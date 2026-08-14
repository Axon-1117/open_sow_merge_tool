"""Fast, conservative OOXML analysis for multi-branch submission.

This module deliberately does not write workbooks.  It indexes the package
parts that are needed to classify a source delta and a target branch.  The
slower Excel/openpyxl writer is only used after the user starts the batch.
"""

from __future__ import annotations

import copy
import os
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from functools import lru_cache

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
M = "{" + MAIN_NS + "}"


@dataclass(frozen=True)
class FastCell:
    value: object
    formula: str = ""
    style: str = ""
    cell_type: str = ""


@dataclass(frozen=True)
class FastRow:
    number: int
    key: str
    cells: dict[int, FastCell]
    style: str = ""


@dataclass
class FastSheet:
    name: str
    member: str
    rows: dict[int, FastRow] = field(default_factory=dict)
    key_rows: dict[str, int] = field(default_factory=dict)
    duplicate_keys: set[str] = field(default_factory=set)
    headers: tuple[str, ...] = ()
    changed_members: set[str] = field(default_factory=set)


@dataclass
class FastWorkbook:
    path: str
    signature: tuple[str, int, int]
    members: dict[str, tuple[int, int]]
    payloads: dict[str, bytes]
    sheets: dict[str, FastSheet]
    shared_strings: tuple[str, ...]


@dataclass
class FastSourceDelta:
    source_path: str
    before_path: str
    changed_sheets: dict[str, dict] = field(default_factory=dict)
    unsupported_reason: str = ""
    incoming_count: int = 0
    changed_count: int = 0


@dataclass
class FastTargetDecision:
    disposition: str
    reason: str = ""
    summary: dict = field(default_factory=dict)
    details: list[dict] = field(default_factory=list)


def _signature(path: str) -> tuple[str, int, int]:
    info = os.stat(path)
    return os.path.normcase(os.path.abspath(path)), int(info.st_size), int(info.st_mtime_ns)


def _text(node: ET.Element | None) -> str:
    if node is None:
        return ""
    return "".join(part.text or "" for part in node.iter(M + "t"))


def _shared_strings(payload: bytes | None) -> tuple[str, ...]:
    if not payload:
        return ()
    root = ET.fromstring(payload)
    return tuple(_text(item) for item in root.findall(M + "si"))


def _cell_value(cell: ET.Element, strings: tuple[str, ...]) -> FastCell:
    formula = (cell.findtext(M + "f") or "").strip()
    value_node = cell.find(M + "v")
    value = value_node.text if value_node is not None else None
    cell_type = cell.attrib.get("t", "")
    if cell_type == "s" and value is not None:
        try:
            value = strings[int(value)]
        except (ValueError, IndexError):
            value = None
    elif cell_type == "inlineStr":
        value = _text(cell.find(M + "is"))
    elif value is not None and cell_type not in {"str", "e"}:
        # Preserve numeric text without an expensive type coercion.  The
        # semantic comparison only needs stable equality for the same OOXML
        # value representation.
        value = value.strip()
    return FastCell(value=value, formula=formula, style=cell.attrib.get("s", ""), cell_type=cell_type)


def _column_number(reference: str) -> int:
    letters = "".join(char for char in reference if char.isalpha())
    result = 0
    for char in letters.upper():
        result = result * 26 + ord(char) - 64
    return result


def _sheet_rows(payload: bytes, strings: tuple[str, ...]) -> dict[int, FastRow]:
    root = ET.fromstring(payload)
    rows: dict[int, FastRow] = {}
    for row in root.findall(".//" + M + "row"):
        try:
            number = int(row.attrib.get("r", "0"))
        except ValueError:
            continue
        cells: dict[int, FastCell] = {}
        for cell in row.findall(M + "c"):
            ref = cell.attrib.get("r", "")
            if not ref:
                continue
            column = _column_number(ref)
            if column:
                cells[column] = _cell_value(cell, strings)
        key_cell = cells.get(1)
        key = str(key_cell.value).strip() if key_cell and key_cell.value not in (None, "") else ""
        rows[number] = FastRow(number, key, cells, row.attrib.get("s", ""))
    return rows


def _sheet_map(payloads: dict[str, bytes]) -> dict[str, str]:
    workbook = ET.fromstring(payloads["xl/workbook.xml"])
    rels = ET.fromstring(payloads["xl/_rels/workbook.xml.rels"])
    targets = {
        item.attrib.get("Id", ""): item.attrib.get("Target", "")
        for item in rels.findall("{" + PACKAGE_REL_NS + "}Relationship")
    }
    result: dict[str, str] = {}
    for sheet in workbook.findall(".//" + M + "sheet"):
        name = sheet.attrib.get("name", "")
        target = targets.get(sheet.attrib.get("{" + REL_NS + "}id", ""), "")
        target = target.lstrip("/")
        target = target if target.startswith("xl/") else os.path.normpath(os.path.join("xl", target)).replace("\\", "/")
        if name and target in payloads:
            result[name] = target
    return result


@lru_cache(maxsize=32)
def _load_cached(signature: tuple[str, int, int]) -> FastWorkbook:
    path = signature[0]
    with zipfile.ZipFile(path, "r") as package:
        names = [name for name in package.namelist() if not name.endswith("/")]
        payloads = {name: package.read(name) for name in names}
    strings = _shared_strings(payloads.get("xl/sharedStrings.xml"))
    sheets: dict[str, FastSheet] = {}
    for name, member in _sheet_map(payloads).items():
        rows = _sheet_rows(payloads[member], strings)
        keys: dict[str, int] = {}
        duplicates: set[str] = set()
        key_counts: dict[str, int] = {}
        for number, row in rows.items():
            if number > 1 and row.key and row.key not in {"string", "唯一值", "type"}:
                key_counts[row.key] = key_counts.get(row.key, 0) + 1
        for number, row in rows.items():
            if number <= 1 or not row.key or row.key in {"string", "唯一值", "type"}:
                continue
            if key_counts.get(row.key, 0) > 1:
                duplicates.add(row.key)
                continue
            if row.key in keys:
                duplicates.add(row.key)
            else:
                keys[row.key] = number
        headers = tuple(
            str(rows.get(1, FastRow(1, "", {})).cells[column].value or "")
            for column in sorted(rows.get(1, FastRow(1, "", {})).cells)
        )
        sheets[name] = FastSheet(name, member, rows, keys, duplicates, headers)
    members = {name: (info.file_size, info.CRC) for name, info in _zip_infos(path)}
    return FastWorkbook(path, signature, members, payloads, sheets, strings)


def _zip_infos(path: str):
    with zipfile.ZipFile(path, "r") as package:
        return [(info.filename, info) for info in package.infolist() if not info.filename.endswith("/")]


def load_workbook_index(path: str) -> FastWorkbook:
    return _load_cached(_signature(path))


def _cell_equal(left: FastCell | None, right: FastCell | None) -> bool:
    if left is None or right is None:
        return left is right
    return (left.value, left.formula) == (right.value, right.formula)


def _schema_compatible(before: FastSheet, after: FastSheet, target: FastSheet | None = None) -> bool:
    if before.headers != after.headers:
        return False
    if target is not None:
        # Ignore formatted blank columns; only compare non-empty header tokens.
        source_headers = {value for value in after.headers if value}
        target_headers = {value for value in target.headers if value}
        if not source_headers.issubset(target_headers):
            return False
    return True


def analyze_source(before_path: str, source_path: str) -> FastSourceDelta:
    before = load_workbook_index(before_path)
    after = load_workbook_index(source_path)
    delta = FastSourceDelta(source_path, before_path)
    if set(before.sheets) != set(after.sheets):
        delta.unsupported_reason = "工作表结构发生变化"
        return delta
    for name in sorted(after.sheets):
        left, right = before.sheets[name], after.sheets[name]
        if left.duplicate_keys or right.duplicate_keys or not _schema_compatible(left, right):
            delta.unsupported_reason = f"{name} 的唯一键或字段结构无法证明稳定"
            return delta
        added = sorted(set(right.key_rows) - set(left.key_rows))
        deleted = sorted(set(left.key_rows) - set(right.key_rows))
        changed: list[dict] = []
        for key in sorted(set(left.key_rows) & set(right.key_rows)):
            old = left.rows[left.key_rows[key]]
            new = right.rows[right.key_rows[key]]
            # A direct cell projection deliberately preserves target formatting.
            # If the source changed row/cell styles at the same time as values,
            # copying only the value would silently lose the designer's intent.
            if old.style != new.style:
                delta.unsupported_reason = f"{name} 的记录行样式发生变化，需要人工合并"
                return delta
            columns = sorted(set(old.cells) | set(new.cells))
            for column in columns:
                old_cell, new_cell = old.cells.get(column), new.cells.get(column)
                if old_cell is not None and new_cell is not None and old_cell.style != new_cell.style:
                    delta.unsupported_reason = f"{name} 的单元格样式发生变化，需要人工合并"
                    return delta
                if not _cell_equal(old.cells.get(column), new.cells.get(column)):
                    changed.append({"key": key, "before_row": old.number, "after_row": new.number, "column": column})
        if added or deleted or changed:
            delta.changed_sheets[name] = {"added": added, "deleted": deleted, "changed": changed}
            delta.incoming_count += len(added) + len(deleted) + len(changed)
        elif before.payloads.get(left.member) != after.payloads.get(right.member):
            # A changed sheet with no provable record key cannot be treated as
            # an empty delta.  It is a manual merge case, not "already applied".
            delta.unsupported_reason = f"{name} 发生变化但没有可证明的唯一键"
            return delta
    delta.changed_count = delta.incoming_count
    return delta


def analyze_target(delta: FastSourceDelta, target_path: str) -> FastTargetDecision:
    if delta.unsupported_reason:
        return FastTargetDecision("manual", delta.unsupported_reason)
    before = load_workbook_index(delta.before_path)
    source = load_workbook_index(delta.source_path)
    # Classification is read-only; do not deep-copy a large workbook here.
    # The writer takes its own copy only when it needs to re-index a candidate.
    target = load_workbook_index(target_path)
    details: list[dict] = []
    for sheet_name, changes in delta.changed_sheets.items():
        source_sheet = source.sheets[sheet_name]
        before_sheet = before.sheets[sheet_name]
        target_sheet = target.sheets.get(sheet_name)
        if target_sheet is None or not _schema_compatible(before_sheet, source_sheet, target_sheet):
            return FastTargetDecision("manual", f"目标缺少可映射的工作表或字段：{sheet_name}")
        for key in changes["added"]:
            if key in target_sheet.key_rows:
                details.append({"sheet": sheet_name, "key": key, "kind": "manual", "reason": "目标已存在同一唯一键"})
                continue
            row = source_sheet.rows[source_sheet.key_rows[key]]
            if row.number < max(source_sheet.rows):
                details.append({"sheet": sheet_name, "key": key, "kind": "manual", "reason": "新增记录不在源表尾部"})
            elif row.style and row.style != "0" or any(cell.style and cell.style != "0" for cell in row.cells.values()):
                details.append({"sheet": sheet_name, "key": key, "kind": "manual", "reason": "新增记录带有非默认样式"})
            else:
                details.append({"sheet": sheet_name, "key": key, "kind": "add", "row": row.number})
        for key in changes["deleted"]:
            target_row = target_sheet.key_rows.get(key)
            if target_row is None:
                details.append({"sheet": sheet_name, "key": key, "kind": "already"})
            elif not _row_equal(target_sheet.rows[target_row], before_sheet.rows[before_sheet.key_rows[key]]):
                details.append({"sheet": sheet_name, "key": key, "kind": "manual", "reason": "目标记录有独立修改"})
            else:
                details.append({"sheet": sheet_name, "key": key, "kind": "delete", "row": target_row})
        for change in changes["changed"]:
            key = change["key"]
            target_row = target_sheet.key_rows.get(key)
            if target_row is None:
                details.append({"sheet": sheet_name, "key": key, "kind": "manual", "reason": "目标不存在对应记录"})
                continue
            old = before_sheet.rows[change["before_row"]].cells.get(change["column"])
            new = source_sheet.rows[change["after_row"]].cells.get(change["column"])
            mine = target_sheet.rows[target_row].cells.get(change["column"])
            if _cell_equal(mine, new):
                kind = "already"
            elif _cell_equal(mine, old):
                kind = "change"
            else:
                kind = "manual"
            details.append({"sheet": sheet_name, "key": key, "kind": kind, "column": change["column"]})
    manual = [item for item in details if item["kind"] == "manual"]
    auto = [item for item in details if item["kind"] in {"add", "delete", "change"}]
    already = [item for item in details if item["kind"] == "already"]
    if manual:
        return FastTargetDecision("manual", manual[0].get("reason", "需要人工合并"), {"auto": len(auto), "already": len(already), "manual": len(manual)}, details)
    if not auto:
        return FastTargetDecision("already_applied", "目标已经包含源分支变更", {"already": len(already)}, details)
    return FastTargetDecision("one_click", f"可一键处理 {len(auto)} 项", {"auto": len(auto), "already": len(already)}, details)


def _row_equal(left: FastRow, right: FastRow) -> bool:
    columns = set(left.cells) | set(right.cells)
    return all(_cell_equal(left.cells.get(column), right.cells.get(column)) for column in columns)


def cache_clear() -> None:
    _load_cached.cache_clear()


def apply_one_click_plan(
    before_path: str,
    source_path: str,
    target_path: str,
    output_path: str,
    decision: FastTargetDecision,
) -> None:
    """Materialize a previously classified safe plan into a candidate.

    This is intentionally called only after the source commit has been
    reconciled and the target has passed a fresh SVN status check.  The fast
    analyzer remains read-only; this function is the first workbook write.
    """
    if decision.disposition != "one_click":
        raise ValueError(f"不能物化 {decision.disposition} 动作")
    source = load_workbook_index(source_path)
    target = copy.deepcopy(load_workbook_index(target_path))
    changed_payloads: dict[str, bytes] = {}
    try:
        for detail in decision.details:
            kind = detail.get("kind")
            if kind not in {"add", "delete", "change"}:
                continue
            sheet = str(detail["sheet"])
            source_ws = source.sheets[sheet]
            target_ws = target.sheets[sheet]
            key = str(detail["key"])
            source_row = source_ws.key_rows.get(key)
            target_row = target_ws.key_rows.get(key)
            root = ET.fromstring(changed_payloads.get(target_ws.member, target.payloads[target_ws.member]))
            sheet_data = root.find(M + "sheetData")
            if sheet_data is None:
                raise ValueError(f"目标工作表缺少 sheetData：{sheet}")
            if kind == "add":
                if target_row is not None:
                    continue
                target_row = max(target_ws.rows or {1}) + 1
                _append_row_xml(sheet_data, source_ws, source_row, target_ws, target_row)
            elif kind == "delete":
                if target_row is not None:
                    _remove_row_xml(sheet_data, target_row)
            elif kind == "change" and target_row is not None:
                source_column = int(detail["column"])
                target_column = _target_column_for_source(source_ws, target_ws, source_column)
                if target_column:
                    source_cell = source_ws.rows[source_row].cells.get(source_column)
                    if source_cell is not None:
                        _set_cell_xml(sheet_data, target_row, target_column, source_cell)
            _update_dimension(root, sheet_data)
            changed_payloads[target_ws.member] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            # Re-index the logical target rows after a structural operation so
            # subsequent changes in the same Sheet use current row numbers.
            target = _reindex_changed_target(target, target_ws.member, changed_payloads[target_ws.member])
        _safe_rewrite_package(target_path, output_path, changed_payloads)
    finally:
        pass


def _target_column_for_source(source_ws: FastSheet, target_ws: FastSheet, source_column: int) -> int | None:
    source_header = source_ws.headers[source_column - 1] if source_column <= len(source_ws.headers) else ""
    if source_header not in (None, ""):
        for column, header in enumerate(target_ws.headers, 1):
            if header == source_header:
                return column
    return source_column if source_column <= len(target_ws.headers) else None


def _append_row_xml(sheet_data: ET.Element, source_ws: FastSheet, source_row: int | None, target_ws: FastSheet, target_row: int) -> None:
    if source_row is None:
        raise ValueError("源新增记录缺少唯一键行")
    row = ET.Element(M + "row", {"r": str(target_row)})
    for column, source_cell in source_ws.rows[source_row].cells.items():
        target_column = _target_column_for_source(source_ws, target_ws, column)
        if not target_column:
            continue
        cell = ET.SubElement(row, M + "c", {"r": f"{_column_name(target_column)}{target_row}"})
        if source_cell.style and source_cell.style != "0":
            cell.attrib["s"] = source_cell.style
        _write_cell_payload(cell, source_cell)
    sheet_data.append(row)


def _remove_row_xml(sheet_data: ET.Element, row_number: int) -> None:
    for row in list(sheet_data.findall(M + "row")):
        if row.attrib.get("r") == str(row_number):
            sheet_data.remove(row)
            return


def _update_dimension(root: ET.Element, sheet_data: ET.Element) -> None:
    max_row = 1
    max_column = 1
    for row in sheet_data.findall(M + "row"):
        try:
            max_row = max(max_row, int(row.attrib.get("r", "1")))
        except ValueError:
            pass
        for cell in row.findall(M + "c"):
            max_column = max(max_column, _column_number(cell.attrib.get("r", "A1")))
    dimension = root.find(M + "dimension")
    if dimension is not None:
        dimension.attrib["ref"] = f"A1:{_column_name(max_column)}{max_row}"


def _set_cell_xml(sheet_data: ET.Element, row_number: int, column: int, source_cell: FastCell) -> None:
    row = next((item for item in sheet_data.findall(M + "row") if item.attrib.get("r") == str(row_number)), None)
    if row is None:
        row = ET.Element(M + "row", {"r": str(row_number)})
        sheet_data.append(row)
    reference = f"{_column_name(column)}{row_number}"
    cell = next((item for item in row.findall(M + "c") if item.attrib.get("r") == reference), None)
    if cell is None:
        cell = ET.SubElement(row, M + "c", {"r": reference})
    _write_cell_payload(cell, source_cell)


def _write_cell_payload(cell: ET.Element, source_cell: FastCell) -> None:
    for child in list(cell):
        if child.tag in {M + "f", M + "v", M + "is"}:
            cell.remove(child)
    if source_cell.formula:
        cell.attrib.pop("t", None)
        ET.SubElement(cell, M + "f").text = source_cell.formula
        return
    value = source_cell.value
    if value is None or value == "":
        cell.attrib.pop("t", None)
        return
    if source_cell.cell_type in {"s", "inlineStr", "str"} or isinstance(value, str) and not _looks_numeric(value):
        cell.attrib["t"] = "inlineStr"
        inline = ET.SubElement(cell, M + "is")
        ET.SubElement(inline, M + "t").text = str(value)
    else:
        cell.attrib.pop("t", None)
        ET.SubElement(cell, M + "v").text = str(value)


def _looks_numeric(value: str) -> bool:
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def _column_name(column: int) -> str:
    value = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        value = chr(65 + remainder) + value
    return value


def _reindex_changed_target(target: FastWorkbook, member: str, payload: bytes) -> FastWorkbook:
    strings = target.shared_strings
    rows = _sheet_rows(payload, strings)
    sheet = next(item for item in target.sheets.values() if item.member == member)
    keys: dict[str, int] = {}
    for number, row in rows.items():
        if number > 1 and row.key and row.key not in {"string", "唯一值", "type"} and row.key not in keys:
            keys[row.key] = number
    target.sheets[sheet.name] = FastSheet(sheet.name, member, rows, keys, set(), sheet.headers)
    target.payloads[member] = payload
    return target


def _safe_rewrite_package(source_path: str, output_path: str, changed_payloads: dict[str, bytes]) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    temporary = output_path + ".tmp-fast-merge"
    with zipfile.ZipFile(source_path, "r") as source, zipfile.ZipFile(temporary, "w") as destination:
        for info in source.infolist():
            payload = changed_payloads.get(info.filename, source.read(info.filename))
            destination.writestr(info, payload)
    os.replace(temporary, output_path)
