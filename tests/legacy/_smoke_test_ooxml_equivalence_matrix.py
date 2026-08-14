"""OpenSpec 4.2 regressions for complete OOXML package equivalence.

The fixtures are intentionally tiny ZIP packages.  This keeps the tests fast
while proving that comparison is based on the complete member set and each
member's uncompressed bytes, not ZIP ordering/timestamps or an openpyxl subset.
"""

from __future__ import annotations

import inspect
import os
import zipfile
from itertools import combinations

import sow_merge_tool as smt
from _test_temp_utils import make_temp_dir


_CORE_MEMBERS = {
    "[Content_Types].xml": (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Default Extension="xml" ContentType="application/xml"/>'
        b"</Types>"
    ),
    "_rels/.rels": (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    ),
    "xl/workbook.xml": b"<workbook><sheets><sheet name=\"Data\"/></sheets></workbook>",
    "xl/worksheets/sheet1.xml": b"<worksheet><sheetData/></worksheet>",
}


def _write_package(
    path: str,
    members: dict[str, bytes],
    *,
    reverse: bool = False,
    timestamp=(2020, 1, 2, 3, 4, 6),
) -> None:
    names = sorted(members, reverse=reverse)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        for name in names:
            info = zipfile.ZipInfo(name, date_time=timestamp)
            info.compress_type = zipfile.ZIP_DEFLATED
            package.writestr(info, members[name])


def _field(value, *names, default=None):
    if isinstance(value, dict):
        for name in names:
            if name in value:
                return value[name]
        return default
    for name in names:
        if hasattr(value, name):
            return getattr(value, name)
    return default


def _is_ready(result) -> bool:
    explicit = _field(result, "ready", "complete", "comparison_complete")
    if explicit is not None:
        return bool(explicit)
    status = str(_field(result, "status", "state", default="")).lower()
    return status in {"ready", "complete", "ok", "equal", "different"}


def _is_package_equal(result):
    value = _field(
        result,
        "package_equal",
        "package_equivalent",
        "equivalent",
        "equal",
    )
    return None if value is None else bool(value)


def _raw_hash_equal(result):
    value = _field(result, "raw_sha256_equal", "raw_hash_equal", "sha256_equal")
    return None if value is None else bool(value)


def _reason(result) -> str:
    return str(_field(result, "reason", "error", "message", default="") or "")


def _compare(left: str, right: str):
    comparator = getattr(smt, "compare_ooxml_packages", None)
    assert callable(comparator), "missing compare_ooxml_packages API"
    return comparator(left, right)


def test_zip_metadata_and_entry_order_are_ignored() -> None:
    root = make_temp_dir("sow_ooxml_zip_metadata_")
    left = os.path.join(root, "left.xlsx")
    right = os.path.join(root, "right.xlsx")
    _write_package(left, dict(_CORE_MEMBERS), timestamp=(2020, 1, 2, 3, 4, 6))
    _write_package(
        right,
        dict(_CORE_MEMBERS),
        reverse=True,
        timestamp=(2026, 7, 28, 12, 30, 0),
    )

    result = _compare(left, right)
    assert _is_ready(result), result
    assert _is_package_equal(result) is True, result
    raw_equal = _raw_hash_equal(result)
    if raw_equal is not None:
        assert raw_equal is False, "fixture must differ as raw ZIP bytes"


def test_member_set_and_member_bytes_are_complete_evidence() -> None:
    root = make_temp_dir("sow_ooxml_member_difference_")
    baseline = os.path.join(root, "baseline.xlsx")
    changed_part = os.path.join(root, "changed-part.xlsx")
    missing_part = os.path.join(root, "missing-part.xlsx")
    _write_package(baseline, dict(_CORE_MEMBERS))

    changed = dict(_CORE_MEMBERS)
    changed["xl/worksheets/sheet1.xml"] = (
        b"<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>9</v></c></row></sheetData></worksheet>"
    )
    _write_package(changed_part, changed)

    missing = dict(_CORE_MEMBERS)
    del missing["xl/worksheets/sheet1.xml"]
    _write_package(missing_part, missing)

    for right in (changed_part, missing_part):
        result = _compare(baseline, right)
        assert _is_ready(result), result
        assert _is_package_equal(result) is False, (right, result)
        assert _reason(result), "different packages need an auditable reason"


def test_xlsm_vba_payload_is_compared() -> None:
    root = make_temp_dir("sow_ooxml_vba_difference_")
    left = os.path.join(root, "left.xlsm")
    right = os.path.join(root, "right.xlsm")
    left_members = dict(_CORE_MEMBERS)
    right_members = dict(_CORE_MEMBERS)
    left_members["xl/vbaProject.bin"] = b"VBA-PROJECT-LEFT\x00\x01"
    right_members["xl/vbaProject.bin"] = b"VBA-PROJECT-RIGHT\x00\x01"
    _write_package(left, left_members)
    _write_package(right, right_members, reverse=True)

    result = _compare(left, right)
    assert _is_ready(result), result
    assert _is_package_equal(result) is False, result
    assert "vba" in _reason(result).lower() or _reason(result), result


def test_unreadable_and_missing_packages_fail_closed() -> None:
    root = make_temp_dir("sow_ooxml_error_")
    valid = os.path.join(root, "valid.xlsx")
    corrupt = os.path.join(root, "corrupt.xlsx")
    missing = os.path.join(root, "missing.xlsx")
    _write_package(valid, dict(_CORE_MEMBERS))
    with open(corrupt, "wb") as handle:
        handle.write(b"not-an-ooxml-package")

    for right in (corrupt, missing):
        result = _compare(valid, right)
        assert not _is_ready(result), (right, result)
        assert _is_package_equal(result) is not True, (right, result)
        assert _reason(result), "error result must explain why comparison is unavailable"


def _call_context_builder(base, mine, theirs, merged, pristine):
    builder = getattr(smt, "build_merge_launch_context", None)
    assert callable(builder), "missing build_merge_launch_context API"
    signature = inspect.signature(builder)
    candidates = {
        "base_path": base,
        "source_base_path": base,
        "mine_path": mine,
        "theirs_path": theirs,
        "merged_path": merged,
        "target_pristine_path": pristine,
    }
    kwargs = {
        name: value
        for name, value in candidates.items()
        if name in signature.parameters
    }
    try:
        return builder(**kwargs)
    except TypeError:
        try:
            return builder(base, mine, theirs, merged, target_pristine_path=pristine)
        except TypeError:
            return builder(base, mine, theirs, merged, pristine)


def _identity_map(context) -> dict[str, object]:
    result = {}
    aliases = {
        "base": ("source_base", "source_base_identity", "base", "base_identity"),
        "mine": ("mine", "mine_identity", "mine_working"),
        "theirs": ("theirs", "theirs_identity", "theirs_incoming"),
        "target_pristine": (
            "target_pristine",
            "target_pristine_identity",
            "target_wc_pristine",
        ),
    }
    identities = _field(context, "identities", "versions")
    if isinstance(identities, dict):
        for role in aliases:
            if role in identities:
                result[role] = identities[role]
    for role, names in aliases.items():
        if role in result:
            continue
        identity = _field(context, *names)
        if identity is not None:
            result[role] = identity
    assert set(result) == set(aliases), (
        "launch context must expose all four independent identities",
        result,
    )
    return result


def _matrix_result(matrix, left_role, right_role, left_identity, right_identity):
    getter = getattr(matrix, "get", None)
    if callable(getter):
        attempts = (
            (left_role, right_role),
            (left_identity, right_identity),
            ((left_role, right_role),),
            (frozenset((left_role, right_role)),),
        )
        for args in attempts:
            try:
                found = getter(*args)
            except (KeyError, TypeError, AttributeError):
                continue
            if found is not None:
                return found

    comparisons = _field(matrix, "comparisons", "pairs", "results")
    if comparisons is None and isinstance(matrix, dict):
        comparisons = matrix
    if isinstance(comparisons, dict):
        keys = (
            (left_role, right_role),
            (right_role, left_role),
            frozenset((left_role, right_role)),
            f"{left_role}:{right_role}",
            f"{right_role}:{left_role}",
            f"{left_role}-{right_role}",
            f"{right_role}-{left_role}",
        )
        for key in keys:
            if key in comparisons:
                return comparisons[key]
        for key, value in comparisons.items():
            text = str(key).lower()
            if left_role in text and right_role in text:
                return value
    raise AssertionError(f"matrix pair missing: {left_role}/{right_role}: {matrix!r}")


def test_complete_four_identity_matrix_and_clean_target_semantics() -> None:
    root = make_temp_dir("sow_ooxml_matrix_")
    base = os.path.join(root, "Design.xlsx.r10")
    mine = os.path.join(root, "Design.xlsx")
    theirs = os.path.join(root, "Design.xlsx.r12")
    pristine = os.path.join(root, "Design.target-pristine.xlsx")

    base_members = dict(_CORE_MEMBERS)
    mine_members = dict(_CORE_MEMBERS)
    mine_members["xl/worksheets/sheet1.xml"] = b"<worksheet><sheetData><row r=\"2\"/></sheetData></worksheet>"
    theirs_members = dict(_CORE_MEMBERS)
    theirs_members["xl/worksheets/sheet1.xml"] = b"<worksheet><sheetData><row r=\"3\"/></sheetData></worksheet>"
    _write_package(base, base_members)
    _write_package(mine, mine_members, timestamp=(2021, 1, 1, 0, 0, 0))
    _write_package(
        pristine,
        mine_members,
        reverse=True,
        timestamp=(2026, 7, 28, 10, 0, 0),
    )
    _write_package(theirs, theirs_members)

    context = _call_context_builder(base, mine, theirs, mine, pristine)
    identities = _identity_map(context)
    builder = getattr(smt, "build_equivalence_matrix", None)
    assert callable(builder), "missing build_equivalence_matrix API"
    try:
        matrix = builder(identities)
    except TypeError:
        matrix = builder(tuple(identities.values()))

    observed = {}
    for left_role, right_role in combinations(identities, 2):
        comparison = _matrix_result(
            matrix,
            left_role,
            right_role,
            identities[left_role],
            identities[right_role],
        )
        assert _is_ready(comparison), (left_role, right_role, comparison)
        observed[frozenset((left_role, right_role))] = _is_package_equal(comparison)

    assert len(observed) == 6, observed
    assert observed[frozenset(("mine", "target_pristine"))] is True
    assert observed[frozenset(("base", "mine"))] is False
    assert observed[frozenset(("base", "target_pristine"))] is False


def test_unavailable_pristine_still_has_complete_fail_closed_matrix() -> None:
    root = make_temp_dir("sow_ooxml_matrix_missing_pristine_")
    base = os.path.join(root, "Design.xlsx.r10")
    mine = os.path.join(root, "Design.xlsx")
    theirs = os.path.join(root, "Design.xlsx.r12")
    for path in (base, mine, theirs):
        _write_package(path, dict(_CORE_MEMBERS))

    context = _call_context_builder(base, mine, theirs, mine, None)
    identities = _identity_map(context)
    matrix = smt.build_equivalence_matrix(context)
    comparisons = _field(matrix, "comparisons", "pairs", "results")
    assert len(comparisons) == 6, comparisons
    for other_role in ("base", "mine", "theirs"):
        comparison = _matrix_result(
            matrix,
            other_role,
            "target_pristine",
            identities[other_role],
            identities["target_pristine"],
        )
        assert not _is_ready(comparison), comparison
        assert "missing input path" in str(_field(comparison, "reason", default="")), comparison


def main() -> None:
    tests = (
        test_zip_metadata_and_entry_order_are_ignored,
        test_member_set_and_member_bytes_are_complete_evidence,
        test_xlsm_vba_payload_is_compared,
        test_unreadable_and_missing_packages_fail_closed,
        test_complete_four_identity_matrix_and_clean_target_semantics,
        test_unavailable_pristine_still_has_complete_fail_closed_matrix,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}")
    print(f"PASS: OOXML equivalence matrix ({len(tests)} tests)")


if __name__ == "__main__":
    main()
