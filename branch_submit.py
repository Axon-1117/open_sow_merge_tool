"""Multi-branch SVN submission workflow for the Excel merge tool.

The module deliberately keeps SVN side effects behind ``TortoiseProc`` and
keeps the merge engine pure enough to exercise in a temporary fixture.  The
host application imports it lazily so the legacy TortoiseSVN diff/merge launch
path remains unchanged.
"""

from __future__ import annotations

import copy
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import time
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

from openpyxl import load_workbook


DEFAULT_BRANCHES = ("develop", "release", "sandbox")
SUPPORTED_EXTENSIONS = (".xlsx",)
STATE_VERSION = 1


def _sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _safe_copy(src: str, dst: str) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
    tmp = dst + f".tmp-{os.getpid()}-{uuid.uuid4().hex}"
    shutil.copy2(src, tmp)
    os.replace(tmp, dst)


def _json_default(value):
    if isinstance(value, Path):
        return str(value)
    raise TypeError(f"unsupported JSON value: {type(value).__name__}")


def settings_dir() -> str:
    root = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    return os.path.join(root, "SowMergeTool", "branch_submit")


def load_settings() -> dict:
    path = os.path.join(settings_dir(), "settings.json")
    try:
        with open(path, "r", encoding="utf-8") as stream:
            data = json.load(stream)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(data: dict) -> None:
    root = settings_dir()
    os.makedirs(root, exist_ok=True)
    path = os.path.join(root, "settings.json")
    tmp = path + f".tmp-{os.getpid()}"
    with open(tmp, "w", encoding="utf-8") as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2, default=_json_default)
    os.replace(tmp, path)


def discover_branches(wc_root: str, allowed: Iterable[str] = DEFAULT_BRANCHES) -> list[str]:
    """Return existing, safe, immediate branch directories in deterministic order."""
    root = os.path.abspath(wc_root)
    allowed_set = {str(item).strip() for item in allowed if str(item).strip()}
    if not os.path.isdir(os.path.join(root, ".svn")):
        raise ValueError(f"不是 SVN 工作副本根目录：{root}")
    result = []
    for entry in os.scandir(root):
        if not entry.is_dir() or entry.name.startswith("."):
            continue
        if entry.name not in allowed_set:
            continue
        result.append(entry.name)
    return sorted(result, key=lambda item: (item != "develop", item))


def _validate_branch_name(branch: str, branches: Iterable[str]) -> str:
    value = str(branch or "").strip()
    if value not in set(branches):
        raise ValueError(f"分支不在白名单中：{value or '<empty>'}")
    if value == "master":
        raise ValueError("master 分支不进入多分支自动提交")
    return value


def _validate_relative_file(relative_path: str) -> str:
    value = str(relative_path or "").replace("\\", "/").strip("/")
    if not value or value.startswith("../") or "/../" in f"/{value}/":
        raise ValueError(f"非法配置相对路径：{relative_path!r}")
    if Path(value).suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"v1 只支持 .xlsx：{relative_path}")
    return value


def _relative_source_files(wc_root: str, source_branch: str, selected: Iterable[str]) -> list[str]:
    branch_root = os.path.abspath(os.path.join(wc_root, source_branch))
    result = []
    for path in selected:
        absolute = os.path.abspath(path)
        try:
            relative = os.path.relpath(absolute, branch_root)
        except ValueError as exc:
            raise ValueError(f"文件不在源分支目录：{path}") from exc
        relative = _validate_relative_file(relative)
        if not os.path.isfile(absolute):
            raise ValueError(f"文件不存在：{absolute}")
        result.append(relative)
    if not result:
        raise ValueError("至少选择一个 .xlsx 文件")
    return sorted(dict.fromkeys(result))


def infer_context_from_files(initial_paths: Iterable[str]) -> tuple[str, str, list[str]]:
    """Infer one SVN root and source branch from Explorer-selected workbooks."""
    paths = [os.path.abspath(str(path)) for path in initial_paths if str(path).strip()]
    if not paths:
        raise ValueError("没有可用的右键文件")
    wc_root = ""
    source_branch = ""
    for path in paths:
        if not os.path.isfile(path):
            raise ValueError(f"右键文件不存在：{path}")
        if Path(path).suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError(f"右键入口目前只支持 .xlsx：{path}")
        probe = os.path.dirname(path)
        current_root = ""
        while True:
            if os.path.isfile(os.path.join(probe, ".svn", "wc.db")):
                current_root = probe
                break
            parent = os.path.dirname(probe)
            if parent == probe:
                break
            probe = parent
        if not current_root:
            raise ValueError(f"文件不在 SVN 工作副本中：{path}")
        relative = os.path.relpath(path, current_root).replace("\\", "/")
        parts = relative.split("/")
        if len(parts) < 2:
            raise ValueError(f"无法从路径识别源分支：{path}")
        current_branch = parts[0]
        if not wc_root:
            wc_root, source_branch = current_root, current_branch
        elif os.path.normcase(current_root) != os.path.normcase(wc_root) or current_branch != source_branch:
            raise ValueError("右键选择的文件必须位于同一 SVN 工作副本和同一源分支")
    branches = discover_branches(wc_root)
    _validate_branch_name(source_branch, branches)
    _relative_source_files(wc_root, source_branch, paths)
    return wc_root, source_branch, paths


def _wc_revision(core, path: str) -> int | None:
    root = core._find_svn_wc_root_for_path(path)
    if not root:
        return None
    relative = os.path.relpath(os.path.abspath(path), root).replace("\\", "/")
    revision, _author, _reason = core._wc_node_metadata(root, relative)
    return revision


def _has_conflict(core, path: str) -> bool:
    try:
        return bool(core._detect_svn_conflict_files(path)) or bool(core._has_svn_conflict_artifacts(path))
    except Exception:
        return False


def _openpyxl_signature(path: str) -> tuple[dict[str, tuple[int, int]], dict[str, tuple[int, int]]]:
    """Return conservative sheet dimensions for unsupported-structure gating."""
    wb = load_workbook(path, read_only=False, data_only=False)
    try:
        dims = {}
        merges = {}
        for ws in wb.worksheets:
            dims[ws.title] = (ws.max_row, ws.max_column)
            merges[ws.title] = len(ws.merged_cells.ranges)
        return dims, merges
    finally:
        wb.close()


@dataclass
class FilePlan:
    relative_path: str
    source_before: str = ""
    source_after: str = ""
    source_before_hash: str = ""
    source_after_hash: str = ""
    source_revision: int | None = None
    target_before_hash: dict[str, str] = field(default_factory=dict)
    target_candidate_hash: dict[str, str] = field(default_factory=dict)
    target_candidates: dict[str, str] = field(default_factory=dict)
    target_summaries: dict[str, dict] = field(default_factory=dict)
    target_reasons: dict[str, str] = field(default_factory=dict)


@dataclass
class BranchSubmitBatch:
    batch_id: str
    wc_root: str
    source_branch: str
    target_branches: list[str]
    files: list[FilePlan]
    message: str
    source_status: str = "pending"
    target_status: dict[str, str] = field(default_factory=dict)
    created_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))
    updated_at: str = ""
    source_revision_after: int | None = None
    error: str = ""

    @property
    def folder(self) -> str:
        return os.path.join(settings_dir(), "batches", self.batch_id)

    @property
    def state_path(self) -> str:
        return os.path.join(self.folder, "batch.json")

    def save(self) -> None:
        os.makedirs(self.folder, exist_ok=True)
        self.updated_at = datetime.now().isoformat(timespec="seconds")
        payload = asdict(self)
        payload["state_version"] = STATE_VERSION
        tmp = self.state_path + f".tmp-{os.getpid()}"
        with open(tmp, "w", encoding="utf-8") as stream:
            json.dump(payload, stream, ensure_ascii=False, indent=2, default=_json_default)
        os.replace(tmp, self.state_path)

    @classmethod
    def load(cls, path: str) -> "BranchSubmitBatch":
        with open(path, "r", encoding="utf-8") as stream:
            payload = json.load(stream)
        if payload.get("state_version", STATE_VERSION) != STATE_VERSION:
            raise ValueError("批次状态版本不兼容")
        files = [FilePlan(**item) for item in payload.pop("files", [])]
        payload.pop("state_version", None)
        return cls(files=files, **payload)


class BranchSubmitEngine:
    """Preflight and commit orchestration; UI and subprocesses are injectable."""

    def __init__(self, wc_root: str, *, allowed_branches=DEFAULT_BRANCHES, runner: Callable | None = None):
        self.wc_root = os.path.abspath(wc_root)
        self.allowed_branches = tuple(allowed_branches)
        self.core = None
        self.runner = runner or self._default_runner

    def _load_core(self):
        if self.core is None:
            import sow_merge_tool as core
            self.core = core
        return self.core

    @staticmethod
    def _default_runner(args, *, timeout=300):
        return subprocess.run(args, capture_output=True, text=True, errors="replace", timeout=timeout)

    def _tortoise(self, command: str, paths: list[str], *, message: str | None = None) -> int:
        core = self._load_core()
        exe = core._find_tortoise_proc_exe()
        if not exe or (os.path.dirname(exe) and not os.path.isfile(exe)):
            raise RuntimeError("未找到 TortoiseProc.exe")
        temp_paths = []
        args = [exe, f"/command:{command}"]
        if len(paths) == 1:
            args.append(f"/path:{paths[0]}")
        else:
            pathfile = os.path.join(tempfile.gettempdir(), f"{core.APP_NAME}_pathfile_{uuid.uuid4().hex}.txt")
            with open(pathfile, "w", encoding="utf-16-le", newline="") as stream:
                stream.write("\n".join(paths))
            temp_paths.append(pathfile)
            args.extend([f"/pathfile:{pathfile}", "/deletepathfile"])
        if message is not None:
            logmsg = os.path.join(tempfile.gettempdir(), f"{core.APP_NAME}_logmsg_{uuid.uuid4().hex}.txt")
            with open(logmsg, "w", encoding="utf-8-sig", newline="") as stream:
                stream.write(message)
            temp_paths.append(logmsg)
            args.append(f"/logmsgfile:{logmsg}")
        args.append("/closeonend:1")
        try:
            result = self.runner(args, timeout=3600 if command == "commit" else 600)
            return int(getattr(result, "returncode", 1))
        finally:
            for path in temp_paths:
                try:
                    os.remove(path)
                except OSError:
                    pass

    def _update(self, paths: list[str]) -> None:
        code = self._tortoise("update", paths)
        if code != 0:
            raise RuntimeError(f"SVN update 未成功（退出码 {code}）")

    def _source_plan(self, source_path: str) -> tuple[str, int | None]:
        core = self._load_core()
        if _has_conflict(core, source_path):
            raise RuntimeError(f"源文件存在 SVN 冲突：{source_path}")
        before = core._try_export_svn_base_from_working_copy(source_path)
        if not before:
            raise RuntimeError(f"无法读取源文件 SVN pristine：{source_path}")
        return before, _wc_revision(core, source_path)

    def preflight(self, source_branch: str, target_branches: Iterable[str], selected_files: Iterable[str], message: str) -> BranchSubmitBatch:
        branches = discover_branches(self.wc_root, self.allowed_branches)
        source = _validate_branch_name(source_branch, branches)
        targets = []
        for target in target_branches:
            value = _validate_branch_name(target, branches)
            if value == source:
                raise ValueError("目标分支不能与源分支相同")
            if value not in targets:
                targets.append(value)
        if not targets:
            raise ValueError("至少选择一个目标分支")
        message = str(message or "").strip()
        if not message:
            raise ValueError("SVN 提交说明不能为空")
        relative_files = _relative_source_files(self.wc_root, source, selected_files)
        batch = BranchSubmitBatch(
            batch_id=datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:8],
            wc_root=self.wc_root,
            source_branch=source,
            target_branches=targets,
            files=[],
            message=message,
            target_status={branch: "pending" for branch in targets},
        )
        core = self._load_core()
        for relative in relative_files:
            source_path = os.path.join(self.wc_root, source, *relative.split("/"))
            before, revision = self._source_plan(source_path)
            before_copy = os.path.join(batch.folder, "source-before", relative.replace("/", "__"))
            after_copy = os.path.join(batch.folder, "source-after", relative.replace("/", "__"))
            _safe_copy(before, before_copy)
            _safe_copy(source_path, after_copy)
            plan = FilePlan(
                relative_path=relative,
                source_before=before_copy,
                source_after=after_copy,
                source_before_hash=_sha256(before_copy),
                source_after_hash=_sha256(after_copy),
                source_revision=revision,
            )
            for target in targets:
                target_path = os.path.join(self.wc_root, target, *relative.split("/"))
                if not os.path.isfile(target_path):
                    plan.target_reasons[target] = "目标文件不存在"
                    batch.target_status[target] = "blocked"
                    continue
                if _has_conflict(core, target_path):
                    plan.target_reasons[target] = "目标文件存在 SVN 冲突"
                    batch.target_status[target] = "blocked"
                    continue
                plan.target_before_hash[target] = _sha256(target_path)
                conflicts, candidate, _map, summary, reason = core._cross_branch_source_delta_premerge(
                    before_copy, target_path, after_copy
                )
                plan.target_summaries[target] = dict(summary)
                if reason or conflicts:
                    plan.target_reasons[target] = reason or f"存在 {len(conflicts)} 个待解决冲突"
                    batch.target_status[target] = "blocked"
                    continue
                candidate_copy = os.path.join(batch.folder, "candidates", target, relative.replace("/", "__"))
                _safe_copy(candidate, candidate_copy)
                plan.target_candidates[target] = candidate_copy
                plan.target_candidate_hash[target] = _sha256(candidate_copy)
            batch.files.append(plan)
        blocked = any(value == "blocked" for value in batch.target_status.values())
        for target in targets:
            if batch.target_status.get(target) != "blocked":
                batch.target_status[target] = "ready"
        if blocked:
            batch.error = "预检查未通过：请处理阻断项后重新分析"
        else:
            batch.source_status = "ready"
        batch.save()
        return batch

    def _verify_source_after_commit(self, batch: BranchSubmitBatch) -> None:
        core = self._load_core()
        for plan in batch.files:
            path = os.path.join(batch.wc_root, batch.source_branch, *plan.relative_path.split("/"))
            if not os.path.isfile(path) or _sha256(path) != plan.source_after_hash:
                raise RuntimeError(f"源文件提交后内容与计划不一致：{plan.relative_path}")
            pristine = core._try_export_svn_base_from_working_copy(path)
            if not pristine or _sha256(pristine) != plan.source_after_hash:
                raise RuntimeError(f"未验证源文件 SVN 提交成功：{plan.relative_path}")
        batch.source_revision_after = max((_wc_revision(core, os.path.join(batch.wc_root, batch.source_branch, *p.relative_path.split("/"))) or 0 for p in batch.files), default=0) or None

    def _verify_source_before_commit(self, batch: BranchSubmitBatch) -> None:
        core = self._load_core()
        for plan in batch.files:
            path = os.path.join(batch.wc_root, batch.source_branch, *plan.relative_path.split("/"))
            if _has_conflict(core, path):
                raise RuntimeError(f"源文件提交前存在 SVN 冲突：{plan.relative_path}")
            if not os.path.isfile(path) or _sha256(path) != plan.source_after_hash:
                raise RuntimeError(f"源文件已偏离预览内容，请重新预检查：{plan.relative_path}")

    def commit(self, batch: BranchSubmitBatch, *, stop_on_failure: bool = True) -> BranchSubmitBatch:
        if batch.source_status not in ("ready", "committed"):
            raise RuntimeError(f"批次不可提交：source_status={batch.source_status}")
        core = self._load_core()
        source_paths = [os.path.join(batch.wc_root, batch.source_branch, *p.relative_path.split("/")) for p in batch.files]
        if batch.source_status == "ready":
            try:
                self._verify_source_before_commit(batch)
                code = self._tortoise("commit", source_paths, message=batch.message)
                if code != 0:
                    batch.source_status = "cancelled" if code == 1 else "failed"
                    batch.error = f"源分支提交未完成（退出码 {code}）"
                    batch.save()
                    return batch
                self._verify_source_after_commit(batch)
                batch.source_status = "committed"
                batch.save()
            except Exception as exc:
                batch.source_status = "failed"
                batch.error = str(exc)
                batch.save()
                return batch
        footer = f"[MultiBranchSync] batch={batch.batch_id} source={batch.source_branch}@r{batch.source_revision_after or 'unknown'}"
        target_message = batch.message.rstrip() + "\n\n" + footer
        for target in batch.target_branches:
            if batch.target_status.get(target) in ("committed", "already_present"):
                continue
            try:
                target_paths = [os.path.join(batch.wc_root, target, *p.relative_path.split("/")) for p in batch.files]
                # A cancelled commit leaves the verified candidate in the WC.
                # Resume can safely continue that exact candidate; any other
                # content must be updated and re-projected from a fresh target.
                planned_candidates = {
                    os.path.abspath(os.path.join(batch.wc_root, target, *plan.relative_path.split("/"))): plan.target_candidate_hash.get(target)
                    for plan in batch.files
                    if plan.target_candidate_hash.get(target)
                }
                can_resume_candidate = bool(planned_candidates) and all(
                    os.path.isfile(path) and _sha256(path) == expected
                    for path, expected in planned_candidates.items()
                )
                if not can_resume_candidate:
                    self._update(target_paths)
                commit_paths = []
                for plan in batch.files:
                    target_path = os.path.join(batch.wc_root, target, *plan.relative_path.split("/"))
                    if _has_conflict(core, target_path):
                        raise RuntimeError(f"更新后目标存在 SVN 冲突：{plan.relative_path}")
                    conflicts, candidate, _map, summary, reason = core._cross_branch_source_delta_premerge(
                        plan.source_before, target_path, plan.source_after
                    )
                    if reason or conflicts:
                        raise RuntimeError(reason or f"目标冲突：{plan.relative_path}")
                    plan.target_summaries[target] = dict(summary)
                    if summary.get("applied_count", 0):
                        _safe_copy(candidate, target_path)
                        commit_paths.append(target_path)
                if not commit_paths:
                    batch.target_status[target] = "already_present"
                    batch.save()
                    continue
                code = self._tortoise("commit", commit_paths, message=target_message)
                if code != 0:
                    raise RuntimeError(f"目标提交未完成（退出码 {code}）")
                for path in commit_paths:
                    pristine = core._try_export_svn_base_from_working_copy(path)
                    if not pristine or _sha256(pristine) != _sha256(path):
                        raise RuntimeError(f"未验证目标提交成功：{path}")
                batch.target_status[target] = "committed"
                batch.save()
            except Exception as exc:
                batch.target_status[target] = "cancelled" if "未完成" in str(exc) else "failed"
                batch.error = str(exc)
                batch.save()
                if stop_on_failure:
                    break
        return batch


def _choose_files(root, source_root: str) -> list[str]:
    from tkinter import filedialog
    return list(filedialog.askopenfilenames(
        parent=root,
        title="选择源分支 Excel（仅支持 .xlsx）",
        initialdir=source_root,
        filetypes=[("Excel Workbook", "*.xlsx")],
    ))


def launch_ui(initial_paths: Iterable[str] | None = None) -> None:
    """Launch a compact, user-confirmed branch-submit workflow."""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Excel 合并器 - 多分支 SVN 提交")
    root.geometry("720x560")
    settings = load_settings()
    default_root = settings.get("wc_root") if settings.get("wc_root") else None
    initial_files = [os.path.abspath(str(path)) for path in (initial_paths or []) if str(path).strip()]
    inferred_source = ""
    if initial_files:
        try:
            wc_root, inferred_source, initial_files = infer_context_from_files(initial_files)
        except Exception as exc:
            messagebox.showerror("无法识别右键文件", str(exc), parent=root)
            root.destroy()
            return
    else:
        wc_root = filedialog.askdirectory(parent=root, title="选择 SVN 工作副本根目录", initialdir=default_root or os.getcwd())
        if not wc_root:
            root.destroy()
            return
    try:
        branches = discover_branches(wc_root)
    except Exception as exc:
        messagebox.showerror("工作副本无效", str(exc), parent=root)
        root.destroy()
        return
    if len(branches) < 2:
        messagebox.showerror("分支不足", "白名单中至少需要两个已存在分支。", parent=root)
        root.destroy()
        return
    form = ttk.Frame(root, padding=16)
    form.pack(fill="both", expand=True)
    ttk.Label(form, text="源分支").grid(row=0, column=0, sticky="w")
    source_var = tk.StringVar(value=inferred_source or branches[0])
    source_box = ttk.Combobox(form, textvariable=source_var, values=branches, state="readonly", width=20)
    source_box.grid(row=0, column=1, sticky="w", padx=(8, 0))
    ttk.Label(form, text="目标分支（可多选）").grid(row=1, column=0, sticky="nw", pady=(12, 0))
    target_list = tk.Listbox(form, selectmode=tk.MULTIPLE, exportselection=False, height=5)
    target_list.grid(row=1, column=1, sticky="ew", pady=(12, 0))
    selected_var = tk.StringVar(value="尚未选择文件")
    file_holder = {"paths": list(initial_files)}

    def refresh_targets(*_args):
        target_list.delete(0, tk.END)
        for branch in branches:
            if branch != source_var.get():
                target_list.insert(tk.END, branch)

    def show_selected_files():
        if not file_holder["paths"]:
            selected_var.set("尚未选择文件")
            return
        source_root = os.path.join(wc_root, source_var.get())
        selected_var.set("\n".join(os.path.relpath(path, source_root) for path in file_holder["paths"]))

    def source_changed(_event=None):
        refresh_targets()
        source_root = os.path.abspath(os.path.join(wc_root, source_var.get()))
        def belongs_to_source(path: str) -> bool:
            try:
                return os.path.normcase(os.path.commonpath((source_root, os.path.abspath(path)))) == os.path.normcase(source_root)
            except ValueError:
                return False
        if any(not belongs_to_source(path) for path in file_holder["paths"]):
            file_holder["paths"] = []
        show_selected_files()

    refresh_targets()
    show_selected_files()
    source_box.bind("<<ComboboxSelected>>", source_changed)
    def choose_files():
        picked = _choose_files(root, os.path.join(wc_root, source_var.get()))
        if picked:
            file_holder["paths"] = picked
            show_selected_files()
    ttk.Button(form, text="选择 Excel", command=choose_files).grid(row=2, column=0, sticky="nw", pady=(12, 0))
    ttk.Label(form, textvariable=selected_var, justify="left", wraplength=560).grid(row=2, column=1, sticky="w", pady=(12, 0))
    ttk.Label(form, text="提交说明").grid(row=3, column=0, sticky="nw", pady=(12, 0))
    message = tk.Text(form, height=5, width=60)
    message.grid(row=3, column=1, sticky="ew", pady=(12, 0))
    status = tk.StringVar(value="先选择分支、文件和提交说明，再点击预检查")
    ttk.Label(form, textvariable=status, foreground="#555", wraplength=650).grid(row=4, column=0, columnspan=2, sticky="w", pady=(16, 0))
    form.columnconfigure(1, weight=1)
    def run():
        selected = [target_list.get(index) for index in target_list.curselection()]
        try:
            engine = BranchSubmitEngine(wc_root)
            batch = engine.preflight(source_var.get(), selected, file_holder["paths"], message.get("1.0", tk.END))
            if batch.source_status != "ready":
                raise RuntimeError(batch.error or "预检查未通过")
            lines = [f"批次 {batch.batch_id} 预检查通过。", f"源分支：{batch.source_branch}"]
            lines.extend(f"目标 {branch}：{batch.target_status[branch]}" for branch in batch.target_branches)
            if not messagebox.askyesno("确认提交", "\n".join(lines) + "\n\n将依次打开源分支和目标分支提交窗口，取消会停止后续分支。继续吗？", parent=root):
                return
            save_settings({"wc_root": wc_root, "last_source_branch": source_var.get()})
            result = engine.commit(batch)
            status.set(f"批次完成：source={result.source_status}; " + ", ".join(f"{k}={v}" for k, v in result.target_status.items()))
            messagebox.showinfo("批次结果", status.get(), parent=root)
        except Exception as exc:
            status.set(str(exc))
            messagebox.showerror("多分支提交失败", str(exc), parent=root)
    ttk.Button(form, text="预检查并提交", command=run).grid(row=5, column=1, sticky="e", pady=(18, 0))
    root.mainloop()


def prompt_mode() -> str:
    """Return ``branch`` or ``legacy`` from the no-argument startup chooser."""
    import tkinter as tk
    from tkinter import ttk
    root = tk.Tk()
    root.title("Excel 合并器")
    result = {"mode": "legacy"}
    frame = ttk.Frame(root, padding=20)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text="请选择工作模式", font=("Microsoft YaHei", 12, "bold")).pack(pady=(0, 16))
    def choose(mode):
        result["mode"] = mode
        root.destroy()
    ttk.Button(frame, text="Excel 差异 / 冲突合并", command=lambda: choose("legacy"), width=28).pack(pady=5)
    ttk.Button(frame, text="多分支 SVN 提交", command=lambda: choose("branch"), width=28).pack(pady=5)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()
    return result["mode"]
