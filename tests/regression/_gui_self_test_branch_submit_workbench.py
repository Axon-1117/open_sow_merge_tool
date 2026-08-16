"""Headed smoke test for the dense multi-branch submit workbench."""

from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import time
from unittest.mock import patch

import tkinter as tk

from sow_merge_tool import branch_submit as bs


def _create_fixture():
    root = tempfile.mkdtemp(prefix="branch-submit-gui-")
    os.makedirs(os.path.join(root, ".svn"))
    with sqlite3.connect(os.path.join(root, ".svn", "wc.db")) as conn:
        conn.executescript(
            """
            create table REPOSITORY (id integer primary key, root text, uuid text);
            create table NODES (
              wc_id integer, local_relpath text, op_depth integer, parent_relpath text,
              repos_id integer, repos_path text, revision integer, presence text,
              moved_here integer, moved_to text, kind text, properties blob, depth text,
              checksum text, symlink_target text, changed_revision integer,
              changed_date integer, changed_author text, translated_size integer,
              last_mod_time integer, dav_cache blob, file_external integer,
              inherited_props blob
            );
            create table ACTUAL_NODE (
              wc_id integer, local_relpath text, parent_relpath text, properties blob,
              conflict_old text, conflict_new text, conflict_working text,
              prop_reject text, changelist text, text_mod text,
              tree_conflict_data text, conflict_data blob, older_checksum text,
              left_checksum text, right_checksum text
            );
            insert into REPOSITORY values (1, 'file:///repo', 'gui-fixture');
            """
        )
        for index in range(32):
            name = "develop" if index == 0 else "master" if index == 1 else f"feature_{index:02d}"
            folder = os.path.join(root, name)
            os.makedirs(folder)
            open(os.path.join(folder, "seed.xlsx"), "wb").close()
            conn.execute(
                """insert into NODES
                (wc_id,local_relpath,op_depth,parent_relpath,repos_id,repos_path,revision,presence,moved_here,moved_to,kind,changed_revision,changed_author,file_external)
                values (1,?,0,'',1,?,1,'normal',0,'','dir',1,'tester',0)""",
                (name, "sheets/" + name),
            )
    return root


def main():
    fixture = _create_fixture()
    root = tk.Tk()
    root.withdraw()
    items = [
        bs.SvnChangeItem(
            path=os.path.join(fixture, "develop", f"配置_{index:03d}.xlsx"),
            relative_path=f"配置_{index:03d}.xlsx", extension=".xlsx", node_kind="file",
            node_status="modified" if index % 4 else "unversioned",
            text_status="modified", prop_status="normal", versioned=index % 4 != 0,
            checked=index % 4 != 0, selectable=True,
        )
        for index in range(200)
    ]
    try:
        context = bs.BranchContext(fixture, "develop", os.path.join(fixture, "develop"))
        with patch.object(bs, "scan_changes", lambda *_args, **_kwargs: items):
            app = bs.BranchSubmitWorkbench(root, context)
            root.deiconify()
            deadline = time.time() + 5
            while len(app.items) != 200 and time.time() < deadline:
                root.update(); time.sleep(0.02)
            root.update_idletasks()
            assert len(app.items) == 200
            assert len(app.target_vars) == 31
            assert "master" in app.target_vars and not app.target_vars["master"].get()
            assert len(app.tree.get_children()) == 200
            assert int(root.winfo_width()) >= 900 and int(root.winfo_height()) >= 620
            assert set(app.tree["columns"]) == {"check", "path", "extension", "status", "property", "lock", "switched", "changelist"}
            assert "预检查（必需）" in app.preflight_button.cget("text")
            assert app.submit_button.instate(["disabled"]), "multi-branch submit must be preflight-gated"
            target_name = next(name for name in app.target_vars if name != "master")
            app.target_vars[target_name].set(True)
            app._target_selection[target_name] = True
            app.message.insert("1.0", "GUI 门禁测试")
            app._refresh_primary_button()
            assert app.preflight_button.instate(["!disabled"])
            assert app.submit_button.instate(["disabled"]), "selection alone must not enable submit"
            for item in app.items:
                item.checked = item.relative_path == "配置_001.xlsx"
            action = bs.BatchFileAction(
                branch=target_name,
                relative_path="配置_001.xlsx",
                operation="modify",
                state="confirmation_required",
                reason="目标分支同一单元格已有独立修改",
            )
            plan = bs.FilePlan(
                relative_path="配置_001.xlsx",
                operation="modify",
                actions={target_name: action},
                target_summaries={target_name: {"confirmation": 1}},
                target_details={target_name: [{
                    "kind": "confirmation", "sheet": "Data", "key": "1001",
                    "field": "value", "before": "旧", "source": "新", "target": "目标值",
                    "reason": action.reason,
                }]},
            )
            app.current_batch = bs.BranchSubmitBatch(
                batch_id="gui-manual",
                wc_root=fixture,
                source_branch="develop",
                target_branches=[target_name],
                files=[plan],
                message="GUI 门禁测试",
                scope_path=os.path.join(fixture, "develop"),
                source_status="ready",
                target_status={target_name: "confirmation_required"},
            )
            app._approved_preflight_signature = app._request_signature()
            app._render_target_statuses()
            root.update_idletasks()
            assert app.confirmation_alert.winfo_manager() == "pack"
            assert "内容重叠" in app.confirmation_alert_var.get()
            assert app.submit_button.instate(["disabled"]), "confirmation items must keep submit gated"
            app._open_confirmation_dialog()
            root.update_idletasks()
            dialog_tree = app._confirmation_dialog_tree
            assert dialog_tree is not None
            assert tuple(dialog_tree["columns"]) == ("branch", "file", "state", "reason")
            dialog_rows = dialog_tree.get_children()
            assert len(dialog_rows) == 1 and dialog_tree.set(dialog_rows[0], "state") == "待确认"
            app._set_confirmation_dialog_row(target_name, "配置_001.xlsx", "completed", "已确认采用源修改")
            root.update_idletasks()
            assert len(dialog_tree.get_children()) == 1, "completed row must remain in the current dialog"
            assert dialog_tree.set(dialog_rows[0], "state") == "已确认"
            app._confirmation_dialog.grab_release()
            app._confirmation_dialog.destroy()
            app._confirmation_dialog = None
            app._confirmation_dialog_tree = None
            app._confirmation_dialog_button = None
            app._confirmation_exclude_button = None
            app._confirmation_detail = None
            app._confirmation_dialog_summary_var = None
            app._confirmation_dialog_rows = {}
            action.state = "ready"
            action.confirmed = True
            app.current_batch.target_status[target_name] = "ready"
            app._render_target_statuses()
            app._refresh_primary_button()
            root.update_idletasks()
            assert not app.confirmation_alert.winfo_manager()
            assert app.submit_button.instate(["!disabled"]), "confirmed items may pass the gate"
            hold = float(os.environ.get("SOW_GUI_TEST_HOLD", "0") or 0)
            deadline = time.time() + hold
            while time.time() < deadline:
                root.update(); time.sleep(0.03)
        print("PASS: branch-submit workbench 32 branches / 200 files")
    finally:
        try: root.destroy()
        except tk.TclError: pass
        shutil.rmtree(fixture, ignore_errors=True)


if __name__ == "__main__":
    main()
