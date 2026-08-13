"""Headed smoke test for the dense multi-branch submit workbench."""

from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import time
from unittest.mock import patch

import tkinter as tk

import branch_submit as bs


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
