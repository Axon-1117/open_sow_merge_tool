"""OpenSpec 4.5 regressions for SVN author labels and audit logging."""

from __future__ import annotations

import os
import shutil
import sqlite3
import subprocess
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

import sow_merge_tool as smt
from _test_temp_utils import make_temp_dir


def _make_book(path: str, marker: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "marker"
    worksheet["A2"] = marker
    workbook.save(path)
    workbook.close()


def _create_local_wc_fixture():
    root = make_temp_dir("sow_svn_author_wc_")
    os.makedirs(os.path.join(root, ".svn"), exist_ok=True)
    db_path = os.path.join(root, ".svn", "wc.db")
    with sqlite3.connect(db_path) as connection:
        connection.execute(
            """
            create table REPOSITORY (
                id integer primary key,
                root text,
                uuid text
            )
            """
        )
        connection.execute(
            "insert into REPOSITORY (id, root, uuid) values (1, ?, 'fixture-repo')",
            ("http://svn.example.test/repository",),
        )
        connection.execute(
            """
            create table NODES (
                local_relpath text,
                op_depth integer,
                kind text,
                presence text,
                repos_id integer,
                changed_revision integer,
                changed_author text,
                checksum text
            )
            """
        )
        connection.executemany(
            """
            insert into NODES (
                local_relpath, op_depth, kind, presence,
                repos_id, changed_revision, changed_author, checksum
            ) values (?, 0, 'file', 'normal', 1, ?, ?, null)
            """,
            (
                ("source-left/Design.xlsx", 200, "alice"),
                ("target/Design.xlsx", 205, "carol"),
                ("source-right/Design.xlsx", 210, "bob"),
            ),
        )
        connection.commit()

    target_dir = os.path.join(root, "target")
    os.makedirs(target_dir, exist_ok=True)
    base = os.path.join(target_dir, "Design.xlsx.merge-left.r200")
    mine = os.path.join(target_dir, "Design.xlsx")
    theirs = os.path.join(target_dir, "Design.xlsx.merge-right.r210")
    pristine = os.path.join(target_dir, "Design.target-pristine.xlsx")
    _make_book(base, "base")
    _make_book(mine, "mine-local")
    _make_book(theirs, "theirs")
    _make_book(pristine, "mine-pristine")
    context = smt.build_merge_launch_context(
        base,
        mine,
        theirs,
        mine,
        target_pristine_path=pristine,
    )
    return root, context


def test_sidecar_suffix_is_stripped_for_exact_author_lookup() -> None:
    _root, context = _create_local_wc_fixture()
    logs = []
    with patch.object(smt, "_dlog", lambda message: logs.append(str(message))):
        identities = smt.resolve_svn_author_metadata(context)

    base = identities["base"]
    mine = identities["mine"]
    theirs = identities["theirs"]
    pristine = identities["target_pristine"]
    assert base.revision == 200
    assert base.author == "alice"
    assert base.author_status == "resolved"
    assert "source-left/Design.xlsx@r200" in base.author_source.replace("\\", "/")
    assert theirs.revision == 210
    assert theirs.author == "bob"
    assert theirs.author_status == "resolved"
    assert "source-right/Design.xlsx@r210" in theirs.author_source.replace("\\", "/")
    assert mine.revision == 205
    assert mine.author == "carol"
    assert mine.author_status == "resolved"
    assert pristine.revision == 205
    assert pristine.author == "carol"

    # Successful sources are part of the audit contract, not label-only data.
    joined = "\n".join(logs)
    assert "wc-db-exact-sidecar" in joined and "AUTHOR" in joined, (
        "resolved author source was not logged",
        logs,
    )


def test_cross_branch_repository_path_requires_exact_author_node() -> None:
    root, context = _create_local_wc_fixture()
    db_path = os.path.join(root, ".svn", "wc.db")
    with sqlite3.connect(db_path) as connection:
        connection.executemany(
            """
            insert into NODES (
                local_relpath, op_depth, kind, presence,
                repos_id, changed_revision, changed_author, checksum
            ) values (?, 0, 'file', 'normal', 1, ?, ?, null)
            """,
            (
                # Same basename/revision on another branch must not steal the
                # conflict-source author attribution.
                ("target-copy/Design.xlsx", 200, "wrong-base-author"),
                ("target-copy/Design.xlsx", 210, "wrong-after-author"),
                # This row proves a basename fallback would have returned an
                # author if repository_identity were ignored.
                ("target/OnlyTarget.xlsx", 333, "wrong-fallback-author"),
            ),
        )
        connection.commit()

    context.identity_for("base").repository_identity = "source-left/Design.xlsx"
    context.identity_for("theirs").repository_identity = "source-right/Design.xlsx"
    with patch.object(
        smt,
        "_run_tortoise_svn_author_probe",
        return_value=(None, "test native probe unavailable"),
    ):
        identities = smt.resolve_svn_author_metadata(context)
    assert identities["base"].author == "alice"
    assert identities["theirs"].author == "bob"
    assert "wc-db-exact-source:source-left/Design.xlsx@r200" in identities["base"].author_source
    assert "wc-db-exact-source:source-right/Design.xlsx@r210" in identities["theirs"].author_source

    base = identities["base"]
    base.path = os.path.join(root, "target", "OnlyTarget.xlsx.merge-left.r333")
    base.revision = 333
    base.repository_identity = "source-missing/OnlyTarget.xlsx"
    with patch.object(
        smt,
        "_run_tortoise_svn_author_probe",
        return_value=(None, "test native probe unavailable"),
    ):
        identities = smt.resolve_svn_author_metadata(context)
    assert identities["base"].author == "未知"
    assert identities["base"].author_status == "unavailable"
    assert "no exact source path source-missing/OnlyTarget.xlsx@r333" in identities["base"].availability_reason


def test_source_author_queries_repository_url_at_exact_peg() -> None:
    root, context = _create_local_wc_fixture()
    base = context.identity_for("base")
    base.revision = 37347
    base.repository_identity = "sheets/release/Gunships护山神兽.xlsx"
    invocations = []

    def fake_run(command, **kwargs):
        invocations.append((list(command), kwargs))
        return SimpleNamespace(
            returncode=0,
            stdout=(
                '<?xml version="1.0"?><log>'
                '<logentry revision="37347"><author>release-owner</author></logentry>'
                "</log>"
            ),
            stderr="",
        )

    with patch.object(smt, "_find_svn_cli_exe", return_value="C:/tools/svn.exe"), patch.object(
        smt.subprocess, "run", side_effect=fake_run
    ):
        author, source = smt._svn_author_for_source_identity(
            root,
            context.merged_path,
            base,
        )

    assert author == "release-owner"
    assert source.startswith("svn-log-url-peg:"), source
    assert len(invocations) == 1
    command, kwargs = invocations[0]
    assert command[:5] == ["C:/tools/svn.exe", "log", "--xml", "--non-interactive", "-r"]
    assert command[5] == "37347:37347"
    assert command[6].startswith(
        "http://svn.example.test/repository/sheets/release/Gunships"
    )
    assert command[6].endswith("@37347")
    assert kwargs["timeout"] == 12


def test_tortoise_only_author_probe_uses_wc_proven_url_and_memory_cache() -> None:
    smt._SVN_AUTHOR_MEMORY_CACHE.clear()
    probe_calls = []
    with patch.object(smt, "_find_svn_cli_exe", return_value=None), patch.object(
        smt,
        "_run_tortoise_svn_author_probe",
        side_effect=lambda url, revision: (
            probe_calls.append((url, revision)) or ("tortoise-author", "tortoise-svn-revprop-author")
        ),
    ):
        author, source = smt._svn_author_for_url_peg(
            "http://svn.example.test/repository/sheets/release/Gunships.xlsx",
            37348,
            trusted_repository_root="http://svn.example.test/repository",
            repository_uuid="fixture-repo",
        )
        cached_author, cached_source = smt._svn_author_for_url_peg(
            "http://svn.example.test/repository/sheets/release/Gunships.xlsx",
            37348,
            trusted_repository_root="http://svn.example.test/repository",
            repository_uuid="fixture-repo",
        )

    assert author == cached_author == "tortoise-author"
    assert source.startswith("tortoise-svn-revprop-author:"), source
    assert cached_source.startswith("svn-author-memory-cache:"), cached_source
    assert probe_calls == [
        ("http://svn.example.test/repository/sheets/release/Gunships.xlsx", 37348)
    ]


def test_tortoise_probe_uses_internal_result_file_and_handles_child_failure() -> None:
    captured = []

    def successful_child(command, **kwargs):
        captured.append((list(command), kwargs))
        with open(command[-1], "w", encoding="utf-8") as stream:
            stream.write('{"author":"isolated-author"}')
        return SimpleNamespace(returncode=0)

    with patch.object(smt, "_find_tortoise_svn_bin_dir", return_value="C:/Tortoise/bin"), patch.object(
        smt.subprocess, "run", side_effect=successful_child
    ):
        author, reason = smt._run_tortoise_svn_author_probe(
            "http://svn.example.test/repository/sheets/release/Gunships.xlsx", 37348
        )

    assert author == "isolated-author"
    assert reason == "tortoise-svn-revprop-author"
    command, kwargs = captured[0]
    assert "--internal-svn-author-query" in command
    assert kwargs["stdout"] is smt.subprocess.DEVNULL
    assert kwargs["stderr"] is smt.subprocess.DEVNULL
    assert kwargs["timeout"] == 8
    assert not os.path.exists(command[-1])

    with patch.object(smt, "_find_tortoise_svn_bin_dir", return_value="C:/Tortoise/bin"), patch.object(
        smt.subprocess, "run", side_effect=subprocess.TimeoutExpired(["probe"], 8)
    ):
        author, reason = smt._run_tortoise_svn_author_probe(
            "http://svn.example.test/repository/sheets/release/Gunships.xlsx", 37348
        )
    assert author is None
    assert reason == "TortoiseSVN author probe timed out"


def test_author_labels_include_revision_status_and_local_edits() -> None:
    _root, context = _create_local_wc_fixture()
    with patch.object(
        smt,
        "_run_tortoise_svn_author_probe",
        return_value=(None, "test native probe unavailable"),
    ):
        identities = smt.resolve_svn_author_metadata(context)

    base_label = smt.format_version_identity(identities["base"])
    assert "Design.xlsx.merge-left.r200" in base_label
    assert "@r200" in base_label
    assert "Author = alice" in base_label

    identities["mine"].locally_modified = True
    mine_label = smt.format_version_identity(identities["mine"], candidate=True)
    assert "Design.xlsx" in mine_label
    assert "@r205" in mine_label
    assert "Author = carol" in mine_label
    assert "本地未提交修改" in mine_label
    assert "合并候选" in mine_label


def test_unknown_author_has_visible_reason_and_log_evidence() -> None:
    root = make_temp_dir("sow_svn_author_unknown_")
    base = os.path.join(root, "Design.xlsx.r10")
    mine = os.path.join(root, "Design.xlsx")
    theirs = os.path.join(root, "Design.xlsx.r12")
    for path, marker in ((base, "base"), (mine, "mine"), (theirs, "theirs")):
        _make_book(path, marker)
    context = smt.build_merge_launch_context(base, mine, theirs, mine)

    logs = []
    with patch.object(smt, "_dlog", lambda message: logs.append(str(message))):
        identities = smt.resolve_svn_author_metadata(context)

    for role in ("base", "mine", "theirs"):
        identity = identities[role]
        assert identity.author == "未知", (role, identity)
        assert identity.author_status == "unavailable", (role, identity)
        assert identity.availability_reason, (role, identity)
        label = smt.format_version_identity(identity)
        assert "Author = 未知" in label, (role, label)
    joined = "\n".join(logs)
    for role in ("base", "mine", "theirs"):
        assert f"AUTHOR role={role} unavailable" in joined, joined
    assert "reason=" in joined and (
        "working-copy metadata unavailable" in joined
        or "sidecar revision unavailable" in joined
        or "wc.db" in joined
    ), joined


def test_startup_diagnostic_log_contains_reproducible_decision_evidence() -> None:
    _root, context = _create_local_wc_fixture()
    # Keep a separate pristine identity and make it equal to Mine so the log
    # contains the clean/modified decision in addition to the six matrix pairs.
    shutil.copy2(context.mine_path, context.target_pristine_path)
    logs = []
    with patch.object(smt, "_dlog", lambda message: logs.append(str(message))):
        # Rebuild while logging so the raw identity/classification block is
        # captured in the same diagnostic evidence stream as the analysis.
        context = smt.build_merge_launch_context(
            context.source_base_path,
            context.mine_path,
            context.theirs_path,
            context.merged_path,
            target_pristine_path=context.target_pristine_path,
        )
        analysis = smt.run_startup_merge_analysis(context)

    assert analysis.context is context
    joined = "\n".join(logs)
    required_tokens = (
        "MERGE_CONTEXT",
        "scenario=cross-branch-merge",
        "raw_base=",
        "raw_mine=",
        "raw_theirs=",
        "EQUIVALENCE",
        "base<->mine",
        "mine<->target_pristine",
        "MINE_LOCAL_STATE",
        "STARTUP_OUTCOME",
        "STARTUP_ANALYSIS_COMPLETE",
        "merged=",
        "unresolved=",
    )
    for token in required_tokens:
        assert token in joined, (token, joined)
    assert joined.count("EQUIVALENCE ") == 6, joined
    assert "wc-db-exact-sidecar" in joined, (
        "diagnostic evidence omits resolved author sources",
        joined,
    )


def main() -> None:
    tests = (
        test_sidecar_suffix_is_stripped_for_exact_author_lookup,
        test_cross_branch_repository_path_requires_exact_author_node,
        test_source_author_queries_repository_url_at_exact_peg,
        test_tortoise_only_author_probe_uses_wc_proven_url_and_memory_cache,
        test_tortoise_probe_uses_internal_result_file_and_handles_child_failure,
        test_author_labels_include_revision_status_and_local_edits,
        test_unknown_author_has_visible_reason_and_log_evidence,
        test_startup_diagnostic_log_contains_reproducible_decision_evidence,
    )
    for test in tests:
        test()
        print(f"PASS: {test.__name__}")
    print(f"PASS: SVN author diagnostics ({len(tests)} tests)")


if __name__ == "__main__":
    main()
