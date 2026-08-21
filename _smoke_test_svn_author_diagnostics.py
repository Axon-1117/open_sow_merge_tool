"""OpenSpec 4.5 regressions for SVN author labels and audit logging."""

from __future__ import annotations

import hashlib
import os
import shutil
import sqlite3
import subprocess
import tempfile
import time
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

import sow_merge_tool as smt


_CASE_DEADLINE: float | None = None


class _OwnedCase:
    def __init__(self, name: str):
        self._temporary = tempfile.TemporaryDirectory(prefix=f"sow_svn_author_{name}_")
        self.root = self._temporary.name
        self.startup_artifacts = os.path.join(self.root, "startup-artifacts")
        os.makedirs(self.startup_artifacts, exist_ok=False)
        self.input_hashes: dict[str, str] = {}
        self.startup_ledger: set[str] = set()
        self.expect_startup_evidence = False

    def record_input(self, path: str) -> None:
        absolute = os.path.abspath(path)
        self.input_hashes[absolute] = _sha256(absolute)

    def cleanup(self) -> None:
        self._temporary.cleanup()


def _sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _checkpoint(label: str) -> None:
    if _CASE_DEADLINE is not None and time.monotonic() > _CASE_DEADLINE:
        raise TimeoutError(f"SVN author diagnostics exceeded 90 seconds at {label}")


def _inside(root: str, path: str) -> bool:
    try:
        return os.path.normcase(os.path.commonpath((os.path.abspath(root), os.path.abspath(path)))) == os.path.normcase(os.path.abspath(root))
    except (TypeError, ValueError):
        return False


def _make_book(path: str, marker: str, *, owned: _OwnedCase) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet["A1"] = "marker"
    worksheet["A2"] = marker
    workbook.save(path)
    workbook.close()
    owned.record_input(path)


def _create_local_wc_fixture(owned: _OwnedCase):
    root = owned.root
    os.makedirs(os.path.join(root, ".svn"), exist_ok=True)
    db_path = os.path.join(root, ".svn", "wc.db")
    connection = sqlite3.connect(db_path)
    try:
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
    finally:
        connection.close()

    target_dir = os.path.join(root, "target")
    os.makedirs(target_dir, exist_ok=True)
    base = os.path.join(target_dir, "Design.xlsx.merge-left.r200")
    mine = os.path.join(target_dir, "Design.xlsx")
    theirs = os.path.join(target_dir, "Design.xlsx.merge-right.r210")
    pristine = os.path.join(target_dir, "Design.target-pristine.xlsx")
    _make_book(base, "base", owned=owned)
    _make_book(mine, "mine-local", owned=owned)
    _make_book(theirs, "theirs", owned=owned)
    _make_book(pristine, "mine-pristine", owned=owned)
    # Record only after all WC and workbook setup is complete.
    owned.record_input(db_path)
    context = smt.build_merge_launch_context(
        base,
        mine,
        theirs,
        mine,
        target_pristine_path=pristine,
    )
    return root, context


def test_sidecar_suffix_is_stripped_for_exact_author_lookup(owned: _OwnedCase) -> None:
    _root, context = _create_local_wc_fixture(owned)
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

    joined = "\n".join(logs)
    assert "wc-db-exact-sidecar" in joined and "AUTHOR" in joined, (
        "resolved author source was not logged",
        logs,
    )


def test_cross_branch_repository_path_requires_exact_author_node(owned: _OwnedCase) -> None:
    root, context = _create_local_wc_fixture(owned)
    db_path = os.path.join(root, ".svn", "wc.db")
    connection = sqlite3.connect(db_path)
    try:
        connection.executemany(
            """
            insert into NODES (
                local_relpath, op_depth, kind, presence,
                repos_id, changed_revision, changed_author, checksum
            ) values (?, 0, 'file', 'normal', 1, ?, ?, null)
            """,
            (
                ("target-copy/Design.xlsx", 200, "wrong-base-author"),
                ("target-copy/Design.xlsx", 210, "wrong-after-author"),
                ("target/OnlyTarget.xlsx", 333, "wrong-fallback-author"),
            ),
        )
        connection.commit()
    finally:
        connection.close()
    owned.record_input(db_path)

    context.identity_for("base").repository_identity = "source-left/Design.xlsx"
    context.identity_for("theirs").repository_identity = "source-right/Design.xlsx"
    identities = smt.resolve_svn_author_metadata(context)
    assert identities["base"].author == "alice"
    assert identities["theirs"].author == "bob"
    assert "wc-db-exact-source:source-left/Design.xlsx@r200" in identities["base"].author_source
    assert "wc-db-exact-source:source-right/Design.xlsx@r210" in identities["theirs"].author_source

    base = identities["base"]
    base.path = os.path.join(root, "target", "OnlyTarget.xlsx.merge-left.r333")
    base.revision = 333
    base.repository_identity = "source-missing/OnlyTarget.xlsx"
    identities = smt.resolve_svn_author_metadata(context)
    assert identities["base"].author == "未知"
    assert identities["base"].author_status == "unavailable"
    assert "no exact source path source-missing/OnlyTarget.xlsx@r333" in identities["base"].availability_reason


def test_source_author_queries_repository_url_at_exact_peg(owned: _OwnedCase) -> None:
    root, context = _create_local_wc_fixture(owned)
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
        author, source = smt._svn_author_for_source_identity(root, context.merged_path, base)

    assert author == "release-owner"
    assert source.startswith("svn-log-url-peg:"), source
    assert len(invocations) == 1
    command, kwargs = invocations[0]
    assert command[:5] == ["C:/tools/svn.exe", "log", "--xml", "--non-interactive", "-r"]
    assert command[5] == "37347:37347"
    assert command[6].startswith("http://svn.example.test/repository/sheets/release/Gunships")
    assert command[6].endswith("@37347")
    assert kwargs["timeout"] == 12


def test_tortoise_only_author_probe_uses_wc_proven_url_and_memory_cache(owned: _OwnedCase) -> None:
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


def test_tortoise_probe_uses_internal_result_file_and_handles_child_failure(owned: _OwnedCase) -> None:
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
    assert _inside(owned.startup_artifacts, command[-1]), command
    assert not os.path.exists(command[-1])

    with patch.object(smt, "_find_tortoise_svn_bin_dir", return_value="C:/Tortoise/bin"), patch.object(
        smt.subprocess, "run", side_effect=subprocess.TimeoutExpired(["probe"], 8)
    ):
        author, reason = smt._run_tortoise_svn_author_probe(
            "http://svn.example.test/repository/sheets/release/Gunships.xlsx", 37348
        )
    assert author is None
    assert reason == "TortoiseSVN author probe timed out"
    assert not os.listdir(owned.startup_artifacts), os.listdir(owned.startup_artifacts)


def test_author_labels_include_revision_status_and_local_edits(owned: _OwnedCase) -> None:
    _root, context = _create_local_wc_fixture(owned)
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


def test_unknown_author_has_visible_reason_and_log_evidence(owned: _OwnedCase) -> None:
    root = owned.root
    base = os.path.join(root, "Design.xlsx.r10")
    mine = os.path.join(root, "Design.xlsx")
    theirs = os.path.join(root, "Design.xlsx.r12")
    for path, marker in ((base, "base"), (mine, "mine"), (theirs, "theirs")):
        _make_book(path, marker, owned=owned)
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


def test_startup_diagnostic_log_contains_reproducible_decision_evidence(owned: _OwnedCase) -> None:
    _root, context = _create_local_wc_fixture(owned)
    shutil.copy2(context.mine_path, context.target_pristine_path)
    owned.record_input(context.target_pristine_path)
    logs = []
    owned.expect_startup_evidence = True
    with patch.object(smt, "_dlog", lambda message: logs.append(str(message))):
        context = smt.build_merge_launch_context(
            context.source_base_path,
            context.mine_path,
            context.theirs_path,
            context.merged_path,
            target_pristine_path=context.target_pristine_path,
        )
        analysis = smt.run_startup_merge_analysis(
            context,
            owned_startup_paths=owned.startup_ledger,
        )

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


def _forbid_subprocess(*_args, **_kwargs):
    raise AssertionError("author fixture path must not launch an external subprocess")


def _run_owned_case(test) -> None:
    owned = _OwnedCase(test.__name__)
    cache_before = dict(smt._SVN_AUTHOR_MEMORY_CACHE)
    primary: BaseException | None = None
    cleanup_errors: list[str] = []
    try:
        with (
            patch.object(smt, "_find_svn_cli_exe", lambda: None),
            patch.object(smt, "_find_tortoise_svn_bin_dir", lambda: None),
            patch.object(smt.subprocess, "run", _forbid_subprocess),
            patch.object(smt.tempfile, "gettempdir", lambda: owned.startup_artifacts),
        ):
            _checkpoint(f"before:{test.__name__}")
            test(owned)
            _checkpoint(f"after:{test.__name__}")
    except BaseException as exc:
        primary = exc
    finally:
        try:
            smt._SVN_AUTHOR_MEMORY_CACHE.clear()
            smt._SVN_AUTHOR_MEMORY_CACHE.update(cache_before)
        except Exception as exc:
            cleanup_errors.append(f"author cache restore failed: {type(exc).__name__}: {exc}")
        try:
            for path, expected_hash in owned.input_hashes.items():
                actual_hash = _sha256(path)
                if actual_hash != expected_hash:
                    cleanup_errors.append(
                        f"input hash changed: {path} {expected_hash} -> {actual_hash}"
                    )
        except Exception as exc:
            cleanup_errors.append(f"input hash verification failed: {type(exc).__name__}: {exc}")
        try:
            evidence_sink: list[dict] = []
            evidence = smt._consume_owned_startup_temp_paths(
                owned.startup_ledger,
                evidence_sink,
            )
            if tuple(evidence_sink) != evidence or owned.startup_ledger:
                cleanup_errors.append("startup ledger was not consumed exactly once")
            if owned.expect_startup_evidence and not evidence:
                cleanup_errors.append("startup diagnostic created no owned cleanup evidence")
            for fact in evidence:
                if (
                    not _inside(owned.startup_artifacts, fact.get("path", ""))
                    or not fact.get("removed")
                    or fact.get("exists_after")
                    or fact.get("error")
                ):
                    cleanup_errors.append(f"startup cleanup evidence invalid: {fact}")
        except Exception as exc:
            cleanup_errors.append(f"startup ledger cleanup failed: {type(exc).__name__}: {exc}")
        try:
            if os.path.isdir(owned.startup_artifacts) and os.listdir(owned.startup_artifacts):
                cleanup_errors.append(
                    f"SVN author result/temp residue: {os.listdir(owned.startup_artifacts)}"
                )
        except Exception as exc:
            cleanup_errors.append(f"startup artifact audit failed: {type(exc).__name__}: {exc}")
        try:
            owned.cleanup()
            if os.path.lexists(owned.root):
                cleanup_errors.append(f"owned temporary root retained: {owned.root}")
        except Exception as exc:
            cleanup_errors.append(f"owned temporary cleanup failed: {type(exc).__name__}: {exc}")
    if primary is not None:
        for detail in cleanup_errors:
            try:
                primary.add_note(detail)
            except Exception:
                pass
        raise primary
    if cleanup_errors:
        raise AssertionError("; ".join(cleanup_errors))


def main() -> None:
    global _CASE_DEADLINE
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
    _CASE_DEADLINE = time.monotonic() + 90.0
    try:
        for test in tests:
            _run_owned_case(test)
            print(f"PASS: {test.__name__}")
        print(f"PASS: SVN author diagnostics ({len(tests)} tests)")
    finally:
        _CASE_DEADLINE = None


if __name__ == "__main__":
    main()
