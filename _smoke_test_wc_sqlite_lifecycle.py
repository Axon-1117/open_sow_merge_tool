"""Focused no-GUI lifecycle gate for read-only working-copy SQLite queries."""
from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import sqlite3
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import sow_merge_tool as sm


_CASE = "wc-sqlite-lifecycle"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _inside(root: Path, path: str) -> bool:
    try:
        return os.path.normcase(os.path.commonpath((str(root.resolve()), os.path.abspath(path)))) == os.path.normcase(str(root.resolve()))
    except (TypeError, ValueError):
        return False


def _checkpoint(deadline: float, label: str) -> None:
    if time.monotonic() > deadline:
        raise TimeoutError(f"wc SQLite lifecycle exceeded 90 seconds at {label}")


def _forbid_subprocess(*_args, **_kwargs):
    raise AssertionError("WC SQLite lifecycle fixture must not start a subprocess")


class _TrackedConnection:
    def __init__(self, raw, calls: list["_TrackedConnection"]):
        self._raw = raw
        self.close_calls = 0
        self.execute_calls = 0
        self.closed = False
        calls.append(self)

    def execute(self, *args, **kwargs):
        self.execute_calls += 1
        return self._raw.execute(*args, **kwargs)

    def close(self) -> None:
        self._raw.close()
        self.close_calls += 1
        self.closed = True

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return self._raw.__exit__(exc_type, exc, traceback)

    def __getattr__(self, name):
        return getattr(self._raw, name)


@dataclass(frozen=True)
class _Fixture:
    root: Path
    db_path: Path
    target: Path
    pristine: Path
    artifacts: Path
    input_hashes: dict[Path, str]


def _write_fixture(root: Path) -> _Fixture:
    svn_root = root / ".svn"
    svn_root.mkdir()
    db_path = svn_root / "wc.db"
    connection = sqlite3.connect(db_path)
    try:
        connection.execute("create table REPOSITORY (id integer primary key, root text, uuid text)")
        connection.execute(
            "insert into REPOSITORY (id, root, uuid) values (1, ?, ?)",
            ("https://svn.example.test/repo", "fixture-uuid"),
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
        connection.execute("create table ACTUAL_NODE (local_relpath text, conflict_data blob)")
        connection.executemany(
            """
            insert into NODES (
                local_relpath, op_depth, kind, presence, repos_id,
                changed_revision, changed_author, checksum
            ) values (?, 0, 'file', 'normal', 1, ?, ?, ?)
            """,
            (
                ("source-left/Design.xlsx", 200, "alice", None),
                ("source-left/Design.xlsx", 210, "bob", None),
                ("target/Design.xlsx", 205, "carol", None),
            ),
        )
        connection.execute(
            "insert into ACTUAL_NODE (local_relpath, conflict_data) values (?, ?)",
            (
                "target/Design.xlsx",
                "((subversion source-left/Design.xlsx 200) (subversion source-left/Design.xlsx 210))",
            ),
        )
        connection.commit()
    finally:
        connection.close()

    target_dir = root / "target"
    target_dir.mkdir()
    target = target_dir / "Design.xlsx"
    workbook = Workbook()
    try:
        workbook.active.title = "Data"
        workbook.active["A1"] = "id@id"
        workbook.active["A2"] = "string"
        workbook.active["A3"] = "target"
        workbook.save(target)
    finally:
        workbook.close()
    sha1 = hashlib.sha1(target.read_bytes()).hexdigest()
    connection = sqlite3.connect(db_path)
    try:
        connection.execute(
            "update NODES set checksum = ? where local_relpath = ?",
            (f"$sha1${sha1}", "target/Design.xlsx"),
        )
        connection.commit()
    finally:
        connection.close()
    pristine = svn_root / "pristine" / sha1[:2] / f"{sha1}.svn-base"
    pristine.parent.mkdir(parents=True)
    shutil.copyfile(target, pristine)
    artifacts = root / "startup-artifacts"
    artifacts.mkdir()
    input_paths = (db_path, target, pristine)
    return _Fixture(root, db_path, target, pristine, artifacts, {path: _sha256(path) for path in input_paths})


def _exercise_all_wc_queries(
    fixture: _Fixture, deadline: float, ledger: set[str]
) -> list[_TrackedConnection]:
    real_connect = sqlite3.connect
    tracked: list[_TrackedConnection] = []

    def tracked_connect(*args, **kwargs):
        return _TrackedConnection(real_connect(*args, **kwargs), tracked)

    with (
        patch.object(sm.sqlite3, "connect", tracked_connect),
        patch.object(sm, "_find_svn_cli_exe", lambda: None),
        patch.object(sm, "_find_tortoise_svn_bin_dir", lambda: None),
        patch.object(sm.subprocess, "run", _forbid_subprocess),
        patch.object(sm.tempfile, "gettempdir", lambda: str(fixture.artifacts)),
    ):
        _checkpoint(deadline, "cross-source")
        sources, reason = sm._read_cross_branch_sources_from_wc(str(fixture.root), "target/Design.xlsx")
        assert sources == [("source-left/Design.xlsx", 200), ("source-left/Design.xlsx", 210)]
        assert reason == "wc-conflict-data"

        _checkpoint(deadline, "node")
        revision, author, reason = sm._wc_node_metadata(str(fixture.root), "target/Design.xlsx")
        assert (revision, author, reason) == (205, "carol", "wc-db-current-node")

        _checkpoint(deadline, "exact-source")
        author, reason = sm._wc_author_for_exact_sidecar_revision(
            str(fixture.root),
            "Design.xlsx.merge-left.r200",
            200,
            repository_identity="source-left/Design.xlsx",
        )
        assert author == "alice" and reason == "wc-db-exact-source:source-left/Design.xlsx@r200"

        _checkpoint(deadline, "exact-sidecar")
        author, reason = sm._wc_author_for_exact_sidecar_revision(
            str(fixture.root), "Design.xlsx.merge-left.r200", 200
        )
        assert author == "alice" and reason == "wc-db-exact-sidecar:source-left/Design.xlsx@r200"

        _checkpoint(deadline, "repository")
        repo_root, repo_uuid, reason = sm._wc_repository_metadata_for_local_path(
            str(fixture.root), "target/Design.xlsx"
        )
        assert (repo_root, repo_uuid, reason) == (
            "https://svn.example.test/repo",
            "fixture-uuid",
            "wc-db-repository-root",
        )

        _checkpoint(deadline, "pristine")
        copied = sm._copy_wc_pristine_local(str(fixture.target), owned_paths=ledger)
        assert copied and _inside(fixture.artifacts, copied) and Path(copied).is_file()
        assert _sha256(Path(copied)) == _sha256(fixture.pristine)

    assert len(tracked) == 6, tracked
    for index, connection in enumerate(tracked):
        assert connection.execute_calls == 1, (index, connection.execute_calls)
        assert connection.close_calls == 1 and connection.closed, (
            index,
            connection.close_calls,
            connection.closed,
        )
    return tracked


def run_case() -> None:
    temporary = tempfile.TemporaryDirectory(prefix="sow_wc_sqlite_lifecycle_")
    fixture: _Fixture | None = None
    ledger: set[str] = set()
    evidence: tuple[dict, ...] = ()
    evidence_sink: list[dict] = []
    primary: BaseException | None = None
    cleanup_errors: list[str] = []
    try:
        deadline = time.monotonic() + 90.0
        fixture = _write_fixture(Path(temporary.name))
        _tracked = _exercise_all_wc_queries(fixture, deadline, ledger)

        renamed = fixture.db_path.with_name("wc.db.lifecycle-rename")
        os.replace(fixture.db_path, renamed)
        os.replace(renamed, fixture.db_path)
        assert _sha256(fixture.db_path) == fixture.input_hashes[fixture.db_path]

        evidence = sm._consume_owned_startup_temp_paths(ledger, evidence_sink)
        assert tuple(evidence_sink) == evidence and not ledger and len(evidence) == 1, evidence
        fact = evidence[0]
        assert (
            _inside(fixture.artifacts, fact.get("path", ""))
            and fact.get("removed") is True
            and fact.get("exists_after") is False
            and not fact.get("error")
        ), fact
        for path, expected in fixture.input_hashes.items():
            assert _sha256(path) == expected, path
        _checkpoint(deadline, "post-query-cleanup")
    except BaseException as exc:
        primary = exc
    finally:
        if fixture is not None:
            try:
                for path, expected in fixture.input_hashes.items():
                    if _sha256(path) != expected:
                        cleanup_errors.append(f"input hash changed: {path}")
            except Exception as exc:
                cleanup_errors.append(f"input hash audit failed: {type(exc).__name__}: {exc}")
            try:
                if ledger:
                    final_sink: list[dict] = []
                    final_evidence = sm._consume_owned_startup_temp_paths(ledger, final_sink)
                    if tuple(final_sink) != final_evidence or ledger:
                        cleanup_errors.append("startup ledger was not consumed exactly once")
                    for fact in final_evidence:
                        if (
                            not _inside(fixture.artifacts, fact.get("path", ""))
                            or not fact.get("removed")
                            or fact.get("exists_after")
                            or fact.get("error")
                        ):
                            cleanup_errors.append(f"startup cleanup evidence invalid: {fact}")
            except Exception as exc:
                cleanup_errors.append(f"startup ledger cleanup failed: {type(exc).__name__}: {exc}")
            try:
                if fixture.artifacts.is_dir() and tuple(fixture.artifacts.iterdir()):
                    cleanup_errors.append(
                        f"startup artifact residue: {tuple(path.name for path in fixture.artifacts.iterdir())}"
                    )
            except Exception as exc:
                cleanup_errors.append(f"startup artifact audit failed: {type(exc).__name__}: {exc}")
        try:
            temporary.cleanup()
            if os.path.lexists(temporary.name):
                cleanup_errors.append(f"owned WC root retained: {temporary.name}")
        except Exception as exc:
            cleanup_errors.append(f"owned WC cleanup failed: {type(exc).__name__}: {exc}")
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
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--case", choices=(_CASE,))
    parser.add_argument("--list-cases", action="store_true")
    args = parser.parse_args()
    if args.list_cases:
        print(_CASE)
        return
    run_case()
    print("PASS " + (args.case or _CASE))


if __name__ == "__main__":
    main()
