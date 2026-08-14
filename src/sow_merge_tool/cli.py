"""Stable CLI entry point with a cheap branch-submit dispatch path."""

from __future__ import annotations

import sys


def _run_branch_submit(argv: list[str]) -> None:
    from .branch_submit import launch_ui

    initial_paths: list[str] = []
    index = 0
    while index < len(argv):
        value = argv[index]
        if value == "--branch-file" and index + 1 < len(argv):
            initial_paths.append(argv[index + 1])
            index += 2
            continue
        if value not in {"--branch-submit", "--branch-file"} and not value.startswith("-"):
            initial_paths.append(value)
        index += 1
    launch_ui(initial_paths)


def run() -> None:
    argv = list(sys.argv[1:])
    if "--branch-submit" in argv:
        _run_branch_submit(argv)
        return
    if argv and argv[0] == "--internal-svn-status-query":
        from .svn_status_provider import internal_status_entrypoint

        raise SystemExit(internal_status_entrypoint(argv[1:]))
    from .legacy_core import run_entrypoint

    run_entrypoint()


if __name__ == "__main__":
    run()
