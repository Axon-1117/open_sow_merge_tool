"""SVN adapter seam; fail-closed behavior remains in the provider module."""

from __future__ import annotations

from threading import Event

from ..svn_status_provider import SvnStatusRecord, scan_status


def scan_svn_status(path: str, *, cancel_event: Event | None = None) -> list[SvnStatusRecord]:
    return scan_status(path, cancel_event=cancel_event)
