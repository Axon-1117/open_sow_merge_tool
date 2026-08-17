"""Helpers for test temporary directories in sandboxed environments."""

import os
import uuid


def _default_test_tmp_root() -> str:
    repo_root = os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    )
    return os.path.join(repo_root, "tmp", "test_tmp")


def make_temp_dir(prefix: str) -> str:
    """Create a writable unique temp directory for tests.

    Priority:
    1) $SOW_TEST_TMPDIR
    2) <repo>/tmp/test_tmp
    """
    base = (os.environ.get("SOW_TEST_TMPDIR") or "").strip() or _default_test_tmp_root()
    os.makedirs(base, exist_ok=True)
    pfx = prefix or "sow_tmp_"
    for _ in range(200):
        candidate = os.path.join(base, f"{pfx}{uuid.uuid4().hex[:8]}")
        try:
            os.mkdir(candidate)
            return candidate
        except FileExistsError:
            continue
    raise RuntimeError(f"Failed to create temp dir under: {base}")


def visible_render_text(rendered_text: object, *, placeholder: str = "\u200b") -> str:
    """Return what a user sees after removing zero-width Tk index padding."""
    return str(rendered_text).replace(str(placeholder), "")
