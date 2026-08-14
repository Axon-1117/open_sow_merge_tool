"""Public Python package for the Excel merge tool.

The legacy implementation is kept behind this package boundary while the
remaining modules are migrated.  Re-exporting the legacy symbols here keeps
the executable and existing automation on one import path during the
refactor; new code must depend on the service and adapter modules instead of
reaching into this compatibility surface.
"""

from . import legacy_core as _legacy_core

for _name, _value in vars(_legacy_core).items():
    if _name not in {"__name__", "__package__", "__loader__", "__spec__"}:
        globals()[_name] = _value

__all__ = tuple(
    _name
    for _name in globals()
    if not _name.startswith("__") and _name not in {"_legacy_core", "_name", "_value"}
)
