"""External-system adapters used by application services."""

from .svn import scan_svn_status

__all__ = ["scan_svn_status"]
