"""Stable CLI entry point used by the console scripts and PyInstaller spec."""

from .legacy_core import run_entrypoint


def run() -> None:
    run_entrypoint()


if __name__ == "__main__":
    run()
