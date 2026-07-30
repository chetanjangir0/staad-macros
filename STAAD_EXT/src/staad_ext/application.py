"""Public desktop application entry point."""

from staad_ext.desktop import StaadExtApplication

__all__ = ["StaadExtApplication", "launch"]


def launch() -> int:
    return StaadExtApplication().run()
