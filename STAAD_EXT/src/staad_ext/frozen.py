"""Entry point used by the frozen Windows application."""

from staad_ext.application import launch


if __name__ == "__main__":
    raise SystemExit(launch())
