# STAAD_EXT

An extensible Python project for small STAAD.Pro 2025 utilities.

## Setup

Use the same Python architecture (normally 64-bit) as STAAD.Pro:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e .
```

Open a STAAD model, select one or more analytical beam members, then run:

```powershell
staad-ext std-to-dxf
```

For development and tests, install `python -m pip install -e ".[dev]"`, then run
`pytest`.

The exporter prompts for a DXF destination, projection plane, label option, and
text scale. It attaches to the running STAAD.Pro instance; it does not start a
new one.

## Layout

- `staad_ext.openstaad`: reusable STAAD.Pro COM adapter
- `staad_ext.dxf`: dependency-free ASCII DXF writer
- `staad_ext.macros`: individual user-facing macros
- `tests`: unit tests that do not require STAAD.Pro
