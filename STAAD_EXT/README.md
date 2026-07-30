# STAAD_EXT

An extensible Python project for small STAAD.Pro 2025 utilities.

## Setup

Use the same Python architecture (normally 64-bit) as STAAD.Pro:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e .
```

Open the STAAD_EXT main screen with:

```powershell
staad-ext
```

Choose **Plate Summary** to show a table for the selected analytical members.
Tapered I-sections are split into a tapered web, top flange, and bottom flange;
all other section types are kept as whole-section items. The direct command is
`staad-ext plate-summary`.


Choose **STD to DXF**, then select one or more analytical beam members in an
open STAAD model before exporting. The direct command remains available for
development and automation: `staad-ext std-to-dxf`.

For development and tests, install `python -m pip install -e ".[dev]"`, then run
`pytest`.

The exporter prompts for a DXF destination, projection plane, label option, and
text scale. It attaches to the running STAAD.Pro instance; it does not start a
new one.

## Layout

- `staad_ext.openstaad`: reusable STAAD.Pro COM adapter
- `staad_ext.application`: desktop home screen and utility registry
- `staad_ext.workflows`: user-facing utility workflows
- `staad_ext.dxf`: dependency-free ASCII DXF writer
- `staad_ext.macros`: individual user-facing macros
- `tests`: unit tests that do not require STAAD.Pro

## Windows releases

Pushing a version tag such as `v0.1.0` runs the
`Release Windows executable` GitHub Actions workflow. It tests the project,
builds the application in one-folder mode, creates ZIP and MSI packages with
SHA-256 checksums, and attaches them to a new GitHub release. No repository
secrets are required. Install `STAAD_EXT-windows-x64.msi`, or download and
extract `STAAD_EXT-windows-x64.zip` and run `STAAD_EXT.exe` directly.

The executable is not code-signed. Windows may therefore display an
`Unknown publisher` or Microsoft Defender SmartScreen warning when it is
downloaded or launched. The one-folder package avoids the self-extracting
behavior of a one-file PyInstaller build and disables UPX compression, reducing
common antivirus heuristic triggers.

To test the packaging locally:

```powershell
python -m pip install -e ".[dev,release]"
pyinstaller --noconfirm --clean STAAD_EXT.spec
```
