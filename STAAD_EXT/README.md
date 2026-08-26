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

The exporter prompts for a DXF destination, projection plane, label option, and
text scale. It attaches to the running STAAD.Pro instance; it does not start a
new one.

Choose **STD to GA Drawing** for the general-arrangement version of the same
selection: every member carries a bubbled mark number instead of a label, and
the section descriptions are collected into one MEMBER SIZE SCHEDULE placed in a
chosen corner. Members sharing a section size *and* a steel grade share a mark.
The direct command is `staad-ext std-to-ga-dxf`.

Choose **Taper Optimizer** to size the tapered I sections of a 2D frame down to
the lightest plates that still pass STAAD's own code check and your deflection
limits. It optimizes the selected tapered members, or every tapered member when
nothing is selected; geometry, loading and non-tapered sections are never
changed. The direct command is `staad-ext optimize-tapers`.

Sizes come off fixed fabrication ladders — web depth from 250mm in 50mm steps,
flange width from 150mm in 25mm steps, both plate thicknesses off the stocked
list, and the flange always thicker than the web. The two flanges are kept
identical, and web depth is solved once per connected node so members
continuing through a node always share a depth.

The model needs a `PARAMETER` / `CHECK CODE` block covering the tapered members:
the optimizer judges candidates by STAAD's design ratios, not by a check of its
own. You supply the utilisation ceiling, the vertical and horizontal deflection
limits, and the load combinations to check them against. Each candidate costs a
full analysis, so the run is capped by an analysis budget; most of the saving
lands in the first ten runs. The default is a dry run that reports what it
would assign and puts every member back on the section it started with.

For development and tests, install `python -m pip install -e ".[dev]"`, then run
`pytest`.

## Layout

- `staad_ext.openstaad`: reusable STAAD.Pro COM adapter
- `staad_ext.application`: desktop home screen and utility registry
- `staad_ext.workflows`: user-facing utility workflows
- `staad_ext.framing`: shared model read and geometry solve for both DXF exporters
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
