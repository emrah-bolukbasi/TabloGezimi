# Tablo Gezimi / Table Navigation

**NVDA add-on, version 1.1.11 (release candidate)** · Maintainer: Emrah Bölükbaşı

NVDA add-on for accessible table navigation, row reading and clipboard copying in desktop Microsoft Excel, web pages and accessible document/PDF tables. Developed from the published 1.1.10 release, with the accumulated improvements of internal builds 1.1.11–1.1.17. These internal versions are consolidated into the next public version, **1.1.11**.

## Features

- Two navigation modes: **mode 1** reads whole rows on Up/Down and cells on Left/Right; **mode 2** reads only the destination cell in each arrow-key direction. Both modes are available in supported Excel, web and PDF tables.
- Configurable pause between spoken row cells, default **150 ms**, range **0–5000 ms**.
- Table navigation enabled/disabled state persists across NVDA restarts; the navigation mode itself resets to mode 1 each time table navigation is enabled.
- Excel arrow-key movement is not restricted to `CurrentRegion`. Empty target cells announce their address (for example A5).
- Excel row/column/table read and copy operations omit empty values; copied table data is scoped to the current contiguous region (`CurrentRegion`). Fully blank rows and columns separate regions, while Excel-defined table boundaries can be used during target reading.
- Clipboard copying of the current row, column, cell or entire table.
- Nine bundled language resources: Turkish, English, Spanish, Russian, Arabic, Simplified Chinese, German, French, Italian. **Newer dialog and mode announcements are currently written in Turkish in the implementation; updated translations should be reviewed before a multilingual public release.**

## Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| Alt+Shift+O | Toggle table navigation (saved across restarts) |
| Ctrl+Shift+Alt+Up Arrow | Switch mode 1 / mode 2 |
| Ctrl+Shift+Alt+Down Arrow | Set cell pause duration (milliseconds) |
| Up/Down Arrow | Mode 1: read previous/next row; Mode 2: read previous/next single cell in current column |
| Left/Right Arrow | Move to previous/next cell |
| Alt+Shift+Up Arrow | Copy current row |
| Alt+Shift+Down Arrow | Copy current table |
| Alt+Shift+Left Arrow | Copy current column |
| Alt+Shift+Right Arrow | Copy current cell |

The mode-switch command and pause dialog shortcut are available in NVDA's Input Gestures. Copy commands require table navigation to be enabled. Mode 1 is selected whenever the feature is switched on, even after a saved enabled state is restored at NVDA startup.

## Installation

Download `TabloGezimi-1.1.11.nvda-addon` from the **v1.1.11** GitHub release *after it has been published*, open the file, confirm installation in NVDA and restart NVDA. At present this is a preparation package: **the new release is not yet published**.

## Compatibility and limitations

Declared minimum NVDA version: 2026.1.0; declared last-tested version: 2026.3.0, inherited from the existing manifest, **not independently verified here**. Requires Windows for Excel COM integration. Excel Online, LibreOffice and Google Sheets are not covered by the Excel-specific backend. PDF/document table support depends on accessible table metadata exposed to NVDA; image-only/scanned PDFs require separate OCR. Some table boundaries are only detectable when a completely empty row/column separates data blocks. Cell formatting, merged cells and applications may affect outputs.

## Source, license and release notes

Python implementation: `globalPlugins/tabloGezimi/`; native DLL source and build instructions: `native/`; bundled binary DLLs: `globalPlugins/tabloGezimi/x86/` and `x64/`; translations: `locale/`; HTML help: `doc/`. Native binaries are supplied; they have **not** been independently rebuilt for this update. See [CHANGELOG.md](CHANGELOG.md), [LICENSE](LICENSE) and [COPYRIGHT.md](COPYRIGHT.md). GPL-2.0-only.

Maintainer: Emrah Bölükbaşı · [GitHub repository](https://github.com/emrah-bolukbasi/TabloGezimi)
