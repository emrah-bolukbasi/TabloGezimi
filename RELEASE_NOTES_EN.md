# Table Navigation 1.1.11 — Update from 1.1.10

Table Navigation is an NVDA add-on for accessible tables in Microsoft Excel, web browsers and accessible PDF/document readers.

## What's new

- **Two navigation modes.** Mode 1 reads whole rows with Up/Down Arrow; mode 2 reads only the selected cell while using arrow keys. Available in desktop Excel, web browsers and accessible PDF tables. Toggle with **Ctrl+Shift+Alt+Up Arrow**. Each new table-navigation session starts in mode 1.
- **Configurable cell pause.** Add a pause between cells in full-row speech. Default **150 ms**, configurable from **0 to 5000 ms** with **Ctrl+Shift+Alt+Down Arrow** in a Turkish input dialog; the shortcut is customizable via NVDA Input Gestures.
- **Freer Excel navigation.** Arrow keys can move beyond the current data region. An empty destination cell announces its address (for example **A5**).
- **Smarter Excel data operations.** Blank values are omitted during bulk reading and copying. Copying is scoped to the active contiguous table region, with fully blank rows or columns separating regions.
- **Persistent toggle.** When table navigation is left enabled (or disabled), this state is restored after restarting NVDA. Toggle with **Alt+Shift+O**.
- **Preserved copy commands.** Alt+Shift+Up/Down/Left/Right Arrow still copy the row/table/column/cell in either navigation mode.

## Installation

Download **TabloGezimi-1.1.11.nvda-addon** from the **v1.1.11** release assets. Open it, accept NVDA's installation prompt, then restart NVDA. The release is **not yet published**; these are prepared release notes.

## Compatibility and limitations

- Manifest declares NVDA 2026.1.0 minimum and 2026.3.0 last tested; compatibility must still be verified on the target NVDA/Windows system before publishing.
- Excel integration uses desktop Windows Excel COM. PDF tables require accessible table information; no OCR is bundled.
- Separate data blocks without fully empty rows/columns between them may not always be detected as separate tables.
- The add-on retains nine language resources; newly added mode and dialog announcements are currently written in Turkish and should be reviewed for translation before release.

License: GPL-2.0-only. Maintainer: Emrah Bölükbaşı.
