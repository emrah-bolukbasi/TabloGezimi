# Tablo Gezimi / Table Navigation

**NVDA add-on, version 1.1.10** · Maintainer: Emrah Bölükbaşı

An NVDA add-on to navigate accessible tables and read complete rows with their column headings. The add-on supports desktop Microsoft Excel via Windows COM and the accessible table interfaces exposed by web browsers and document/PDF readers.

## Features

- Move through accessible table rows and cells with the arrow keys while table navigation is enabled.
- Read a complete row, including column headings, with Up/Down Arrow.
- Copy the current cell, row, column, or table to the clipboard.
- For desktop Excel, detect the active cell's contiguous data region (`CurrentRegion`).
- Localized UI messages and help in **Turkish, English, Spanish, Russian, Arabic, Simplified Chinese, German, French and Italian**. Unsupported NVDA interface languages fall back to English.

## Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| Alt+Shift+O | Toggle table navigation |
| Up/Down Arrow | Previous/next row, read the row |
| Left/Right Arrow | Previous/next cell |
| Alt+Shift+Up Arrow | Copy current row |
| Alt+Shift+Down Arrow | Copy entire table |
| Alt+Shift+Left Arrow | Copy current column |
| Alt+Shift+Right Arrow | Copy current cell |

The navigation mode must be enabled for the copy commands.

## Installation

Download `TabloGezimi-1.1.10.nvda-addon` from [Releases](https://github.com/emrah-bolukbasi/TabloGezimi/releases), open it, confirm installation in NVDA, then restart NVDA.

## Compatibility and limitations

The manifest declares NVDA 2026.1.0 minimum and 2026.3.0 last tested. These values come from the supplied add-on; independent testing against these NVDA releases has not been performed here. The Excel-specific component targets installed Microsoft Excel on Windows. Excel Online, Google Sheets, LibreOffice, and untagged or scanned PDF tables are not guaranteed to work. PDF support depends on the PDF reader exposing NVDA-compatible table information; no OCR is included.

Excel's first row in the active contiguous data region is treated as a heading row. Very large tables may take time to read or copy. The bundled DLLs are Windows binaries for x86 and x64.

## Source and translations

`globalPlugins/tabloGezimi/` contains the Python add-on implementation, `locale/` contains `.po` and `.mo` translations, and `doc/` contains localized HTML help. The `TabloOkuma.dll` binaries are included in `globalPlugins/tabloGezimi/x86/` and `x64/`.

The original native C source, export definition, and build script are now included in [`native/`](native/). The supplied x86/x64 binaries have been verified byte-for-byte against the native author’s separate source archive. **The source has not been independently rebuilt into byte-identical binaries.** The author has approved GNU GPL version 2 only (`GPL-2.0-only`) for this project, including the native DLL sources and distributed binaries. See [LICENSE](LICENSE) and [COPYRIGHT.md](COPYRIGHT.md).

## Maintainer

Emrah Bölükbaşı · [GitHub profile](https://github.com/emrah-bolukbasi)

## License

Copyright (C) 2026 Emrah Bölükbaşı. Released under [GNU GPL v2 only](LICENSE) (`GPL-2.0-only`), including the supplied native DLL source and distributed DLLs. See [COPYRIGHT.md](COPYRIGHT.md).
