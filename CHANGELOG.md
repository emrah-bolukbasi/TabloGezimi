# Changelog

## 1.1.11 — Cumulative update after 1.1.10 (prepared, not yet published)

### Reading and navigation
- Added a configurable speech pause between cells during full-row reading (150 ms default; 0–5000 ms).
- Ctrl+Shift+Alt+Down Arrow opens the Turkish pause duration entry dialog. The setting is saved and the shortcut is available in NVDA Input Gestures.
- Added a second navigation mode for accessible browser, PDF and desktop Excel tables. Ctrl+Shift+Alt+Up Arrow switches between row-reading mode and single-cell mode. New table-navigation sessions start in mode 1.
- In Excel, arrow-key navigation can go beyond the current data region into other spreadsheet cells.
- When navigating to an empty Excel cell, its address (for example A5) is announced; bulk row/column/table reading and copying still skip empty values.
- In Excel, copying is limited to the active contiguous data region; fully blank rows/columns separate regions. Real Excel table objects are considered during destination reading.

### Settings and preservation
- The open/closed state of table navigation (Alt+Shift+O) is saved and restored after NVDA restarts.
- Existing row, column, cell and table clipboard shortcuts remain available in both navigation modes.
- Retained the included native DLL binaries and nine language resource sets.

### Known limitations
- Accessibility information is needed for browser/PDF tables; scanned images are not processed by OCR.
- Multiple adjacent data blocks without a fully blank separating row or column are not always distinguishable.
- Runtime testing on a Windows machine with actual NVDA, Excel, browsers and accessible PDFs is still required before publication.

## 1.1.10 — Previously published
- Accessible table navigation in Excel, web browsers, PDF/document readers.
- Row reading with headings, cell/row/column/table copying, nine languages.
- Retired the unused Alt+Shift+R command.
