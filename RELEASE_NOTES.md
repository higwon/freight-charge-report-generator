# Release Notes

## v0.2.0

- Detects the source worksheet by required column headers instead of a fixed sheet name.
- Renames `RawData` to `Source` and adds an `ALL` summary sheet.
- Adds expandable Month and Port outline levels to summary sheets.
- Adds application branding, version/author metadata, and success/failure result images.
- Refreshes the UI with clearer progress stages and Korean action labels.

## v0.1.0

- Initial desktop app for generating freight charge summary reports from `LV1DOC`.
- Supports Excel 2003 XML Spreadsheet, `.xls`, and `.xlsx` source files.
- Generates `RawData` plus one summary sheet per Func Code.
- Applies built-in pivot-style formatting and decimal display only when needed.
- Includes a PySide6 UI with detailed progress, result summary, result-file opening, and output-folder opening.
