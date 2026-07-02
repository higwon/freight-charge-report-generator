# Freight Charge Report Generator

Python desktop application for generating formatted freight charge summary reports from forwarding charge Excel files.

## Features

- Finds the source worksheet by required column headers in `.xls`, `.xlsx`, and Excel 2003 XML Spreadsheet files.
- Creates `Source` with all detected source worksheet values.
- Creates `ALL` plus `AE`, `AI`, `OA`, `OE`, and `OI` summary worksheets.
- Groups summaries by Month, Port, and Customer Name.
- Adds Excel outline controls for expanding and collapsing Month and Port details.
- Applies the port selection rules from the specification.
- Applies the built-in pivot-style report formatting without requiring a sample workbook.
- Provides a PySide6 desktop UI with detailed progress, result summary, dialogs, output-file opening, and output-folder opening.
- Logs conversion activity to `logs/app.log`.

## Install

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
python main.py
```

In the application:

1. Select the source Excel file.
2. Select the output folder.
3. Enter an output filename such as `result.xlsx`.
4. Click `Generate`.

## Release Build

GitHub Actions builds a Windows zip artifact on pushes to `main`, pull requests, and manual workflow dispatches.

To publish a release, push a version tag:

```powershell
git tag v0.2.0
git push origin v0.2.0
```

Tags matching `v*` create a GitHub Release and attach the Windows zip package.

## Output

The generated workbook contains:

- `Source`: every row and column from the automatically detected source worksheet.
- `ALL`: the combined total for AE, AI, OA, OE, and OI.
- One sheet per supported Func Code: `AE`, `AI`, `OA`, `OE`, `OI`.

Each summary sheet follows this expandable hierarchy:

```text
Month
    Port
        Customer Name
```

Amounts are summed from `Loc Amt` and formatted as `#,##0.##`, so decimals appear only when needed.

## Project Structure

```text
main.py
requirements.txt
README.md
src/
  app.py
  config.py
  reader.py
  transformer.py
  writer.py
  formatter.py
  models.py
  exceptions.py
  utils.py
ui/
  main_window.py
  worker.py
logs/
output/
```
