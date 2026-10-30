# Excel Ultimate Search

Search many Excel and CSV files for many terms at once. One file, `Excel_Ultimate_Search.pyw`,
that opens without a console window.

## Run it

1. Install [Python 3.9+](https://www.python.org/downloads/) for Windows (tick "Add python.exe to PATH").
2. Download `Excel_Ultimate_Search.pyw` and double-click it. On the first start it installs what it
   needs (PyQt6, openpyxl, xlrd) behind a splash screen and makes Desktop and Start menu shortcuts.
3. Drop files or folders on the left, paste your terms (one per line), press **Search**.

## What it does

- Reads `.xlsx`, `.xlsm`, `.xls` and `.csv`, every sheet, text, numbers and dates.
- Match modes: contains, whole cell, starts with, whole word, regular expression; match case;
  ignore accents (ά = α).
- Results with the cell (e.g. `B3`) and the matched text highlighted; the terms that were not found;
  files that could not be read.
- Filter, sort, double-click to open the file, export to CSV or Excel (with a per-term summary).

## Updates

At every start the app looks on this repository for a newer version (in the background). When there
is one it shows what is new and, with your OK, saves it next to the old file as
`Excel_Ultimate_Search_v<version>.pyw` and opens it. About → *Check for updates* does the same on
demand; a switch there turns the start check off.

**Releasing:** raise `APP_VERSION` in `Excel_Ultimate_Search.pyw`, add a `## <version>` section at
the top of `CHANGELOG.md`, push to `main`.

---
Nikolaos K. Paridis (@nikkpap) · ALU DEV TEAM
