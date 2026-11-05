# Excel Ultimate Search

Search many Excel, CSV, Word and PDF files for many terms at once. One file,
`Excel_Ultimate_Search.pyw`, that opens without a console window. English and Greek interface.

## Run it

1. Install [Python 3.9+](https://www.python.org/downloads/) for Windows (tick "Add python.exe to PATH").
2. Download `Excel_Ultimate_Search.pyw` and double-click it. On the first start it installs what it
   needs (PyQt6, openpyxl, xlrd, RapidFuzz, PyMuPDF) behind a splash screen and makes Desktop and
   Start menu shortcuts.
3. Follow the numbered steps in the window:
   **1 · Files** (drop files or folders) → **2 · Search terms** (one per line) → **3 · Options** → **4 · Search**.

## What it does

- Reads `.xlsx`, `.xlsm`, `.xls`, `.csv`, `.docx` and `.pdf` (PDFs need real text, not scans).
- Match modes: contains, whole cell, starts with, whole word, **similar (typos)** with a % threshold,
  regular expression; match case; ignore accents (ά = α).
- **Greek = Latin**: Greeklish and wrong-keyboard codes (`skyrodema` → σκυρόδεμα, `beton` → μπετόν,
  `Β500C` typed with a Greek Β → B500C).
- **Only these columns / sheets**: `B, D:F`, header names like `Περιγραφή`, sheet names with `*`.
- Results show the cell (`B3`), the matched text highlighted, and a **row preview** with the column
  headers. Double-click opens **Excel at that sheet and cell**.
- **Term lists**: save / load, import from an Excel column or a text file.
- **Folders**: recent folders, and added folders are watched for new files (F5 refreshes).
- **Fast**: uses all processor cores; unchanged files are remembered between searches.
- **Excel report**: checklist found / not found, all results, one sheet per term. Also CSV.

## Updates

At every start the app looks on this repository for a newer version (in the background). When there
is one it shows what is new and, with your OK, saves it next to the old file as
`Excel_Ultimate_Search_v<version>.pyw` and opens it. About → *Check for updates* does the same on
demand; a switch there turns the start check off.

**Releasing:** raise `APP_VERSION` in `Excel_Ultimate_Search.pyw`, add a `## <version>` section at
the top of `CHANGELOG.md`, push to `main`.

---
Nikolaos K. Paridis (@nikkpap) · ALU DEV TEAM
