# Changelog

## 3.0.0

Big update. Greek interface (🌐 button top right; follows Windows' language on first start). Double-click a result opens Excel at that exact sheet and cell. Row preview under the results: the whole row with its column headers, the found cell marked. Word (.docx) and PDF files are searched too. New "Greek = Latin" option: skyrodema finds σκυρόδεμα, beton finds μπετόν, and a code typed with the Greek keyboard (Β500C) finds B500C. New "Similar (typos)" match with a % slider: Portlnd finds portland, B500 C finds B500C. Only these columns / Only these sheets filters (letters, ranges or header names). Saved term lists, and import terms from an Excel column or a text file. Recent folders, and added folders are watched: new files join the list by themselves. Faster: all processor cores, and unchanged files are remembered between searches. Excel report: checklist found / not found, all results, one sheet per term. Not-found terms can be searched again as Similar with one click. The update installs RapidFuzz and PyMuPDF by itself on the first start.

## 2.1.0

Updates over the internet: at every start (in the background) and with About → Check for updates, the app looks for a newer version on GitHub, shows what is new and, with your OK, saves it next to this file as Excel_Ultimate_Search_v<version>.pyw and opens it in place of this window. A switch in About turns the start check off. First start installs PyQt6, openpyxl and xlrd by itself behind a splash screen, and makes Desktop and Start menu shortcuts. No console window.

## 2.0.0

Python / PyQt6 rewrite of the JavaFX Excel Ultimate Search: dark and light theme, drag and drop of files and folders, .xlsx/.xlsm/.xls/.csv, match modes (contains, whole cell, starts with, whole word, regex), ignore Greek accents, highlighted matches with the cell address, terms not found, filter, export to CSV and Excel.
