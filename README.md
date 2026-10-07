# Marksheet Verifier (web version)

The browser version of Marksheet Verifier v2. It does the same two jobs as the Python app:

* **Excel vs Marksheet** (Setup, Results, History tabs): pairs each PDF grade sheet with its Excel/CSV row by
  registration number (ABC ID when the PDF has none) and compares name, father's name, mother's name, date of birth
  and ABC ID.
* **Meritto vs Marksheet**: checks that each marksheet now shows the value the Meritto/TR record says it should.

All files are read **inside your browser**. Nothing is uploaded, and no external service is used for matching.

## Put it on GitHub Pages

1. Create a GitHub repository and upload everything in this folder (keep the folder structure; `node_modules` is not
   needed).
2. In the repository: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, choose `main` and
   `/ (root)`, then **Save**.
3. After a minute the site is live at `https://<your-user>.github.io/<repository>/`.

The `.nojekyll` file tells GitHub to serve the files as they are.

### Run it on your own computer

Browsers block the PDF reader when a page is opened straight from disk (`file://`), so start a small local server:

    python -m http.server 8000        # then open http://localhost:8000

## How to use it

Same workflow as the desktop app:

1. **Setup**: choose the Excel file and add the PDFs. You can use the files button, a whole folder, or drag and drop.
   Columns are auto-detected; change any dropdown to override. Double-click a PDF row to correct what was read.
   Then click **Run comparison**.
2. **Results**: filter, double-click a row for field-by-field detail, **Export Excel report**.
3. **History**: every run is saved in this browser. Re-open, re-export or delete runs, or look up one
   registration number across all runs.
4. **Meritto vs Marksheet**: choose the Meritto file, add the PDFs, check the columns, then **Run check**.
   Double-click a result to adjust it by hand. **Re-check** keeps manual changes.

Matching rules, thresholds, strict/normal mode, statuses and the Excel reports are the same as in the Python version
(see the original README for what each result means).

## Differences from the desktop version

| Desktop (Python) | Web | Why |
|---|---|---|
| OCR for scanned PDFs (Tesseract) | Not available | Needs a program installed on the computer. Scanned PDFs are flagged; type their values by hand (double-click the row). |
| History in a SQLite file; "Choose database file" | History is stored in this browser (IndexedDB) | A web page cannot write files on disk. History is per browser and per computer; clearing site data deletes it. |
| Report saved via a "Save as" dialog | Report is downloaded | Same file names: `verification_report_YYYYMMDD_HHMM.xlsx`, `verification_run_N.xlsx`, `meritto_report_YYYYMMDD_HHMM.xlsx`. |
| Command line (`main.py cli/meritto/history`) | Not included | The Python version still works for scripted use. |

### Known limits

* **Text printed on top of other text in a PDF.** For example, a very long name that runs over the next column's
  label. Both versions return a garbled value, but not the same garbled value, because pdf.js and pdfplumber place
  individual letters slightly differently. Either way the field is reported as a mismatch, never as a match.
* **Whole numbers longer than 15 digits stored as numbers inside an .xlsx written by non-Excel software.** The browser
  reads them as floating point values. Excel itself never stores more than 15 digits, and registration/ABC numbers
  are 12 digits, so normal files are unaffected.
* The page loads the Atkinson Hyperlegible font from Google Fonts. No data is sent with that request. If it is
  blocked, the page uses the system font. To remove it, delete the three `fonts.g…` lines in `index.html`.

## Files

    index.html            the page
    css/styles.css        styling (colours match the Excel report)
    js/core.js            port of core.py + meritto.py: extraction rules, auto-detection, comparison
    js/py.js              Python behaviour JavaScript lacks (round() ties, float printing, \s \d \w \b, sorting)
    js/pdfread.js         read_pdf(): pdf.js in place of pdfplumber
    js/excel.js           reading .xlsx/.xls/.ods/.csv like openpyxl/pandas (dates, numbers, CSV sniffing)
    js/report.js          port of report.py (ExcelJS)
    js/storage.js         port of storage.py (IndexedDB)
    js/app.js             the four tabs (port of gui.py and gui_meritto.py)
    vendor/               pdf.js 4.10.38, SheetJS 0.18.5, ExcelJS 4.4.0 (with their licences)
    tests/                port of tests/test_verifier.py, with fixture files

## Tests

    npm install
    npm test                                   # 32 pass, 4 skipped (real-file tests)
    MV_REAL_DATA=/path/to/real/files npm test  # also runs the real-file tests

## How the port was checked against the Python app

The Python app was the reference. The same inputs were run through both versions and the outputs were compared:

* Scoring, date parsing, number normalising, label detection and line parsing: 62,015 randomized comparisons,
  0 differences.
* PDF field extraction: 21 PDFs (including a real Medhavi grade sheet). 20 are identical; the exception is the
  overlapping-text case above.
* Excel/CSV loading: 18 files (real Learner and Meritto files, multi-sheet "All sheets", `.xls`, dates, times,
  CSV edge cases and their error messages). All 18 are identical.
* Complete runs on the real files with several settings: 43 runs (470 comparison records, 1,086 Meritto checks).
  Every status, score, note and column-detection result is identical.
* Excel reports: 11 reports compared cell by cell, including values, formulas, colours, borders, widths, merges,
  filters and dropdowns. All identical, and the formulas give the same totals when recalculated.
* History: saving and reopening a run gives the same result as the SQLite version.
* The page itself, in Chrome: uploads, runs, filters, manual override, history and downloads. The downloaded
  reports are identical to the Python reports for the same run.

One original Python test (`test_real_truscholar_file`) fails on the current Learner file because it expects a
"Mapping" tab that this version of the file does not have. Its column-detection checks pass in both versions.
