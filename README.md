# Layout Validator

A desktop review tool for instrumentation and control engineering drawings.

Compares instrument tags and junction box references extracted from PDF layout instrument lists against an Excel instrument index. Generates review workbooks and marked PDF copies.

## Features

- Batch processing of PDF files and PDF collections in ZIP archives.
- Extraction of supported horizontal and rotated instrument-list tables.
- Checks for obsolete tags, tags absent from the index, and inconsistent junction box assignments.
- Red PDF markings for obsolete tags; blue markings for unrecognized tags.
- Separate tag-review and junction-box-review Excel reports.
- Coverage warnings and lists of failed or unrecognized drawings.
- Revision-numbered output paths preserve previous runs.
- Light and dark themes; optional drag-and-drop support.

## Run

Python 3.10+ with Tk support. Designed for Windows.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe tools\create_samples.py
.\.venv\Scripts\python.exe lee.py
```

Select `samples/instrument_index.xlsx`, add `samples/synthetic_layout.pdf`, choose an output folder, and run the review.

Drag and drop is optional. To enable it, install `tkinterdnd2`. File-selection buttons work without that package.

## Input requirements

The active Excel worksheet is scanned for `Old_Instrument_Name`, `Instrument_Name`, and `JUNCTION_BOX` headers in the first five rows. If not found, columns L, M, and BU are used respectively.

PDFs must contain extractable text. Supported layouts include instrument tables with `TAG NUMBER`, `ELEVATION`, and `J/B NO.` or supported alternative headers. Image-only scanned PDFs require OCR outside this application. Unsupported tables may not be extracted completely.

## Review behavior

- A current instrument-name match takes precedence over an old-name match.
- A tag found only among old names is reported for deletion review.
- A tag absent from both name sets is reported for engineering review, not automatically declared a design error.
- Junction box values are normalized before comparison. If any occurrence of a tag in a drawing matches the expected junction box, it is accepted under the existing review rule.
- Some tag-only table formats skip junction box checks by design.
- Original PDFs are preserved. Markings are written to separate PDF copies.
- Review results support engineer decisions and do not certify drawing completeness or correctness.

## Synthetic example

`tools/create_samples.py` generates a fictional index and layout containing a valid tag, an obsolete tag, an unknown tag, and a junction box mismatch. Original project drawings and workbooks are not included.

## Verification

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

The test exercises PDF extraction, tag/JB comparison, Excel report generation, and marked PDF creation with synthetic input. It does not validate every supported drawing format or the desktop UI.

## Source and distribution

This repository contains the `lee.py` desktop application, setup instructions, and a synthetic sample generator. No source project PDFs, workbooks, or photographs are included. No redistribution license is granted with this repository. Review dependency license requirements before redistributing packaged executables.
