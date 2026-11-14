# DataPulse – Intelligent Excel & CSV Processing Web Application

[![Python Version](https://img.shields.io/badge/python-3.11%2B-blue.svg)](https://www.python.org/)
[![Framework](https://img.shields.io/badge/Streamlit-1.42%2B-FF4B4B.svg)](https://streamlit.io/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/tests-51%20passed-brightgreen.svg)]()

**DataPulse** is a production-quality web application built with Python and Streamlit. It allows users to upload heterogeneous Excel (`.xlsx`) and CSV (`.csv`) files containing arbitrary columns, intelligently identifies relevant **Name** and **Last Updated Date** columns, normalizes diverse date representations, and filters records based on calendar-day aging thresholds.

---

## 1. Project Overview

Organizations often handle periodic data exports (CRM rosters, task trackers, audit logs, employee sheets) generated from different systems. These files typically:

- Use inconsistent column headers (e.g., `Employee Name` vs. `Resource` vs. `Assigned To`).
- Represent timestamps in multiple formats (e.g., ISO, UK dates, US dates, Excel serials).
- Need consolidation to identify outdated records requiring follow-up.

DataPulse automates this extraction and aging analysis:

- **Primary Filtering Rule:** Retain only records whose last updated date is **strictly more than 5 calendar days older** than the current reference date (`Days Difference > 5`).
- **Boundary Precision:** Records updated exactly 5 days ago (`Difference == 5`) or fewer are strictly excluded.
- **Traceability:** Consolidated exports retain the original source file and worksheet for every matching record.

---

## 2. Key Features

- **Multi-File & Multi-Format Ingestion:**
  - Drag-and-drop support for multiple `.xlsx` and `.csv` files simultaneously.
  - Independent processing of multiple worksheets within Excel workbooks.
  - Automatic detection of CSV delimiters (`,`, `;`, `\t`, `|`) and encodings (`utf-8`, `utf-8-sig`, `latin-1`, `cp1252`).
- **Deterministic Intelligent Column Detection:**
  - Semantic normalization handling CamelCase, underscores, hyphens, and whitespace variations.
  - Fuzzy header matching powered by RapidFuzz token analysis against comprehensive domain dictionaries.
  - Deep content validation: samples row values to verify string types for names and date parseability for dates.
  - Numeric ID false-positive prevention: prevents employee IDs or salaries from being misinterpreted as Excel serial dates.
  - Three-tier confidence scoring (`High`, `Medium`, `Low`) with manual override dropdowns for complete user control.
- **Robust Date Parsing & Normalization:**
  - Parses native Python `date`/`datetime`, Pandas `Timestamp`, and Excel serial integers/floats (e.g., `46331` $\to$ `2026-11-05`).
  - Supports ISO 8601, timezone-aware strings, `DD/MM/YYYY`, `MM/DD/YYYY`, `YYYY-MM-DD`, `DD-MMM-YYYY`, and `MMM DD, YYYY`.
  - Column-context inference and user preference settings to resolve ambiguous numeric dates like `05/11/2026`.
  - Timezone normalization (default: `Asia/Kolkata`) with calendar-day calculation independent of time-of-day offsets.
- **Dual Export Generation & Modal Previews:**
  - **Excel (`filtered_records.xlsx`):** Clean, professional spreadsheet containing the filtered matching records (with bold headers, freeze panes, auto-fit column widths, `DD-MMM-YYYY` date formatting, and Excel auto-filters).
  - **Text (`filtered_records.txt`):** Formatted plain text summary suitable for scripting and quick inspection.
  - **Interactive Modal Previews:** One-click modal viewers for both Excel and TXT reports to inspect the exact output before downloading.
- **Security & Safety:**
  - Fully in-memory processing (`io.BytesIO`) without persisting uploaded data to disk.
  - CSV/Excel formula injection prevention: sanitizes cells starting with `=`, `+`, `-`, `@`.
  - Future dates and malformed records are safely excluded from matching output and logged in the error audit table.

---

## 3. Technology Stack

| Component                      | Technology                  | Purpose                                             |
| ------------------------------ | --------------------------- | --------------------------------------------------- |
| **Programming Language**       | Python 3.11+                | Modern, typed Python core                           |
| **Web UI Framework**           | Streamlit                   | Reactive dashboard layout & state management        |
| **Tabular Data Processing**    | Pandas                      | High-performance DataFrame operations               |
| **Excel Reading & Styling**    | OpenPyXL                    | Multi-sheet reading, workbook styling, freeze panes |
| **Fuzzy Matching**             | RapidFuzz                   | Fast Levenshtein & token-set string matching        |
| **Date Parsing**               | python-dateutil + datetime  | Robust date normalization & parsing                 |
| **Timezone Management**        | zoneinfo (Standard Library) | Standard IANA timezone conversions                  |
| **Unit & Integration Testing** | Pytest                      | Test suite with 51 comprehensive tests              |

---

## 4. Project Structure

```
smart-data-filter/
├── app.py                         # Streamlit application UI and state coordination
├── config/
│   ├── __init__.py
│   └── settings.py                # Central application configuration & keyword dictionaries
├── modules/
│   ├── __init__.py
│   ├── column_detector.py         # Fuzzy matching & content-based column classification
│   ├── data_validator.py          # Field-level validation for names and dates
│   ├── date_parser.py             # Multi-format date parsing & Excel serial handling
│   ├── export_generator.py        # Styled OpenPyXL workbook & TXT report generation
│   ├── file_reader.py             # XLSX & CSV ingestion with encoding/delimiter detection
│   ├── filter_engine.py           # Strict calendar-day aging calculation & filtering
│   └── processing_manager.py      # Pipeline coordinator for multi-file/sheet workflows
├── utils/
│   ├── __init__.py
│   ├── date_utils.py              # Timezone conversions and calendar difference utilities
│   ├── error_handler.py           # Structured error categories & diagnostic logging
│   ├── file_utils.py              # Formula injection sanitization & safe filenames
│   └── header_normalizer.py       # CamelCase, punctuation, and case normalization
├── tests/
│   ├── __init__.py
│   ├── test_column_detection.py   # Unit tests for normalization, fuzzy ranking, content validation
│   ├── test_date_parsing.py       # Unit tests for 17 date format scenarios & leap years
│   ├── test_export_generation.py  # Unit tests for Excel tabs, styling, formula safety, TXT formats
│   ├── test_file_processing.py    # Integration tests for single/multi XLSX/CSV files
│   └── test_filtering.py          # Unit tests for calendar differences, boundaries, and duplicates
├── sample_data/
│   ├── README.md                  # Documentation of test sample files
│   ├── generate_sample_data.py    # Generator script for sample test datasets
│   ├── employee_updates_2026.xlsx # Multi-sheet Excel test workbook
│   ├── crm_contacts.csv           # Comma-delimited CSV test file
│   └── contractor_tasks.csv       # Semicolon-delimited CSV test file
├── .streamlit/
│   └── config.toml                # Streamlit UI theme and server limits
├── requirements.txt               # Pinned production dependencies
├── .gitignore                     # Git ignore rules
└── README.md                      # Comprehensive project documentation
```

---

## 5. Installation & Setup

### Prerequisites

- Python 3.11 or higher
- Git

### 1. Clone the Repository

```bash
git clone <repository-url>
cd smart-data-filter
```

### 2. Create and Activate a Virtual Environment

```bash
# macOS/Linux
python3 -m venv .venv
source .venv/bin/activate

# Windows (Command Prompt)
python -m venv .venv
.venv\Scripts\activate

# Windows (PowerShell)
python -m venv .venv
.venv\Scripts\Activate.ps1
```

### 3. Install Dependencies

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

---

## 6. Running Locally

Launch the application using Streamlit:

```bash
streamlit run app.py
```

Open your web browser at:

```
http://localhost:8501
```

---

## 7. How to Use the Application

### Step 1: Upload Files

1. Drag and drop any number of `.xlsx` or `.csv` files into the uploader box in **Tab 1**.
2. The dashboard displays the detected file type, file size, worksheet count, and row count.

### Step 2: Review Column Mapping & Configure Parameters

1. Open **Tab 2: Column Mapping & Configuration**.
2. For each worksheet, inspect the automatically detected:
   - **Name Column** (with confidence badge: `HIGH`, `MEDIUM`, or `LOW`).
   - **Last Updated Date Column** (with confidence badge).
3. If needed, override column selections using the dropdown menus.
4. Customize parameters in the sidebar:
   - **Threshold (Calendar Days):** Default is `5` days.
   - **Date Format Preference:** Choose `Auto-detect`, `DD/MM/YYYY`, `MM/DD/YYYY`, or `YYYY-MM-DD`.
   - **Application Timezone:** Default is `Asia/Kolkata`.
   - **Reference Date:** Defaults to today's date in the chosen timezone.
5. Click **"🚀 Process Records"**.

### Step 3: Inspect Processing Results

1. View the KPI metrics cards:
   - _Total Files Processed_, _Total Worksheets_, _Total Records Examined_, _Matching Records (> Threshold)_, _Skipped Records (≤ Threshold)_, and _Processing Errors_.
2. Browse the interactive sub-tabs:
   - **Matching Records:** Full table with search filter by Name or File.
   - **Skipped Records:** Records updated within the threshold window.
   - **Processing Errors:** Records with missing names, invalid dates, or future dates.
   - **File-wise Breakdown:** Summary counts per file and sheet.

### Step 4: Preview & Download Reports

1. Navigate to **Tab 4: Download Reports**.
2. Click **"👁️ View Output"** to preview either the Excel spreadsheet or the Text report in a popup modal.
3. Click **"📥 Download Excel"** to download `filtered_records.xlsx` (clean spreadsheet of matching records with `DD-MMM-YYYY` date formatting).
4. Click **"📥 Download TXT"** to download `filtered_records.txt`.

---

## 8. How Column Detection Works

The detection engine combines semantic header analysis with content validation:

1. **Header Normalization:**
   Converts headers like `Employee_Name`, `EMPLOYEE NAME`, `employeeName`, and `Employee-Name` into the standard form `employee name`.
2. **Dictionary & Fuzzy Matching:**
   Compares normalized headers against curated keyword dictionaries using RapidFuzz ratio, token sort, and token set metrics.
3. **Content-Based Verification:**
   - **Date Candidates:** Evaluates a sample of cell values to check what fraction can be parsed as valid dates. If a column contains 0% parseable dates, its score drops to zero even if the header contains the word "date".
   - **Numeric ID Safeguard:** Raw numbers in columns with non-date headers (e.g., `Salary`, `Amount`, `Employee ID`) are excluded from serial date parsing.
   - **Name Candidates:** Checks that cell values are predominantly text, non-numeric, non-date, and within typical human name lengths (rejecting paragraph comments or ticket descriptions).
4. **Confidence Scoring:**
   - **High ($\ge 0.85$):** Auto-suggested and flagged with green badge.
   - **Medium ($0.60 - 0.84$):** Suggested and flagged for user verification.
   - **Low ($< 0.60$):** Flagged for manual selection.
   - **Close Top Candidates:** If top two candidate scores differ by $\le 0.05$, confirmation is requested.

---

## 9. Supported Date Formats

| Format            | Example                | Resolution Strategy                                   |
| ----------------- | ---------------------- | ----------------------------------------------------- |
| `DD/MM/YYYY`      | `25/11/2026`           | Unambiguous when day $> 12$, or via format preference |
| `MM/DD/YYYY`      | `11/25/2026`           | Unambiguous when day $> 12$, or via format preference |
| `YYYY-MM-DD`      | `2026-11-05`           | ISO standard (unambiguous)                            |
| `DD-MM-YYYY`      | `05-11-2026`           | Hyphenated calendar date                              |
| `DD.MM.YYYY`      | `05.11.2026`           | Dotted European calendar date                         |
| `DD-MMM-YYYY`     | `05-Nov-2026`          | Textual month (unambiguous)                           |
| `MMM DD, YYYY`    | `Nov 05, 2026`         | Textual month (unambiguous)                           |
| `YYYY/MM/DD`      | `2026/11/05`           | Slash date standard                                   |
| ISO Timestamp     | `2026-11-05T14:30:00`  | ISO 8601 string                                       |
| Timestamp with TZ | `2026-11-05T14:30:00Z` | Converted to configured application timezone          |
| Excel Serial Date | `46331`                | Converted to `2026-11-05` (Excel 1899-12-30 epoch)    |
| Native Datetime   | `datetime.datetime`    | Normalized to `datetime.date`                         |

---

## 10. How Filtering Works

- **Formula:**
  $$\text{Days Difference} = \text{Current Reference Date} - \text{Last Updated Date}$$
- **Filtering Condition:**
  $$\text{Days Difference} > \text{Threshold}$$

### Example Table (Assuming Reference Date = 13-Nov-2026, Threshold = 5):

| Name    | Last Updated | Days Difference |   Action    | Reason                         |
| ------- | ------------ | --------------: | :---------: | ------------------------------ |
| John    | 12-Nov-2026  |               1 | **Exclude** | $\le 5$ days                   |
| Alex    | 10-Nov-2026  |               3 | **Exclude** | $\le 5$ days                   |
| David   | 08-Nov-2026  |               5 | **Exclude** | Exactly 5 days (Boundary)      |
| Michael | 07-Nov-2026  |               6 | **Include** | Strictly $> 5$ days            |
| Sarah   | 01-Nov-2026  |              12 | **Include** | Strictly $> 5$ days            |
| Robert  | 20-Oct-2026  |              24 | **Include** | Strictly $> 5$ days            |
| Joey    | 15-Nov-2026  |              -2 | **Flagged** | Future date (logged in Errors) |

---

## 11. Testing

The project includes an automated test suite with **51 tests** covering column detection, date parsing, calendar-day filtering, and end-to-end integration.

To run the complete test suite:

```bash
python -m pytest tests/ -v
```

### Test Suite Structure:

- `tests/test_column_detection.py`: 15 tests (exact match, CamelCase, underscore/hyphen, fuzzy matching, separate datasets, multiple candidates, missing columns, content rejection).
- `tests/test_date_parsing.py`: 17 tests (all formats, timezone shifts, Excel serials, leap years, ambiguous format preferences, empty/invalid strings).
- `tests/test_filtering.py`: 10 tests (strict boundary exclusions, inclusion of 6+ days, future date flags, duplicate name preservation, midnight calendar-day shifts).
- `tests/test_file_processing.py`: 6 integration tests (XLSX, CSV, multi-sheet independence, mixed formats, zero matches, corrupted files).
- `tests/test_export_generation.py`: 3 tests (Excel tabs, OpenPyXL styling, formula injection protection, TXT report formatting).

---

## 12. Deployment Instructions (Streamlit Community Cloud)

To deploy on [Streamlit Community Cloud](https://streamlit.io/cloud):

1. **Push to GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit: DataPulse application"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo-name>.git
   git push -u origin main
   ```
2. **Deploy on Streamlit Cloud:**
   - Log into [share.streamlit.io](https://share.streamlit.io/).
   - Click **"New app"**.
   - Select your repository, branch (`main`), and set **Main file path** to `app.py`.
   - Click **"Deploy!"**.
3. **Environment & Resources:**
   - The app runs on standard Python 3.11 with `requirements.txt` and `.streamlit/config.toml` included in the root directory.

---

## 13. Privacy and File Handling Considerations

- **In-Memory Operations:** Files are read directly from upload streams (`io.BytesIO`) and are not written to permanent storage.
- **Session Isolation:** Streamlit session state is isolated per user connection. When a user closes or resets their session, data is immediately eligible for Python garbage collection.
- **Formula Injection Mitigation:** Any text cells beginning with characters that could trigger spreadsheet formula execution (`=`, `+`, `-`, `@`, `\t`, `\r`) are automatically escaped with a leading single quote (`'`).
- **Data Protection:** No uploaded data or record details are transmitted to external APIs, third-party LLMs, or telemetry services.

---

## 14. License

This project is licensed under the MIT License.
