# DataPulse – Intelligent Ticket & Data Processing

DataPulse is a Streamlit application that processes multiple Excel and CSV files, automatically detects ticket columns, and extracts records inactive for **more than 5 calendar days** (`Days Difference > 5`).

---

## 🚀 Features

- **Multi-File & Multi-Sheet**: Ingests multiple `.xlsx` and `.csv` files and processes worksheets independently.
- **Zero-Config Column Detection**: Automatically detects Name, Last Updated Date, Incident number, SCTask number, and Short Description.
- **Separated Ticket Sections**: Groups records into distinct **Incidents** and **SCTasks** tables in both the UI and exports.
- **Modern Dark UI**: Features glassmorphism styling, `streamlit-shadcn-ui`, and 6 color-coded KPI cards (Files, Sheets, Examined, Matching, Skipped, and Issues).
- **Dual Export Formats**:
  - **Excel (`.xlsx`)**: Formatted multi-table workbook with auto-fit widths and formula injection protection.
  - **Rich Text (`.rtf`)**: Native Word/WordPad document tables that eliminate plain-text font and wrapping misalignments.
- **Interactive Previews**: Modal dialogs to inspect both Excel and RTF outputs before downloading.
- **100% In-Memory**: Zero-disk storage for maximum data privacy.

---

## 🛠️ Tech Stack

- **Python**: 3.11+
- **Frontend**: Streamlit, `streamlit-shadcn-ui`
- **Data & Export**: Pandas, OpenPyXL, RTF generator
- **Parsing & Matching**: RapidFuzz, `python-dateutil`, `pytz`
- **Testing**: Pytest (55 tests)

---

## ⚙️ Quick Start

```bash
# Clone repository
git clone <repository-url>
cd cog_proj_2

# Setup virtual environment
python -m venv .venv
.venv\Scripts\activate   # Windows
# source .venv/bin/activate  # macOS/Linux

# Install dependencies & run
pip install -r requirements.txt
streamlit run app.py
```
Open **`http://localhost:8501`** in your browser.

---

## 📖 Usage

1. **Upload & Filter**: Drop Excel/CSV files, set threshold (default: 5 days), and click **Process Records**.
2. **Review Summary**: View KPI metrics and review separated Incident and SCTask tables.
3. **Download**: Click **👁️ View Output** to preview, then download **`.xlsx`** or **`.rtf`** reports.

---

## 📅 Filtering Logic

$$\text{Days Difference} = \text{Reference Date} - \text{Last Updated Date}$$

- **$> 5$ days**: Included (matching).
- **$\le 5$ days**: Excluded (skipped).
- **Future / Invalid dates**: Logged as errors with row and sheet traceability.

---

## 🧪 Testing

Run the 55 automated tests:
```bash
python -m pytest tests/ -v
```
