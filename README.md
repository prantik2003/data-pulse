# DataPulse – Intelligent Excel & CSV Processing

DataPulse is a Python-based Streamlit application that processes multiple Excel and CSV files, automatically detects Name and Last Updated Date columns, and identifies outdated records.

**Primary Rule:** Records are included only when they have not been updated for more than 5 calendar days.

## 🚀 Features

- Upload multiple Excel (`.xlsx`) and CSV (`.csv`) files.
- Process multiple worksheets independently.
- Automatically detect Name and Date columns using fuzzy matching and content validation.
- Support multiple date formats, Excel serial dates, and timezone conversion.
- Customize column mappings, date preferences, timezone, and aging threshold.
- View matching records, skipped records, and processing errors.
- Export filtered records to Excel and TXT files.
- Preview reports before downloading.
- Secure in-memory file processing.

## 🛠️ Tech Stack

- Python 3.11+
- Streamlit
- Pandas
- OpenPyXL
- RapidFuzz
- python-dateutil
- Pytest

## 📁 Project Structure

text

```text
smart-data-filter/
├── app.py
├── config/
├── modules/
├── utils/
├── tests/
├── sample_data/
├── .streamlit/
├── requirements.txt
└── README.md
```

## ⚙️ Installation

bash

```bash
git clone <repository-url>
cd smart-data-filter

python -m venv .venv
```

Activate the virtual environment:

bash

```bash
# macOS/Linux
source .venv/bin/activate

# Windows
.venv\Scripts\activate
```

Install dependencies:

bash

```bash
pip install -r requirements.txt
```

## ▶️ Run Application

bash

```bash
streamlit run app.py
```

Open `http://localhost:8501` in your browser.

## 📖 Usage

1. Upload Excel or CSV files.
2. Review and configure detected column mappings.
3. Set the date threshold and other preferences.
4. Click **Process Records**.
5. Review results and download Excel or TXT reports.

## 📅 Filtering Logic

python

```python
Days Difference = Reference Date - Last Updated Date

Include if Days Difference > 5
```

- 0–5 days: Excluded.
- 6+ days: Included.
- Future dates: Logged as errors.

## 🧪 Testing

The project includes 51 automated tests covering column detection, date parsing, filtering, file processing, and report generation.

bash

```bash
python -m pytest tests/ -v
```
