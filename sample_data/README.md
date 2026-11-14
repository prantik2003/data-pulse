# Sample Datasets for DataPulse

This directory contains realistic sample datasets designed to test all capabilities of the **DataPulse** application.

## Files Provided

### 1. `employee_updates_2026.xlsx`

A multi-sheet Excel workbook:

- **Sheet 'Engineering':**
  - Columns: `Employee_Name`, `Department`, `Last_Updated`, `Project_Status`
  - Date format: Standard ISO (`YYYY-MM-DD`) and slash formats (`DD/MM/YYYY`).
  - Contains records with differences ranging from 1 day to 45 days, plus missing values and exact 5-day boundary records.
- **Sheet 'Marketing':**
  - Columns: `Resource`, `Campaign`, `Modified On`, `Role`
  - Demonstrates independent column detection per worksheet (different column headings from Engineering).

### 2. `crm_contacts.csv`

A comma-delimited CSV file:

- Columns: `Full Name`, `Account`, `Last Activity Date`, `Priority`
- Date format: Textual month format (e.g. `07-Nov-2026`, `20-Oct-2026`).

### 3. `contractor_tasks.csv`

A semicolon-delimited CSV file (`sep=;`):

- Columns: `Assigned To;Task Title;Modification Timestamp;Hourly Rate`
- Delimiter: Semicolon `;`
- Date format: ISO timestamp (e.g. `2026-11-05T14:30:00Z`).
- Tests delimiter auto-detection and timestamp timezone normalization.

---

## Regenerating Sample Data

To regenerate the sample files at any time:

```bash
python sample_data/generate_sample_data.py
```
