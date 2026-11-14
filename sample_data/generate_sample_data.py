"""Script to generate realistic sample test data files for DataPulse."""
import io
from pathlib import Path
import pandas as pd

SAMPLE_DIR = Path(__file__).parent


def generate_samples():
    """Create sample XLSX and CSV files with varied structures, headers, and date representations."""
    SAMPLE_DIR.mkdir(parents=True, exist_ok=True)

    # 1. employee_updates_2026.xlsx (Multi-sheet)
    # Sheet 1: Engineering
    df_eng = pd.DataFrame([
        {"Employee_Name": "Michael Scott", "Department": "Engineering", "Last_Updated": "2026-11-07", "Project_Status": "Active"},   # 6 days diff -> MATCH
        {"Employee_Name": "Sarah Connor", "Department": "Engineering", "Last_Updated": "2026-11-01", "Project_Status": "Active"},    # 12 days diff -> MATCH
        {"Employee_Name": "David Miller", "Department": "Engineering", "Last_Updated": "2026-11-08", "Project_Status": "Pending"},   # 5 days diff -> SKIPPED (boundary)
        {"Employee_Name": "Alex Rivera", "Department": "Engineering", "Last_Updated": "2026-11-10", "Project_Status": "Active"},    # 3 days diff -> SKIPPED
        {"Employee_Name": "John Doe", "Department": "Engineering", "Last_Updated": "2026-11-12", "Project_Status": "Active"},       # 1 day diff -> SKIPPED
        {"Employee_Name": "Invalid Date Person", "Department": "Engineering", "Last_Updated": "not-a-date", "Project_Status": "Hold"}, # Error
        {"Employee_Name": None, "Department": "Engineering", "Last_Updated": "2026-10-15", "Project_Status": "Active"},            # Missing Name
    ])

    # Sheet 2: Marketing (different column names)
    df_mkt = pd.DataFrame([
        {"Resource": "Rachel Green", "Campaign": "Winter Promo", "Modified On": "20-Oct-2026", "Role": "Lead"},    # 24 days diff -> MATCH
        {"Resource": "Monica Geller", "Campaign": "Branding", "Modified On": "05-Nov-2026", "Role": "Manager"},    # 8 days diff -> MATCH
        {"Resource": "Chandler Bing", "Campaign": "Copywriting", "Modified On": "11-Nov-2026", "Role": "Specialist"}, # 2 days diff -> SKIPPED
        {"Resource": "Joey Tribbiani", "Campaign": "Media", "Modified On": "2026-11-15", "Role": "Actor"},         # Future date -> Error
    ])

    xlsx_path = SAMPLE_DIR / "employee_updates_2026.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_eng.to_excel(writer, sheet_name="Engineering", index=False)
        df_mkt.to_excel(writer, sheet_name="Marketing", index=False)
    print(f"Generated: {xlsx_path}")

    # 2. crm_contacts.csv (Comma-delimited)
    df_crm = pd.DataFrame([
        {"Full Name": "Robert Brown", "Account": "Enterprise A", "Last Activity Date": "15/10/2026", "Priority": "High"},    # 29 days diff -> MATCH
        {"Full Name": "Emily Watson", "Account": "Mid-Market B", "Last Activity Date": "02/11/2026", "Priority": "Medium"},  # 11 days diff -> MATCH
        {"Full Name": "Daniel Craig", "Account": "Gov Sector", "Last Activity Date": "08/11/2026", "Priority": "Low"},       # 5 days diff -> SKIPPED
        {"Full Name": "Anna Taylor", "Account": "Retail C", "Last Activity Date": "13/11/2026", "Priority": "High"},        # 0 days diff -> SKIPPED
    ])
    csv_crm_path = SAMPLE_DIR / "crm_contacts.csv"
    df_crm.to_csv(csv_crm_path, index=False)
    print(f"Generated: {csv_crm_path}")

    # 3. contractor_tasks.csv (Semicolon-delimited with ISO timestamps)
    df_contractors = pd.DataFrame([
        {"Assigned To": "Bruce Wayne", "Task Title": "Security Audit", "Modification Timestamp": "2026-10-01T09:30:00Z", "Hourly Rate": 150},  # 43 days -> MATCH
        {"Assigned To": "Clark Kent", "Task Title": "Press Release", "Modification Timestamp": "2026-11-06T18:00:00Z", "Hourly Rate": 85},     # 7 days -> MATCH
        {"Assigned To": "Diana Prince", "Task Title": "Diplomacy Brief", "Modification Timestamp": "2026-11-11T12:00:00Z", "Hourly Rate": 120}, # 2 days -> SKIPPED
    ])
    csv_contractor_path = SAMPLE_DIR / "contractor_tasks.csv"
    df_contractors.to_csv(csv_contractor_path, sep=";", index=False)
    print(f"Generated: {csv_contractor_path}")


if __name__ == "__main__":
    generate_samples()
