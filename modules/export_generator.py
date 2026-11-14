"""Export generation module for Excel (.xlsx) and Text (.txt) reports."""
import datetime
import io
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from modules.filter_engine import FilterResult, Record
from utils.error_handler import ProcessingError
from utils.file_utils import sanitize_cell_for_export


class ExportGenerator:
    """Generates styled Excel (.xlsx) workbooks and plain text (.txt) reports."""

    # Styling constants
    HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")  # Navy / Slate Blue
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="64748B")
    REGULAR_FONT = Font(name="Calibri", size=11)
    BOLD_FONT = Font(name="Calibri", size=11, bold=True)
    METRIC_LABEL_FONT = Font(name="Calibri", size=11, bold=True, color="1E293B")
    METRIC_VAL_FONT = Font(name="Calibri", size=11, color="0F172A")
    ERROR_HEADER_FILL = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")  # Crimson

    THIN_BORDER = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0"),
    )

    @classmethod
    def generate_excel_bytes(
        cls,
        matching_records: List[Record],
        errors: Optional[List[ProcessingError]] = None,
        summary_stats: Optional[Dict[str, Any]] = None,
        date_format: str = "%d-%b-%Y",
    ) -> bytes:
        """Create a formatted Excel workbook in memory containing the filtered records and return as bytes."""
        wb = openpyxl.Workbook()
        ws_matching = wb.active
        ws_matching.title = "Filtered Records"
        cls._build_matching_sheet(ws_matching, matching_records, date_format)

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    @classmethod
    def _build_summary_sheet(cls, ws: Any, stats: Dict[str, Any]):
        """Format and populate the Summary tab with high-level processing KPIs."""
        ws.views.sheetView[0].showGridLines = True

        # Report Header
        ws["A1"] = "DataPulse – Processing Report Summary"
        ws["A1"].font = cls.TITLE_FONT
        ws["A2"] = f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Application Timezone: {stats.get('timezone', 'Asia/Kolkata')}"
        ws["A2"].font = cls.SUBTITLE_FONT

        # Key Metrics Table
        metrics = [
            ("Processing Date", str(stats.get("processing_date", ""))),
            ("Filtering Threshold", f"More than {stats.get('threshold_days', 5)} calendar days"),
            ("Comparison Logic", "Strictly Greater Than (Days Difference > Threshold)"),
            ("Total Files Processed", stats.get("total_files", 0)),
            ("Total Worksheets Processed", stats.get("total_sheets", 0)),
            ("Total Records Examined", stats.get("total_examined", 0)),
            ("Total Matching Records", stats.get("total_matching", 0)),
            ("Total Skipped Records (<= Threshold)", stats.get("total_skipped", 0)),
            ("Total Processing Errors", stats.get("total_errors", 0)),
            ("  - Invalid Date Count", stats.get("invalid_date_count", 0)),
            ("  - Missing Date Count", stats.get("missing_date_count", 0)),
            ("  - Missing Name Count", stats.get("missing_name_count", 0)),
            ("  - Future Date Count", stats.get("future_date_count", 0)),
        ]

        row_start = 4
        ws.cell(row=row_start, column=1, value="Metric").fill = cls.HEADER_FILL
        ws.cell(row=row_start, column=1).font = cls.HEADER_FONT
        ws.cell(row=row_start, column=2, value="Value").fill = cls.HEADER_FILL
        ws.cell(row=row_start, column=2).font = cls.HEADER_FONT

        for i, (label, val) in enumerate(metrics, start=row_start + 1):
            cell_label = ws.cell(row=i, column=1, value=label)
            cell_val = ws.cell(row=i, column=2, value=val)

            cell_label.font = cls.METRIC_LABEL_FONT
            cell_label.border = cls.THIN_BORDER
            cell_val.font = cls.METRIC_VAL_FONT
            cell_val.border = cls.THIN_BORDER

            if isinstance(val, int) or isinstance(val, float):
                cell_val.alignment = Alignment(horizontal="right")
            else:
                cell_val.alignment = Alignment(horizontal="left")

        # Auto-adjust column widths
        ws.column_dimensions["A"].width = 42
        ws.column_dimensions["B"].width = 48

    @classmethod
    def _build_matching_sheet(cls, ws: Any, records: List[Record], date_format: str):
        """Format and populate the Matching Records tab."""
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"

        headers = ["Name", "Last Updated", "Days Difference", "Source File", "Source Sheet"]
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill = cls.HEADER_FILL
            cell.font = cls.HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = cls.THIN_BORDER

        for row_idx, rec in enumerate(records, start=2):
            # Formula injection prevention
            safe_name = sanitize_cell_for_export(rec.name)
            safe_file = sanitize_cell_for_export(rec.source_file)
            safe_sheet = sanitize_cell_for_export(rec.source_sheet)

            c_name = ws.cell(row=row_idx, column=1, value=safe_name)
            c_date = ws.cell(row=row_idx, column=2, value=rec.last_updated.strftime(date_format))
            c_diff = ws.cell(row=row_idx, column=3, value=rec.days_difference)
            c_file = ws.cell(row=row_idx, column=4, value=safe_file)
            c_sheet = ws.cell(row=row_idx, column=5, value=safe_sheet)

            c_name.alignment = Alignment(horizontal="left", vertical="center")
            c_date.alignment = Alignment(horizontal="center", vertical="center")
            c_diff.alignment = Alignment(horizontal="right", vertical="center")
            c_file.alignment = Alignment(horizontal="left", vertical="center")
            c_sheet.alignment = Alignment(horizontal="left", vertical="center")

            for c in [c_name, c_date, c_diff, c_file, c_sheet]:
                c.font = cls.REGULAR_FONT
                c.border = cls.THIN_BORDER

        # Enable auto-filter
        if records:
            ws.auto_filter.ref = f"A1:E{len(records) + 1}"
        else:
            ws.auto_filter.ref = "A1:E1"
            ws.cell(row=2, column=1, value="No records met the filtering criteria.").font = cls.SUBTITLE_FONT

        # Adjust column widths
        cls._adjust_column_widths(ws, headers)

    @classmethod
    def _build_errors_sheet(cls, ws: Any, errors: List[ProcessingError]):
        """Format and populate the Processing Errors tab."""
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"

        headers = ["Source File", "Source Sheet", "Row Number", "Column Name", "Error Category", "Explanation", "Original Value"]
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill = cls.ERROR_HEADER_FILL
            cell.font = cls.HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = cls.THIN_BORDER

        for row_idx, err in enumerate(errors, start=2):
            safe_file = sanitize_cell_for_export(err.source_file)
            safe_sheet = sanitize_cell_for_export(err.source_sheet)
            safe_col = sanitize_cell_for_export(err.column_name or "")
            safe_exp = sanitize_cell_for_export(err.explanation)
            safe_orig = sanitize_cell_for_export(str(err.original_value) if err.original_value is not None else "")
            cat_val = err.error_category.value if hasattr(err.error_category, "value") else str(err.error_category)
            safe_cat = sanitize_cell_for_export(cat_val)

            c_file = ws.cell(row=row_idx, column=1, value=safe_file)
            c_sheet = ws.cell(row=row_idx, column=2, value=safe_sheet)
            c_row = ws.cell(row=row_idx, column=3, value=err.row_number if err.row_number is not None else "N/A")
            c_col = ws.cell(row=row_idx, column=4, value=safe_col)
            c_cat = ws.cell(row=row_idx, column=5, value=safe_cat)
            c_exp = ws.cell(row=row_idx, column=6, value=safe_exp)
            c_orig = ws.cell(row=row_idx, column=7, value=safe_orig)

            c_file.alignment = Alignment(horizontal="left", vertical="center")
            c_sheet.alignment = Alignment(horizontal="left", vertical="center")
            c_row.alignment = Alignment(horizontal="center", vertical="center")
            c_col.alignment = Alignment(horizontal="left", vertical="center")
            c_cat.alignment = Alignment(horizontal="center", vertical="center")
            c_exp.alignment = Alignment(horizontal="left", vertical="center")
            c_orig.alignment = Alignment(horizontal="left", vertical="center")

            for c in [c_file, c_sheet, c_row, c_col, c_cat, c_exp, c_orig]:
                c.font = cls.REGULAR_FONT
                c.border = cls.THIN_BORDER

        if errors:
            ws.auto_filter.ref = f"A1:G{len(errors) + 1}"
        else:
            ws.auto_filter.ref = "A1:G1"
            ws.cell(row=2, column=1, value="No processing errors were encountered.").font = cls.SUBTITLE_FONT

        cls._adjust_column_widths(ws, headers)

    @classmethod
    def _adjust_column_widths(cls, ws: Any, headers: List[str]):
        """Dynamically adjust column widths based on header and cell content length."""
        for col_idx, header in enumerate(headers, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = len(str(header))
            # Inspect first 100 rows
            for row in range(1, min(ws.max_row + 1, 100)):
                val = ws.cell(row=row, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    @classmethod
    def generate_txt_report(
        cls,
        matching_records: List[Record],
        processing_date: datetime.date,
        threshold_days: int = 5,
        date_format: str = "%d-%b-%Y",
    ) -> str:
        """Create a clean, standardized plain text report for download."""
        lines = [
            "DataPulse - RESULTS",
            "",
            f"Processing Date: {processing_date.strftime(date_format)}",
            f"Filtering Threshold: More than {threshold_days} days",
            "=" * 50,
            "",
        ]

        if not matching_records:
            lines.extend([
                "No records satisfied the filtering condition.",
                f"All examined records were updated within the last {threshold_days} days or contained validation errors.",
                "",
                "=" * 50,
                "Total Matching Records: 0",
            ])
            return "\n".join(lines)

        for rec in matching_records:
            lines.extend([
                "-" * 48,
                f"Name: {rec.name}",
                f"Last Updated: {rec.last_updated.strftime(date_format)}",
                f"Days Difference: {rec.days_difference}",
                f"Source: {rec.source_file} (Sheet: {rec.source_sheet})",
            ])

        lines.extend([
            "-" * 48,
            "",
            "=" * 50,
            f"Total Matching Records: {len(matching_records)}",
            "=" * 50,
        ])

        return "\n".join(lines)
