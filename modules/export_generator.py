"""Export generation module for Excel (.xlsx) and Text (.txt) reports."""
import datetime
import io
import textwrap
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

    @staticmethod
    def _sanitize_cell_str(val: Any) -> str:
        """Sanitize a cell value for ASCII table display, stripping tabs, newlines, and nulls."""
        if val is None:
            return ""
        s = str(val).strip()
        if s.lower() in ("nan", "none", "null"):
            return ""
        return " ".join(s.split())

    @classmethod
    def format_ascii_table(
        cls,
        headers: List[str],
        rows: List[List[str]],
        target_widths: Optional[List[int]] = None,
        max_widths: Optional[Dict[int, int]] = None,
    ) -> str:
        """Render a well-defined ASCII border table with column alignment and word wrapping.
        
        Args:
            headers: List of column header names.
            rows: List of row data lists.
            target_widths: Optional pre-synchronized column widths.
            max_widths: Optional dict mapping column index to max allowed character width.
        """
        if not rows:
            return ""

        if max_widths is None:
            max_widths = {1: 42, 2: 24, 4: 20}

        num_cols = len(headers)
        cleaned_rows: List[List[str]] = []
        for r in rows:
            cleaned_rows.append([cls._sanitize_cell_str(r[i]) if i < len(r) else "" for i in range(num_cols)])

        if target_widths is not None:
            col_widths = list(target_widths)
        else:
            col_widths = [len(h) for h in headers]
            for r in cleaned_rows:
                for i in range(num_cols):
                    col_widths[i] = max(col_widths[i], len(r[i]))
            for i, max_w in max_widths.items():
                if i < num_cols:
                    col_widths[i] = min(col_widths[i], max_w)

        sep = "+" + "+".join("-" * (w + 2) for w in col_widths) + "+"
        header_line = "|" + "|".join(f" {h:<{col_widths[i]}} " for i, h in enumerate(headers)) + "|"

        lines = [sep, header_line, sep]
        for r in cleaned_rows:
            wrapped_cells = [textwrap.wrap(r[i], width=col_widths[i]) or [""] for i in range(num_cols)]
            max_sublines = max(len(c) for c in wrapped_cells)
            for sub_idx in range(max_sublines):
                line_parts = []
                for i in range(num_cols):
                    sub_val = wrapped_cells[i][sub_idx] if sub_idx < len(wrapped_cells[i]) else ""
                    line_parts.append(f" {sub_val:<{col_widths[i]}} ")
                lines.append("|" + "|".join(line_parts) + "|")
            lines.append(sep)

        return "\r\n".join(lines)

    @classmethod
    def _build_matching_sheet(cls, ws: Any, records: List[Record], date_format: str):
        """Format and populate the worksheet with clear separation between Incidents and SCTasks."""
        ws.views.sheetView[0].showGridLines = True

        inc_records = [r for r in records if r.incident]
        sctask_records = [r for r in records if r.sctask]
        other_records = [r for r in records if not r.incident and not r.sctask]

        current_row = 1

        # 1. INCIDENTS SECTION
        if inc_records:
            ws.cell(row=current_row, column=1, value=f"INCIDENTS ({len(inc_records)} Records)")
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            c_banner = ws.cell(row=current_row, column=1)
            c_banner.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            c_banner.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
            c_banner.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[current_row].height = 24
            current_row += 1

            inc_headers = ["Incident", "Short Description", "Name", "Last Updated", "Source File"]
            for col_idx, h in enumerate(inc_headers, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=h)
                cell.fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
                cell.font = cls.HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = cls.THIN_BORDER
            ws.row_dimensions[current_row].height = 20
            current_row += 1

            for rec in inc_records:
                c1 = ws.cell(row=current_row, column=1, value=sanitize_cell_for_export(rec.incident))
                c2 = ws.cell(row=current_row, column=2, value=sanitize_cell_for_export(rec.short_description))
                c3 = ws.cell(row=current_row, column=3, value=sanitize_cell_for_export(rec.name))
                c4 = ws.cell(row=current_row, column=4, value=rec.last_updated.strftime(date_format))
                c5 = ws.cell(row=current_row, column=5, value=sanitize_cell_for_export(rec.source_file))

                c1.alignment = Alignment(horizontal="center", vertical="center")
                c2.alignment = Alignment(horizontal="left", vertical="center")
                c3.alignment = Alignment(horizontal="left", vertical="center")
                c4.alignment = Alignment(horizontal="center", vertical="center")
                c5.alignment = Alignment(horizontal="left", vertical="center")

                for c in [c1, c2, c3, c4, c5]:
                    c.font = cls.REGULAR_FONT
                    c.border = cls.THIN_BORDER
                current_row += 1

            current_row += 1  # Blank line separator

        # 2. SERVICE CATALOG TASKS (SCTASKS) SECTION
        if sctask_records:
            ws.cell(row=current_row, column=1, value=f"SERVICE CATALOG TASKS (SCTASKS) ({len(sctask_records)} Records)")
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            c_banner = ws.cell(row=current_row, column=1)
            c_banner.fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
            c_banner.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
            c_banner.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[current_row].height = 24
            current_row += 1

            sctask_headers = ["SCTask", "Short Description", "Name", "Last Updated", "Source File"]
            for col_idx, h in enumerate(sctask_headers, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=h)
                cell.fill = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")
                cell.font = cls.HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = cls.THIN_BORDER
            ws.row_dimensions[current_row].height = 20
            current_row += 1

            for rec in sctask_records:
                c1 = ws.cell(row=current_row, column=1, value=sanitize_cell_for_export(rec.sctask))
                c2 = ws.cell(row=current_row, column=2, value=sanitize_cell_for_export(rec.short_description))
                c3 = ws.cell(row=current_row, column=3, value=sanitize_cell_for_export(rec.name))
                c4 = ws.cell(row=current_row, column=4, value=rec.last_updated.strftime(date_format))
                c5 = ws.cell(row=current_row, column=5, value=sanitize_cell_for_export(rec.source_file))

                c1.alignment = Alignment(horizontal="center", vertical="center")
                c2.alignment = Alignment(horizontal="left", vertical="center")
                c3.alignment = Alignment(horizontal="left", vertical="center")
                c4.alignment = Alignment(horizontal="center", vertical="center")
                c5.alignment = Alignment(horizontal="left", vertical="center")

                for c in [c1, c2, c3, c4, c5]:
                    c.font = cls.REGULAR_FONT
                    c.border = cls.THIN_BORDER
                current_row += 1

            current_row += 1

        # 3. OTHER RECORDS SECTION (generic records without ticket IDs)
        if other_records:
            ws.cell(row=current_row, column=1, value=f"OTHER MATCHING RECORDS ({len(other_records)} Records)")
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            c_banner = ws.cell(row=current_row, column=1)
            c_banner.fill = PatternFill(start_color="475569", end_color="475569", fill_type="solid")
            c_banner.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
            c_banner.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[current_row].height = 24
            current_row += 1

            other_headers = ["Short Description", "Name", "Last Updated", "Source File"] if any(r.short_description for r in other_records) else ["Name", "Last Updated", "Source File"]
            for col_idx, h in enumerate(other_headers, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=h)
                cell.fill = PatternFill(start_color="64748B", end_color="64748B", fill_type="solid")
                cell.font = cls.HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = cls.THIN_BORDER
            ws.row_dimensions[current_row].height = 20
            current_row += 1

            for rec in other_records:
                if len(other_headers) == 4:
                    c1 = ws.cell(row=current_row, column=1, value=sanitize_cell_for_export(rec.short_description))
                    c2 = ws.cell(row=current_row, column=2, value=sanitize_cell_for_export(rec.name))
                    c3 = ws.cell(row=current_row, column=3, value=rec.last_updated.strftime(date_format))
                    c4 = ws.cell(row=current_row, column=4, value=sanitize_cell_for_export(rec.source_file))
                    cells = [c1, c2, c3, c4]
                else:
                    c1 = ws.cell(row=current_row, column=1, value=sanitize_cell_for_export(rec.name))
                    c2 = ws.cell(row=current_row, column=2, value=rec.last_updated.strftime(date_format))
                    c3 = ws.cell(row=current_row, column=3, value=sanitize_cell_for_export(rec.source_file))
                    cells = [c1, c2, c3]

                for c in cells:
                    c.font = cls.REGULAR_FONT
                    c.border = cls.THIN_BORDER
                current_row += 1

        if not records:
            ws.cell(row=1, column=1, value="No records met the filtering criteria.").font = cls.SUBTITLE_FONT

        # Adjust column widths
        cls._adjust_column_widths(ws, ["Incident/SCTask", "Short Description", "Name", "Last Updated", "Source File"])

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
        """Create a clean, standardized plain text report with well-defined ASCII tables."""
        if not matching_records:
            banner = "=" * 90
            lines = [
                banner,
                "DataPulse - CONSOLIDATED REPORT",
                f"Processing Date: {processing_date.strftime(date_format)} | Threshold: More than {threshold_days} calendar days",
                banner,
                "",
                "No records satisfied the filtering condition.",
                f"All examined records were updated within the last {threshold_days} days or contained validation errors.",
                "",
                banner,
                "Total Matching Records: 0",
                banner,
            ]
            return "\r\n".join(lines)

        inc_records = [r for r in matching_records if r.incident]
        sctask_records = [r for r in matching_records if r.sctask]
        other_records = [r for r in matching_records if not r.incident and not r.sctask]

        inc_rows: List[List[str]] = [
            [
                r.incident,
                r.short_description,
                r.name,
                r.last_updated.strftime(date_format),
                r.source_file,
            ]
            for r in inc_records
        ]
        sctask_rows: List[List[str]] = [
            [
                r.sctask,
                r.short_description,
                r.name,
                r.last_updated.strftime(date_format),
                r.source_file,
            ]
            for r in sctask_records
        ]

        # Calculate synchronized target column widths for standard 5-column ticket tables
        target_widths_5 = None
        all_ticket_rows = inc_rows + sctask_rows
        if all_ticket_rows:
            target_widths_5 = [
                max(len("Incident"), len("SCTask"), max((len(cls._sanitize_cell_str(r[0])) for r in all_ticket_rows), default=0)),
                min(max(len("Short Description"), max((len(cls._sanitize_cell_str(r[1])) for r in all_ticket_rows), default=0)), 42),
                min(max(len("Name"), max((len(cls._sanitize_cell_str(r[2])) for r in all_ticket_rows), default=0)), 24),
                max(len("Last Updated"), 11),
                min(max(len("Source File"), max((len(cls._sanitize_cell_str(r[4])) for r in all_ticket_rows), default=0)), 20),
            ]

        t_inc = cls.format_ascii_table(
            ["Incident", "Short Description", "Name", "Last Updated", "Source File"],
            inc_rows,
            target_widths=target_widths_5,
        ) if inc_records else ""

        t_sctask = cls.format_ascii_table(
            ["SCTask", "Short Description", "Name", "Last Updated", "Source File"],
            sctask_rows,
            target_widths=target_widths_5,
        ) if sctask_records else ""

        t_other = ""
        if other_records:
            has_desc = any(r.short_description for r in other_records)
            other_headers = ["Short Description", "Name", "Last Updated", "Source File"] if has_desc else ["Name", "Last Updated", "Source File"]
            other_rows = []
            for r in other_records:
                if has_desc:
                    other_rows.append([r.short_description, r.name, r.last_updated.strftime(date_format), r.source_file])
                else:
                    other_rows.append([r.name, r.last_updated.strftime(date_format), r.source_file])
            t_other = cls.format_ascii_table(other_headers, other_rows)

        # Determine report banner width based on the widest table
        table_widths = [len(t.splitlines()[0]) for t in [t_inc, t_sctask, t_other] if t]
        banner_width = max(max(table_widths, default=90), 90)
        banner = "=" * banner_width

        lines = [
            banner,
            "DataPulse - CONSOLIDATED REPORT",
            f"Processing Date: {processing_date.strftime(date_format)} | Threshold: More than {threshold_days} calendar days",
            banner,
            "",
        ]

        if t_inc:
            lines.append(f"INCIDENTS (Total: {len(inc_records)})")
            lines.append(t_inc)
            lines.append("")

        if t_sctask:
            lines.append(f"SERVICE CATALOG TASKS (SCTASKS) (Total: {len(sctask_records)})")
            lines.append(t_sctask)
            lines.append("")

        if t_other:
            lines.append(f"OTHER MATCHING RECORDS (Total: {len(other_records)})")
            lines.append(t_other)
            lines.append("")

        parts = []
        if inc_records:
            parts.append(f"{len(inc_records)} Incidents")
        if sctask_records:
            parts.append(f"{len(sctask_records)} SCTasks")
        if other_records:
            parts.append(f"{len(other_records)} Other")
        breakdown_str = f" ({', '.join(parts)})" if parts else ""

        lines.extend([
            banner,
            f"Total Matching Records: {len(matching_records)}{breakdown_str}",
            banner,
        ])

        return "\r\n".join(lines)

    @staticmethod
    def escape_rtf(text: Any) -> str:
        """Escape text for RTF syntax, handling special symbols and unicode."""
        if text is None:
            return ""
        s = str(text).strip()
        if s.lower() in ("nan", "none", "null"):
            return ""
        out = []
        for char in s:
            if char == "\\":
                out.append(r"\\")
            elif char == "{":
                out.append(r"\{")
            elif char == "}":
                out.append(r"\}")
            elif ord(char) > 127:
                out.append(f"\\u{ord(char)}?")
            elif char == "\n":
                out.append(" ")
            elif char == "\r":
                pass
            elif char == "\t":
                out.append(" ")
            else:
                out.append(char)
        return " ".join("".join(out).split())

    @classmethod
    def generate_rtf_report(
        cls,
        matching_records: List[Record],
        processing_date: datetime.date,
        threshold_days: int = 5,
        date_format: str = "%d-%b-%Y",
    ) -> str:
        """Create a styled Rich Text Format (.rtf) report with native tables and formatted headers."""
        inc_records = [r for r in matching_records if r.incident]
        sctask_records = [r for r in matching_records if r.sctask]
        other_records = [r for r in matching_records if not r.incident and not r.sctask]

        # RTF document header with color table and font definitions
        rtf = [
            r"{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Calibri;}{\f1\fnil\fcharset0 Arial;}}",
            r"{\colortbl ;\red255\green255\blue255;\red30\green58\blue138;\red37\green99\blue235;\red15\green118\blue110;\red13\green148\blue136;\red203\green213\blue225;\red15\green23\blue42;\red100\green116\blue139;\red248\green250\blue252;\red71\green85\blue105;}",
            r"\viewkind4\uc1\paperw12240\paperh15840\margl720\margr720\margt720\margb720",
            r"\pard\f0\fs28\b\cf2 DataPulse - CONSOLIDATED REPORT\b0\par",
            f"\\fs20\\cf8 Processing Date: {processing_date.strftime(date_format)} | Threshold: More than {threshold_days} calendar days\\par",
            r"\fs18\cf6\strike------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\strike0\cf0\par\par",
        ]

        if not matching_records:
            rtf.append(r"\pard\f0\fs20\cf7 No records satisfied the filtering condition.\par")
            rtf.append(f"\\fs18\\cf8 All examined records were updated within the last {threshold_days} days or contained validation errors.\\par\\par")
            rtf.append(r"\pard\f0\fs22\b\cf2 Total Matching Records: 0\b0\par}")
            return "\n".join(rtf)

        cellx_5 = [1800, 5800, 8000, 9300, 10800]
        aligns_5 = [r"\qc", r"\ql", r"\ql", r"\qc", r"\ql"]

        def render_table(headers: List[str], rows: List[List[str]], header_color_idx: int) -> str:
            t_lines = []
            t_lines.append(r"\trowd\trgaph90\trleft0")
            for cx in cellx_5:
                t_lines.append(f"\\clbrdrt\\brdrs\\brdrw10\\brdrcf6\\clbrdrl\\brdrs\\brdrw10\\brdrcf6\\clbrdrb\\brdrs\\brdrw10\\brdrcf6\\clbrdrr\\brdrs\\brdrw10\\brdrcf6\\clcbpat{header_color_idx}\\cellx{cx}")
            for h in headers:
                t_lines.append(f"\\pard\\intbl\\qc\\f0\\fs19\\b\\cf1 {cls.escape_rtf(h)}\\cell ")
            t_lines.append(r"\row")

            for row_idx, r in enumerate(rows):
                bg = r"\clcbpat9" if row_idx % 2 == 1 else r"\clcbpat1"
                t_lines.append(r"\trowd\trgaph90\trleft0")
                for cx in cellx_5:
                    t_lines.append(f"\\clbrdrt\\brdrs\\brdrw10\\brdrcf6\\clbrdrl\\brdrs\\brdrw10\\brdrcf6\\clbrdrb\\brdrs\\brdrw10\\brdrcf6\\clbrdrr\\brdrs\\brdrw10\\brdrcf6{bg}\\cellx{cx}")
                for val, al in zip(r, aligns_5):
                    t_lines.append(f"\\pard\\intbl{al}\\f0\\fs18\\cf7 {cls.escape_rtf(val)}\\cell ")
                t_lines.append(r"\row")
            return "\n".join(t_lines)

        if inc_records:
            rtf.append(f"\\pard\\f0\\fs22\\b\\cf3 INCIDENTS ({len(inc_records)} Records)\\b0\\par\\fs18")
            inc_headers = ["Incident", "Short Description", "Name", "Last Updated", "Source File"]
            inc_rows = [[r.incident, r.short_description, r.name, r.last_updated.strftime(date_format), r.source_file] for r in inc_records]
            rtf.append(render_table(inc_headers, inc_rows, 3))
            rtf.append(r"\pard\par")

        if sctask_records:
            rtf.append(f"\\pard\\f0\\fs22\\b\\cf5 SERVICE CATALOG TASKS (SCTASKS) ({len(sctask_records)} Records)\\b0\\par\\fs18")
            sctask_headers = ["SCTask", "Short Description", "Name", "Last Updated", "Source File"]
            sctask_rows = [[r.sctask, r.short_description, r.name, r.last_updated.strftime(date_format), r.source_file] for r in sctask_records]
            rtf.append(render_table(sctask_headers, sctask_rows, 5))
            rtf.append(r"\pard\par")

        if other_records:
            rtf.append(f"\\pard\\f0\\fs22\\b\\cf10 OTHER MATCHING RECORDS ({len(other_records)} Records)\\b0\\par\\fs18")
            has_desc = any(r.short_description for r in other_records)
            other_headers = ["Short Description", "Name", "Last Updated", "Source File"] if has_desc else ["Name", "Last Updated", "Source File"]
            cellx_other = [5800, 8000, 9300, 10800] if has_desc else [5000, 7500, 10800]
            aligns_other = [r"\ql", r"\ql", r"\qc", r"\ql"] if has_desc else [r"\ql", r"\qc", r"\ql"]

            t_other = [r"\trowd\trgaph90\trleft0"]
            for cx in cellx_other:
                t_other.append(f"\\clbrdrt\\brdrs\\brdrw10\\brdrcf6\\clbrdrl\\brdrs\\brdrw10\\brdrcf6\\clbrdrb\\brdrs\\brdrw10\\brdrcf6\\clbrdrr\\brdrs\\brdrw10\\brdrcf6\\clcbpat10\\cellx{cx}")
            for h in other_headers:
                t_other.append(f"\\pard\\intbl\\qc\\f0\\fs19\\b\\cf1 {cls.escape_rtf(h)}\\cell ")
            t_other.append(r"\row")

            for row_idx, r in enumerate(other_records):
                bg = r"\clcbpat9" if row_idx % 2 == 1 else r"\clcbpat1"
                t_other.append(r"\trowd\trgaph90\trleft0")
                for cx in cellx_other:
                    t_other.append(f"\\clbrdrt\\brdrs\\brdrw10\\brdrcf6\\clbrdrl\\brdrs\\brdrw10\\brdrcf6\\clbrdrb\\brdrs\\brdrw10\\brdrcf6\\clbrdrr\\brdrs\\brdrw10\\brdrcf6{bg}\\cellx{cx}")
                vals = [r.short_description, r.name, r.last_updated.strftime(date_format), r.source_file] if has_desc else [r.name, r.last_updated.strftime(date_format), r.source_file]
                for val, al in zip(vals, aligns_other):
                    t_other.append(f"\\pard\\intbl{al}\\f0\\fs18\\cf7 {cls.escape_rtf(val)}\\cell ")
                t_other.append(r"\row")
            rtf.append("\n".join(t_other))
            rtf.append(r"\pard\par")

        parts = []
        if inc_records:
            parts.append(f"{len(inc_records)} Incidents")
        if sctask_records:
            parts.append(f"{len(sctask_records)} SCTasks")
        if other_records:
            parts.append(f"{len(other_records)} Other")
        parts_str = ", ".join(parts)
        breakdown_str = f" ({parts_str})" if parts else ""

        rtf.append(r"\pard\par")
        rtf.append(f"\\pard\\f0\\fs22\\b\\cf2 Total Matching Records: {len(matching_records)}{cls.escape_rtf(breakdown_str)}\\b0\\par")
        rtf.append("}")

        return "\n".join(rtf)
