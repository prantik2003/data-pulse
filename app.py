"""DataPulse – Intelligent Excel & CSV Processing Web Application.

A production-quality Streamlit application to parse, validate, and filter records
based on calendar-day aging thresholds across heterogeneous Excel and CSV files.
"""
import datetime
import io
from typing import Any, Dict, List, Optional

import pandas as pd
import streamlit as st

from config.settings import Settings
from modules.column_detector import ColumnDetectionResult
from modules.file_reader import FileData, FileReader
from modules.filter_engine import Record
from modules.processing_manager import ProcessingManager, ProcessingResult, SheetConfig
from utils.date_utils import get_current_date
from utils.file_utils import format_file_size, is_valid_extension


# Page Configuration
st.set_page_config(
    page_title="DataPulse",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling
CUSTOM_CSS = """
<style>
    /* Global layout & typography */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1200px;
    }
    
    /* Header card */
    .app-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        color: white;
        padding: 1.8rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.8rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    }
    .app-header h1 {
        color: white;
        margin: 0;
        font-size: 2.2rem;
        font-weight: 700;
        letter-spacing: -0.025em;
    }
    .app-header p {
        color: #E2E8F0;
        margin: 0.4rem 0 0 0;
        font-size: 1.05rem;
    }
    
    /* KPI Metric Cards */
    .metric-card {
        background: white;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: transform 0.15s ease-in-out;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .metric-val {
        font-size: 2.1rem;
        font-weight: 700;
        color: #0F172A;
        line-height: 1.2;
    }
    .metric-val.matching {
        color: #2563EB;
    }
    .metric-val.skipped {
        color: #64748B;
    }
    .metric-val.error {
        color: #DC2626;
    }
    .metric-label {
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        color: #64748B;
        letter-spacing: 0.05em;
        margin-top: 0.3rem;
    }

    /* Badges */
    .badge {
        display: inline-block;
        padding: 0.25rem 0.6rem;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 600;
        letter-spacing: 0.025em;
    }
    .badge-high {
        background-color: #DEF7EC;
        color: #03543F;
        border: 1px solid #84E1BC;
    }
    .badge-medium {
        background-color: #FEF08A;
        color: #713F12;
        border: 1px solid #FDE047;
    }
    .badge-low {
        background-color: #FEE2E2;
        color: #991B1B;
        border: 1px solid #F87171;
    }
    
    /* Privacy notice */
    .privacy-box {
        background-color: #F1F5F9;
        border-left: 4px solid #3B82F6;
        padding: 0.8rem 1rem;
        border-radius: 0 6px 6px 0;
        font-size: 0.85rem;
        color: #334155;
        margin-top: 1.5rem;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def init_session_state():
    """Initialize persistent Streamlit session state keys."""
    if "files_data" not in st.session_state:
        st.session_state.files_data = []
    if "sheet_configs" not in st.session_state:
        st.session_state.sheet_configs = {}
    if "processing_result" not in st.session_state:
        st.session_state.processing_result = None
    if "uploaded_file_keys" not in st.session_state:
        st.session_state.uploaded_file_keys = []
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0


init_session_state()


def reset_application():
    """Clear all session data and reset state."""
    st.session_state.files_data = []
    st.session_state.sheet_configs = {}
    st.session_state.processing_result = None
    st.session_state.uploaded_file_keys = []
    st.session_state.uploader_key += 1
    st.rerun()


# Render Sidebar Configuration
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/filter.png", width=64)
    st.title("Settings")
    
    st.markdown("### ⚙️ Processing Parameters")
    
    # 1. Filtering Threshold
    threshold_days = st.number_input(
        "Threshold (Calendar Days)",
        min_value=Settings.MIN_THRESHOLD_DAYS,
        max_value=Settings.MAX_THRESHOLD_DAYS,
        value=Settings.DEFAULT_THRESHOLD_DAYS,
        step=1,
        help="Records whose last updated date is MORE THAN this number of calendar days ago will be included. A record exactly equal to this value is excluded.",
    )
    st.caption(f"Rule: **Days Difference > {threshold_days}** (Strictly greater than)")

    # 2. Date Format Preference
    date_pref = st.selectbox(
        "Ambiguous Date Preference",
        options=Settings.DATE_PREFERENCES,
        index=0,
        help="Used to resolve ambiguous numerical dates such as '05/11/2026' when column context cannot determine order.",
    )

    # 3. Application Timezone
    app_timezone = st.selectbox(
        "Application Timezone",
        options=Settings.SUPPORTED_TIMEZONES,
        index=0,
        help="Timezone used to convert timestamps and calculate today's calendar date.",
    )

    # 4. Reference Date
    default_today = get_current_date(app_timezone)
    ref_date_option = st.date_input(
        "Reference Date (Current Date)",
        value=default_today,
        help="Current reference date used to calculate calendar-day differences.",
    )

    st.markdown("---")
    if st.button("🔄 Reset Session", use_container_width=True, type="secondary"):
        reset_application()

    st.markdown("---")
    st.markdown(
        """
        <div class="privacy-box">
            🔒 <strong>Privacy Guarantee</strong><br>
            All processing is done completely in-memory. Uploaded data is never stored on disk or sent to external servers.
        </div>
        """,
        unsafe_allow_html=True,
    )


# Main Header
st.markdown(
    f"""
    <div class="app-header">
        <h1>{Settings.APP_NAME}</h1>
        <p>{Settings.APP_SUBTITLE}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Workflow Tabs
tab_upload, tab_config, tab_results, tab_download = st.tabs([
    "1. 📁 Upload Files",
    "2. 🎛️ Column Mapping & Configuration",
    "3. 📊 Results & Preview",
    "4. 📥 Download Reports",
])


# ==========================================
# TAB 1: UPLOAD FILES
# ==========================================
with tab_upload:
    st.subheader("Upload Excel and CSV Files")
    st.markdown(
        "Upload one or multiple **Excel (.xlsx)** and **CSV (.csv)** files. "
        "The application will automatically detect sheets, encodings, delimiters, and relevant columns."
    )

    uploaded_files = st.file_uploader(
        "Drop files here or click to browse",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key=f"file_uploader_{st.session_state.uploader_key}",
    )

    if uploaded_files:
        current_file_keys = [f.name + str(f.size) for f in uploaded_files]
        if current_file_keys != st.session_state.uploaded_file_keys:
            # Process uploaded files
            with st.spinner("Ingesting and inspecting files..."):
                parsed_files: List[FileData] = []
                for uf in uploaded_files:
                    f_data = FileReader.read_file(uf, uf.name)
                    parsed_files.append(f_data)

                st.session_state.files_data = parsed_files
                st.session_state.uploaded_file_keys = current_file_keys
                # Auto-run column detection
                st.session_state.sheet_configs = ProcessingManager.analyze_uploaded_files(parsed_files)
                # Clear stale results
                st.session_state.processing_result = None

    if st.session_state.files_data:
        st.markdown("### Uploaded Files Overview")
        file_summary_data = []
        for fd in st.session_state.files_data:
            if fd.error:
                file_summary_data.append({
                    "File Name": fd.file_name,
                    "Type": fd.file_type.upper(),
                    "Size": format_file_size(fd.file_size),
                    "Worksheets": 0,
                    "Status": f"⚠️ Error: {fd.error}",
                })
            else:
                total_rows = sum(s.row_count for s in fd.sheets.values())
                file_summary_data.append({
                    "File Name": fd.file_name,
                    "Type": fd.file_type.upper(),
                    "Size": format_file_size(fd.file_size),
                    "Worksheets": len(fd.sheets),
                    "Total Rows": total_rows,
                    "Status": "✅ Ready for processing",
                })

        st.dataframe(pd.DataFrame(file_summary_data), use_container_width=True)
        st.info("Files uploaded successfully. Proceed to **Tab 2: Column Mapping & Configuration** to review detected columns.")
    else:
        st.info("No files uploaded yet. Drag and drop your `.xlsx` or `.csv` files above to get started.")


# ==========================================
# TAB 2: COLUMN MAPPING & CONFIGURATION
# ==========================================
with tab_config:
    st.subheader("Intelligent Column Detection & Overrides")
    st.markdown(
        "DataPulse analyzes both header keywords and underlying cell values to detect the "
        "**Name** and **Last Updated Date** columns for each worksheet independently. "
        "Review the detected columns below or override them if necessary."
    )

    if not st.session_state.files_data:
        st.warning("Please upload at least one Excel or CSV file in Tab 1 first.")
    else:
        has_runnable_sheets = False

        for fd in st.session_state.files_data:
            if fd.error:
                st.error(f"Cannot process **{fd.file_name}**: {fd.error}")
                continue

            with st.expander(f"📄 File: {fd.file_name} ({len(fd.sheets)} sheet(s))", expanded=True):
                for s_name, s_data in fd.sheets.items():
                    if s_data.is_empty:
                        st.warning(f"Sheet **'{s_name}'** has no data rows.")
                        continue

                    has_runnable_sheets = True
                    s_config = st.session_state.sheet_configs.get(fd.file_name, {}).get(s_name)
                    if not s_config:
                        continue

                    det: Optional[ColumnDetectionResult] = s_config.detection_result

                    col1, col2, col3, col4 = st.columns([2, 3, 3, 2])

                    with col1:
                        st.markdown(f"**Sheet:** `{s_name}`")
                        st.caption(f"Rows: {s_data.row_count:,} | Cols: {len(s_data.columns)}")
                        s_config.is_active = st.checkbox(
                            "Process Sheet",
                            value=s_config.is_active,
                            key=f"active_{fd.file_name}_{s_name}",
                        )

                    with col2:
                        name_badge_class = f"badge-{det.name_status}" if det else "badge-low"
                        name_conf_text = f"{det.name_status.upper()} ({int(det.name_confidence * 100)}%)" if det else "N/A"
                        st.markdown(
                            f"Name Column <span class='badge {name_badge_class}'>{name_conf_text}</span>",
                            unsafe_allow_html=True,
                        )
                        name_idx = s_data.columns.index(s_config.name_col) if s_config.name_col in s_data.columns else 0
                        selected_name = st.selectbox(
                            "Name Column",
                            options=s_data.columns,
                            index=name_idx,
                            key=f"name_{fd.file_name}_{s_name}",
                            label_visibility="collapsed",
                        )
                        s_config.name_col = selected_name

                    with col3:
                        date_badge_class = f"badge-{det.date_status}" if det else "badge-low"
                        date_conf_text = f"{det.date_status.upper()} ({int(det.date_confidence * 100)}%)" if det else "N/A"
                        st.markdown(
                            f"Last Updated Column <span class='badge {date_badge_class}'>{date_conf_text}</span>",
                            unsafe_allow_html=True,
                        )
                        date_idx = s_data.columns.index(s_config.date_col) if s_config.date_col in s_data.columns else 0
                        selected_date = st.selectbox(
                            "Date Column",
                            options=s_data.columns,
                            index=date_idx,
                            key=f"date_{fd.file_name}_{s_name}",
                            label_visibility="collapsed",
                        )
                        s_config.date_col = selected_date

                    with col4:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if det and det.needs_confirmation:
                            st.warning("⚠️ Review needed", icon="⚠️")
                        else:
                            st.success("✅ High confidence", icon="✅")

        st.markdown("---")
        process_col1, process_col2 = st.columns([3, 1])
        with process_col1:
            st.markdown(
                f"Ready to filter records with condition: **Last Updated older than {threshold_days} calendar days** "
                f"from reference date **{ref_date_option.strftime('%Y-%m-%d')}**."
            )
        with process_col2:
            process_btn = st.button("🚀 Process Records", type="primary", use_container_width=True)

        if process_btn:
            with st.spinner("Processing files, normalizing dates, and calculating calendar-day differences..."):
                res = ProcessingManager.execute_processing(
                    files_data=st.session_state.files_data,
                    sheet_configs=st.session_state.sheet_configs,
                    threshold_days=threshold_days,
                    current_date=ref_date_option,
                    date_preference=date_pref,
                    timezone_name=app_timezone,
                )
                st.session_state.processing_result = res
                st.success(f"Processing complete! Found {res.stats['total_matching']:,} matching record(s).")


# ==========================================
# TAB 3: RESULTS & PREVIEW
# ==========================================
with tab_results:
    res: Optional[ProcessingResult] = st.session_state.processing_result

    if not res:
        st.info("No processing results yet. Click **'🚀 Process Records'** in Tab 2 to run.")
    else:
        st.subheader("Processing Summary & KPIs")

        # KPI Metrics Row
        kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
        with kpi1:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-val">{res.stats['total_files']}</div>
                    <div class="metric-label">Files Processed</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with kpi2:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-val">{res.stats['total_sheets']}</div>
                    <div class="metric-label">Sheets Processed</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with kpi3:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-val">{res.stats['total_examined']:,}</div>
                    <div class="metric-label">Total Examined</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with kpi4:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-val matching">{res.stats['total_matching']:,}</div>
                    <div class="metric-label">Matching (> {res.stats['threshold_days']}d)</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with kpi5:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-val skipped">{res.stats['total_skipped']:,}</div>
                    <div class="metric-label">Skipped (≤ {res.stats['threshold_days']}d)</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with kpi6:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-val error">{res.stats['total_errors']:,}</div>
                    <div class="metric-label">Errors / Issues</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # Result Details Sub-tabs
        sub_matching, sub_skipped, sub_errors, sub_breakdown = st.tabs([
            f"✅ Matching Records ({len(res.matching_records)})",
            f"⏭️ Skipped Records ({len(res.skipped_records)})",
            f"⚠️ Processing Errors ({len(res.errors)})",
            "📋 File-wise Breakdown",
        ])

        # 1. Matching Records
        with sub_matching:
            if not res.matching_records:
                st.warning(
                    f"No records met the filtering condition (Last Updated date older than {res.stats['threshold_days']} days). "
                    "You may lower the threshold in the sidebar to include more records."
                )
            else:
                matching_data = [r.to_dict() for r in res.matching_records]
                df_matching = pd.DataFrame(matching_data)
                
                # Search filter
                search_query = st.text_input("🔍 Search matching records by Name or File", "")
                if search_query:
                    mask = (
                        df_matching["Name"].astype(str).str.contains(search_query, case=False, na=False) |
                        df_matching["Source File"].astype(str).str.contains(search_query, case=False, na=False)
                    )
                    df_matching = df_matching[mask]

                st.dataframe(df_matching, use_container_width=True, hide_index=True)
                st.caption(f"Showing {len(df_matching):,} of {len(res.matching_records):,} matching records.")

        # 2. Skipped Records
        with sub_skipped:
            if not res.skipped_records:
                st.info("No records were skipped due to the aging threshold.")
            else:
                skipped_data = [r.to_dict() for r in res.skipped_records]
                df_skipped = pd.DataFrame(skipped_data)
                st.dataframe(df_skipped, use_container_width=True, hide_index=True)
                st.caption("These records were updated within the threshold (Days Difference ≤ Threshold) and therefore excluded.")

        # 3. Processing Errors
        with sub_errors:
            if not res.errors:
                st.success("🎉 No processing errors encountered across all files!")
            else:
                error_data = [e.to_dict() for e in res.errors]
                df_errors = pd.DataFrame(error_data)
                st.dataframe(df_errors, use_container_width=True, hide_index=True)
                st.caption("Records with missing dates, invalid date values, missing names, or future dates.")

        # 4. File-wise Breakdown
        with sub_breakdown:
            if res.sheet_breakdown:
                st.dataframe(pd.DataFrame(res.sheet_breakdown), use_container_width=True, hide_index=True)


# ==========================================
# OUTPUT PREVIEW MODALS
# ==========================================
@st.dialog("📊 Excel Output Preview", width="large")
def show_excel_modal(processing_res: ProcessingResult):
    st.subheader(f"Filtered Records ({Settings.DEFAULT_EXCEL_FILENAME})")
    st.caption(
        f"Total Matching Records: **{len(processing_res.matching_records)}** | "
        f"Rule: **Last Updated older than {processing_res.stats.get('threshold_days', 5)} calendar days**"
    )
    if processing_res.matching_records:
        df_modal = pd.DataFrame([
            {
                "Name": r.name,
                "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                "Days Difference": r.days_difference,
                "Source File": r.source_file,
                "Source Sheet": r.source_sheet,
            }
            for r in processing_res.matching_records
        ])
        st.dataframe(df_modal, use_container_width=True, hide_index=True)
    else:
        st.info("No records met the filtering criteria.")

    st.markdown("---")
    st.download_button(
        label=f"📥 Download Excel ({len(processing_res.matching_records)} records)",
        data=processing_res.excel_bytes,
        file_name=Settings.DEFAULT_EXCEL_FILENAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="modal_btn_dl_excel",
        use_container_width=True,
    )


@st.dialog("📄 TXT Output Preview", width="large")
def show_txt_modal(processing_res: ProcessingResult):
    st.subheader(f"Plain Text Report ({Settings.DEFAULT_TXT_FILENAME})")
    st.caption(
        f"Total Matching Records: **{len(processing_res.matching_records)}** | "
        f"Rule: **Last Updated older than {processing_res.stats.get('threshold_days', 5)} calendar days**"
    )
    if processing_res.txt_report:
        st.code(processing_res.txt_report, language="text")
    else:
        st.info("No text report available.")

    st.markdown("---")
    st.download_button(
        label=f"📥 Download TXT ({len(processing_res.matching_records)} records)",
        data=processing_res.txt_report,
        file_name=Settings.DEFAULT_TXT_FILENAME,
        mime="text/plain",
        type="secondary",
        key="modal_btn_dl_txt",
        use_container_width=True,
    )


# ==========================================
# TAB 4: DOWNLOAD RESULTS
# ==========================================
with tab_download:
    res: Optional[ProcessingResult] = st.session_state.processing_result

    if not res:
        st.info("Please process your uploaded files in Tab 2 before downloading reports.")
    else:
        st.subheader("Download Consolidated Reports")
        st.markdown(
            "Download your filtered data in Excel (.xlsx) and Text (.txt) formats, "
            "or click **👁️ View Output** to preview each report in a popup modal."
        )

        dl_col1, dl_col2 = st.columns(2)

        # Excel Download
        with dl_col1:
            st.markdown(
                f"""
                <div style="background: white; border: 1px solid #CBD5E1; border-radius: 10px; padding: 1.5rem; text-align: center;">
                    <img src="https://img.icons8.com/color/96/microsoft-excel-2019--v1.png" width="48" style="margin-bottom: 0.5rem;"/>
                    <h3 style="margin: 0; color: #0F172A;">Excel Report</h3>
                    <p style="color: #64748B; font-size: 0.9rem; margin-top: 0.3rem;">
                        Clean spreadsheet containing all matching records exceeding the aging threshold.
                    </p>
                    <p style="font-weight: 600; color: #1E3A8A;">Filename: <code>{Settings.DEFAULT_EXCEL_FILENAME}</code></p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.markdown("<br>", unsafe_allow_html=True)
            b1_col1, b1_col2 = st.columns(2)
            with b1_col1:
                if st.button("👁️ View Output", key="btn_view_excel", use_container_width=True):
                    show_excel_modal(res)
            with b1_col2:
                st.download_button(
                    label=f"📥 Download Excel ({len(res.matching_records)})",
                    data=res.excel_bytes,
                    file_name=Settings.DEFAULT_EXCEL_FILENAME,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="btn_dl_excel",
                )

        # TXT Download
        with dl_col2:
            st.markdown(
                f"""
                <div style="background: white; border: 1px solid #CBD5E1; border-radius: 10px; padding: 1.5rem; text-align: center;">
                    <img src="https://img.icons8.com/color/96/txt.png" width="48" style="margin-bottom: 0.5rem;"/>
                    <h3 style="margin: 0; color: #0F172A;">Plain Text Report</h3>
                    <p style="color: #64748B; font-size: 0.9rem; margin-top: 0.3rem;">
                        Clean, human-readable text document with formatted record entries.
                    </p>
                    <p style="font-weight: 600; color: #1E3A8A;">Filename: <code>{Settings.DEFAULT_TXT_FILENAME}</code></p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.markdown("<br>", unsafe_allow_html=True)
            b2_col1, b2_col2 = st.columns(2)
            with b2_col1:
                if st.button("👁️ View Output", key="btn_view_txt", use_container_width=True):
                    show_txt_modal(res)
            with b2_col2:
                st.download_button(
                    label=f"📥 Download TXT ({len(res.matching_records)})",
                    data=res.txt_report,
                    file_name=Settings.DEFAULT_TXT_FILENAME,
                    mime="text/plain",
                    type="secondary",
                    use_container_width=True,
                    key="btn_dl_txt",
                )

        st.markdown("---")
        st.button("🔄 Start New Processing Session", on_click=reset_application, type="secondary")
