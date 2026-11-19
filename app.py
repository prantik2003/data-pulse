"""DataPulse – Intelligent Excel & CSV Processing Web Application.

A production-quality Streamlit application to parse, validate, and filter records
based on calendar-day aging thresholds across heterogeneous Excel and CSV files.
"""
import datetime
import importlib
import io
import sys
from typing import Any, Dict, List, Optional

# Ensure workspace modules are reloaded from disk in running Streamlit sessions
for _mod_name in (
    "config.settings",
    "modules.export_generator",
    "modules.processing_manager",
    "modules.filter_engine",
    "modules.file_reader",
    "modules.column_detector",
):
    if _mod_name in sys.modules:
        try:
            importlib.reload(sys.modules[_mod_name])
        except Exception:
            pass

import pandas as pd
import streamlit as st
import streamlit_shadcn_ui as ui

from config.settings import Settings
from modules.column_detector import ColumnDetectionResult
from modules.file_reader import FileData, FileReader
from modules.filter_engine import Record
from modules.processing_manager import ProcessingManager, ProcessingResult, SheetConfig
from utils.date_utils import get_current_date
from utils.file_utils import format_file_size, is_valid_extension

# Safe fallback for hot-reloaded runtime caches
DEFAULT_RTF_FILENAME = getattr(Settings, "DEFAULT_RTF_FILENAME", "filtered_records.rtf")


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
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Global Viewport & Ambient Background */
    .stApp {
        background-color: #070A12 !important;
        background-image: 
            radial-gradient(circle at 10% 12%, rgba(37, 99, 235, 0.20) 0%, transparent 45%),
            radial-gradient(circle at 90% 15%, rgba(124, 58, 237, 0.18) 0%, transparent 45%),
            radial-gradient(circle at 50% 55%, rgba(6, 182, 212, 0.10) 0%, transparent 55%),
            radial-gradient(circle at 85% 90%, rgba(30, 58, 138, 0.22) 0%, transparent 45%),
            radial-gradient(rgba(255, 255, 255, 0.08) 1px, transparent 1px) !important;
        background-size: 100% 100%, 100% 100%, 100% 100%, 100% 100%, 30px 30px !important;
        background-attachment: fixed !important;
        color: #F1F5F9 !important;
    }

    /* Global layout & spacing */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 4rem;
        max-width: 1260px;
    }

    /* Sidebar Glassmorphic Styling */
    [data-testid="stSidebar"] {
        background: rgba(9, 13, 24, 0.92) !important;
        backdrop-filter: blur(24px) !important;
        -webkit-backdrop-filter: blur(24px) !important;
        border-right: 1px solid rgba(255, 255, 255, 0.08) !important;
        box-shadow: 4px 0 24px rgba(0, 0, 0, 0.4) !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: rgba(255, 255, 255, 0.08) !important;
    }

    /* Modern App Header */
    .app-header-container {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.9) 0%, rgba(24, 33, 56, 0.8) 50%, rgba(30, 58, 138, 0.65) 100%);
        backdrop-filter: blur(24px);
        -webkit-backdrop-filter: blur(24px);
        border-radius: 20px;
        padding: 2.2rem 2.4rem;
        margin-bottom: 2rem;
        border: 1px solid rgba(59, 130, 246, 0.28);
        box-shadow: 0 12px 35px -5px rgba(0, 0, 0, 0.6), 0 0 30px -5px rgba(37, 99, 235, 0.25);
        position: relative;
        overflow: hidden;
    }

    .app-header-container::before {
        content: '';
        position: absolute;
        top: -80px;
        right: -80px;
        width: 280px;
        height: 280px;
        background: radial-gradient(circle, rgba(59, 130, 246, 0.35) 0%, rgba(99, 102, 241, 0.15) 50%, transparent 70%);
        border-radius: 50%;
        pointer-events: none;
    }

    .app-header-content {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 1.5rem;
        position: relative;
        z-index: 1;
    }

    .app-header-left {
        display: flex;
        align-items: center;
        gap: 1.2rem;
    }

    .app-header-icon {
        background: rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.2);
        width: 60px;
        height: 60px;
        border-radius: 16px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 30px;
        box-shadow: 0 4px 20px rgba(37, 99, 235, 0.3), inset 0 1px 0 rgba(255, 255, 255, 0.2);
    }

    .app-title-group h1 {
        color: #FFFFFF;
        margin: 0;
        font-size: 2.2rem;
        font-weight: 800;
        letter-spacing: -0.03em;
        line-height: 1.2;
    }

    .app-title-group p {
        color: #94A3B8;
        margin: 0.35rem 0 0 0;
        font-size: 0.98rem;
        font-weight: 500;
    }

    .app-header-badges {
        display: flex;
        gap: 0.65rem;
        flex-wrap: wrap;
    }

    .header-chip {
        display: inline-flex;
        align-items: center;
        gap: 0.45rem;
        padding: 0.4rem 0.85rem;
        background: rgba(15, 23, 42, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 9999px;
        color: #E2E8F0;
        font-size: 0.83rem;
        font-weight: 600;
        backdrop-filter: blur(6px);
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.25);
    }

    .status-dot-green {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: #10B981;
        box-shadow: 0 0 8px #10B981;
        display: inline-block;
    }

    .status-dot-blue {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: #3B82F6;
        box-shadow: 0 0 8px #3B82F6;
        display: inline-block;
    }

    /* Modern Segmented Tab Bar */
    div[data-baseweb="tab-list"] {
        background: rgba(15, 23, 42, 0.85) !important;
        backdrop-filter: blur(20px) !important;
        padding: 6px !important;
        border-radius: 14px !important;
        gap: 8px !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        margin-bottom: 1.8rem !important;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3) !important;
    }

    div[data-baseweb="tab-border"] {
        display: none !important;
    }

    button[data-baseweb="tab"] {
        border-radius: 10px !important;
        padding: 0.7rem 1.6rem !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        color: #94A3B8 !important;
        background: transparent !important;
        border: 1px solid transparent !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }

    button[data-baseweb="tab"]:hover {
        color: #FFFFFF !important;
        background: rgba(255, 255, 255, 0.06) !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, rgba(37, 99, 235, 0.35) 0%, rgba(59, 130, 246, 0.2) 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(59, 130, 246, 0.5) !important;
        box-shadow: 0 0 16px rgba(59, 130, 246, 0.3) !important;
    }

    /* 3-Step Guide Banner */
    .step-banner {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
        gap: 1.2rem;
        margin-bottom: 1.8rem;
    }

    .step-card {
        background: rgba(17, 24, 39, 0.75);
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 1.2rem 1.4rem;
        display: flex;
        align-items: center;
        gap: 1rem;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25);
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
    }

    .step-card:hover {
        border-color: rgba(59, 130, 246, 0.45);
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.4), 0 0 20px rgba(59, 130, 246, 0.15);
    }

    .step-num {
        width: 42px;
        height: 42px;
        border-radius: 12px;
        background: linear-gradient(135deg, rgba(37, 99, 235, 0.35) 0%, rgba(59, 130, 246, 0.18) 100%);
        color: #60A5FA;
        border: 1px solid rgba(96, 165, 250, 0.35);
        font-weight: 800;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.05rem;
        flex-shrink: 0;
        box-shadow: 0 0 12px rgba(37, 99, 235, 0.25);
    }

    .step-text h4 {
        margin: 0;
        font-size: 0.95rem;
        font-weight: 700;
        color: #F8FAFC;
    }

    .step-text p {
        margin: 0.25rem 0 0 0;
        font-size: 0.82rem;
        color: #94A3B8;
        line-height: 1.35;
    }

    /* Colored KPI Metric Boxes */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 1rem;
        margin-bottom: 1.8rem;
    }

    @media (max-width: 1100px) {
        .kpi-grid {
            grid-template-columns: repeat(3, 1fr);
        }
    }

    @media (max-width: 650px) {
        .kpi-grid {
            grid-template-columns: repeat(2, 1fr);
        }
    }

    .kpi-box {
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
        border-radius: 16px;
        padding: 1.3rem 1.15rem;
        text-align: left;
        position: relative;
        overflow: hidden;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        min-height: 142px;
    }

    .kpi-box:hover {
        transform: translateY(-3px);
    }

    /* Box 1: Files Processed (Electric Blue) */
    .kpi-box-blue {
        background: linear-gradient(135deg, rgba(30, 58, 138, 0.45) 0%, rgba(15, 23, 42, 0.85) 100%);
        border: 1px solid rgba(59, 130, 246, 0.4);
        box-shadow: 0 4px 20px rgba(37, 99, 235, 0.15);
    }
    .kpi-box-blue:hover {
        border-color: rgba(96, 165, 250, 0.7);
        box-shadow: 0 8px 25px rgba(37, 99, 235, 0.3), 0 0 15px rgba(59, 130, 246, 0.2);
    }
    .kpi-box-blue::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #2563EB, #60A5FA);
        box-shadow: 0 0 10px #3B82F6;
    }

    /* Box 2: Sheets Processed (Cyan / Aqua) */
    .kpi-box-cyan {
        background: linear-gradient(135deg, rgba(14, 116, 144, 0.45) 0%, rgba(15, 23, 42, 0.85) 100%);
        border: 1px solid rgba(6, 182, 212, 0.4);
        box-shadow: 0 4px 20px rgba(6, 182, 212, 0.15);
    }
    .kpi-box-cyan:hover {
        border-color: rgba(34, 211, 238, 0.7);
        box-shadow: 0 8px 25px rgba(6, 182, 212, 0.3), 0 0 15px rgba(6, 182, 212, 0.2);
    }
    .kpi-box-cyan::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #0891B2, #22D3EE);
        box-shadow: 0 0 10px #06B6D4;
    }

    /* Box 3: Total Examined (Purple / Violet) */
    .kpi-box-purple {
        background: linear-gradient(135deg, rgba(88, 28, 135, 0.45) 0%, rgba(15, 23, 42, 0.85) 100%);
        border: 1px solid rgba(168, 85, 247, 0.4);
        box-shadow: 0 4px 20px rgba(168, 85, 247, 0.15);
    }
    .kpi-box-purple:hover {
        border-color: rgba(192, 132, 252, 0.7);
        box-shadow: 0 8px 25px rgba(168, 85, 247, 0.3), 0 0 15px rgba(168, 85, 247, 0.2);
    }
    .kpi-box-purple::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #7C3AED, #C084FC);
        box-shadow: 0 0 10px #A855F7;
    }

    /* Box 4: Matching Records (Emerald Green) */
    .kpi-box-emerald {
        background: linear-gradient(135deg, rgba(6, 95, 70, 0.5) 0%, rgba(15, 23, 42, 0.85) 100%);
        border: 1px solid rgba(16, 185, 129, 0.5);
        box-shadow: 0 4px 22px rgba(16, 185, 129, 0.2);
    }
    .kpi-box-emerald:hover {
        border-color: rgba(52, 211, 153, 0.8);
        box-shadow: 0 8px 28px rgba(16, 185, 129, 0.35), 0 0 18px rgba(16, 185, 129, 0.25);
    }
    .kpi-box-emerald::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #059669, #34D399);
        box-shadow: 0 0 12px #10B981;
    }

    /* Box 5: Skipped Records (Warm Amber / Gold) */
    .kpi-box-amber {
        background: linear-gradient(135deg, rgba(146, 64, 14, 0.45) 0%, rgba(15, 23, 42, 0.85) 100%);
        border: 1px solid rgba(245, 158, 11, 0.45);
        box-shadow: 0 4px 20px rgba(245, 158, 11, 0.15);
    }
    .kpi-box-amber:hover {
        border-color: rgba(251, 191, 36, 0.75);
        box-shadow: 0 8px 25px rgba(245, 158, 11, 0.3), 0 0 15px rgba(245, 158, 11, 0.2);
    }
    .kpi-box-amber::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #D97706, #FBBF24);
        box-shadow: 0 0 10px #F59E0B;
    }

    /* Box 6: Issues / Errors (Rose Crimson or Clean Slate/Emerald) */
    .kpi-box-crimson {
        background: linear-gradient(135deg, rgba(159, 18, 57, 0.5) 0%, rgba(15, 23, 42, 0.85) 100%);
        border: 1px solid rgba(244, 63, 94, 0.5);
        box-shadow: 0 4px 22px rgba(244, 63, 94, 0.2);
    }
    .kpi-box-crimson:hover {
        border-color: rgba(251, 113, 133, 0.8);
        box-shadow: 0 8px 28px rgba(244, 63, 94, 0.35), 0 0 18px rgba(244, 63, 94, 0.25);
    }
    .kpi-box-crimson::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #E11D48, #FB7185);
        box-shadow: 0 0 12px #F43F5E;
    }

    .kpi-box-clean {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.85) 0%, rgba(30, 41, 59, 0.7) 100%);
        border: 1px solid rgba(52, 211, 153, 0.35);
        box-shadow: 0 4px 20px rgba(16, 185, 129, 0.12);
    }
    .kpi-box-clean:hover {
        border-color: rgba(52, 211, 153, 0.65);
        box-shadow: 0 8px 25px rgba(16, 185, 129, 0.25);
    }
    .kpi-box-clean::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3px;
        background: linear-gradient(90deg, #059669, #34D399);
        box-shadow: 0 0 10px #10B981;
    }

    /* Inner box typography */
    .kpi-box-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 0.35rem;
    }
    .kpi-box-label {
        font-size: 0.76rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        text-align: left;
    }
    .kpi-box-icon {
        font-size: 1.15rem;
    }
    .kpi-box-value {
        font-size: 2.15rem;
        font-weight: 800;
        letter-spacing: -0.03em;
        line-height: 1.15;
        margin: 0.3rem 0;
        text-align: left;
    }
    .kpi-box-desc {
        font-size: 0.78rem;
        font-weight: 500;
        text-align: left;
        margin: 0;
    }

    /* Section Header Banners (Incidents vs SCTasks) */
    .section-banner-inc {
        background: linear-gradient(90deg, #1E3A8A 0%, #2563EB 70%, #3B82F6 100%);
        color: white;
        padding: 0.85rem 1.3rem;
        border-radius: 12px;
        font-weight: 700;
        font-size: 1.05rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin: 1.5rem 0 0.9rem 0;
        box-shadow: 0 4px 16px rgba(37, 99, 235, 0.3), inset 0 1px 0 rgba(255, 255, 255, 0.2);
        border: 1px solid rgba(96, 165, 250, 0.3);
    }

    .section-banner-sctask {
        background: linear-gradient(90deg, #0F766E 0%, #0D9488 70%, #14B8A6 100%);
        color: white;
        padding: 0.85rem 1.3rem;
        border-radius: 12px;
        font-weight: 700;
        font-size: 1.05rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin: 1.5rem 0 0.9rem 0;
        box-shadow: 0 4px 16px rgba(13, 148, 136, 0.3), inset 0 1px 0 rgba(255, 255, 255, 0.2);
        border: 1px solid rgba(45, 212, 191, 0.3);
    }

    .section-banner-other {
        background: linear-gradient(90deg, #1E293B 0%, #334155 70%, #475569 100%);
        color: white;
        padding: 0.85rem 1.3rem;
        border-radius: 12px;
        font-weight: 700;
        font-size: 1.05rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin: 1.5rem 0 0.9rem 0;
        border: 1px solid rgba(255, 255, 255, 0.12);
    }

    .banner-count-badge {
        background: rgba(0, 0, 0, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.2);
        padding: 0.22rem 0.7rem;
        border-radius: 9999px;
        font-size: 0.82rem;
        font-weight: 600;
        backdrop-filter: blur(4px);
    }

    /* Modern Export Cards */
    .export-card {
        background: rgba(17, 24, 39, 0.8);
        backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 20px;
        padding: 2rem;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.35);
        position: relative;
        overflow: hidden;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }

    .export-card:hover {
        border-color: rgba(255, 255, 255, 0.25);
        transform: translateY(-4px);
        box-shadow: 0 16px 40px rgba(0, 0, 0, 0.5), 0 0 25px rgba(59, 130, 246, 0.2);
    }

    .export-card-excel::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: linear-gradient(90deg, #059669 0%, #10B981 50%, #34D399 100%);
        box-shadow: 0 0 12px rgba(16, 185, 129, 0.6);
    }

    .export-card-txt::before,
    .export-card-rtf::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: linear-gradient(90deg, #2563EB 0%, #3B82F6 50%, #60A5FA 100%);
        box-shadow: 0 0 12px rgba(59, 130, 246, 0.6);
    }

    .export-header {
        display: flex;
        align-items: center;
        gap: 1rem;
        margin-bottom: 0.9rem;
    }

    .export-icon-box {
        width: 54px;
        height: 54px;
        border-radius: 14px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 26px;
    }

    .export-icon-excel {
        background: rgba(16, 185, 129, 0.15);
        color: #34D399;
        border: 1px solid rgba(16, 185, 129, 0.3);
        box-shadow: 0 0 16px rgba(16, 185, 129, 0.2);
    }

    .export-icon-txt,
    .export-icon-rtf {
        background: rgba(59, 130, 246, 0.15);
        color: #60A5FA;
        border: 1px solid rgba(59, 130, 246, 0.3);
        box-shadow: 0 0 16px rgba(59, 130, 246, 0.2);
    }

    .export-title-group h3 {
        margin: 0;
        font-size: 1.3rem;
        font-weight: 800;
        color: #F8FAFC;
    }

    .export-title-group .export-badge {
        display: inline-block;
        font-size: 0.72rem;
        font-weight: 700;
        padding: 0.18rem 0.55rem;
        border-radius: 6px;
        margin-top: 0.25rem;
        letter-spacing: 0.04em;
    }

    .badge-excel {
        background: rgba(16, 185, 129, 0.2);
        color: #34D399;
        border: 1px solid rgba(16, 185, 129, 0.4);
    }

    .badge-txt,
    .badge-rtf {
        background: rgba(59, 130, 246, 0.2);
        color: #60A5FA;
        border: 1px solid rgba(59, 130, 246, 0.4);
    }

    .export-features {
        list-style: none;
        padding: 0;
        margin: 1.4rem 0;
    }

    .export-features li {
        font-size: 0.9rem;
        color: #94A3B8;
        margin-bottom: 0.6rem;
        display: flex;
        align-items: center;
        gap: 0.6rem;
    }

    .export-features li span.check {
        color: #34D399;
        font-weight: 700;
    }

    .filename-pill {
        background: rgba(15, 23, 42, 0.85);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 8px;
        padding: 0.45rem 0.9rem;
        font-family: monospace;
        font-size: 0.88rem;
        color: #93C5FD;
        display: inline-block;
        margin-bottom: 1.2rem;
    }

    /* Process Action Card */
    .process-action-card {
        background: linear-gradient(135deg, rgba(17, 24, 39, 0.85) 0%, rgba(30, 58, 138, 0.35) 100%);
        backdrop-filter: blur(16px);
        border: 1px solid rgba(59, 130, 246, 0.35);
        border-radius: 16px;
        padding: 1.6rem 2rem;
        margin-top: 1.6rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 1.5rem;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3), 0 0 25px rgba(37, 99, 235, 0.2);
    }

    .process-action-info h4 {
        margin: 0;
        font-size: 1.1rem;
        font-weight: 700;
        color: #93C5FD;
    }

    .process-action-info p {
        margin: 0.3rem 0 0 0;
        font-size: 0.9rem;
        color: #CBD5E1;
    }

    /* Sidebar modern cards */
    .sidebar-header-box {
        display: flex;
        align-items: center;
        gap: 0.9rem;
        padding-bottom: 1.2rem;
        border-bottom: 1px solid rgba(255, 255, 255, 0.08);
        margin-bottom: 1.2rem;
    }

    .sidebar-header-box h2 {
        margin: 0;
        font-size: 1.3rem;
        font-weight: 800;
        color: #F8FAFC;
    }

    .sidebar-section-title {
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #64748B;
        margin: 1.4rem 0 0.7rem 0;
    }

    .privacy-card {
        background: rgba(17, 24, 39, 0.85);
        border: 1px solid rgba(59, 130, 246, 0.3);
        border-left: 4px solid #3B82F6;
        border-radius: 0 12px 12px 0;
        padding: 1rem 1.1rem;
        font-size: 0.84rem;
        color: #CBD5E1;
        line-height: 1.45;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.25), 0 0 15px rgba(37, 99, 235, 0.1);
    }

    /* File Uploader Dark Styling */
    div[data-testid="stFileUploader"] {
        background: rgba(17, 24, 39, 0.6) !important;
        border: 2px dashed rgba(59, 130, 246, 0.35) !important;
        border-radius: 16px !important;
        padding: 1.2rem !important;
        backdrop-filter: blur(12px) !important;
        transition: all 0.25s ease !important;
    }

    div[data-testid="stFileUploader"]:hover {
        border-color: rgba(96, 165, 250, 0.6) !important;
        box-shadow: 0 0 25px rgba(59, 130, 246, 0.2) !important;
    }

    /* Dataframe Dark Elevation */
    div[data-testid="stDataFrame"] {
        border-radius: 14px !important;
        overflow: hidden !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        background: rgba(17, 24, 39, 0.65) !important;
        backdrop-filter: blur(16px) !important;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3) !important;
    }

    /* Primary & Secondary Buttons */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(147, 197, 253, 0.35) !important;
        box-shadow: 0 4px 16px rgba(37, 99, 235, 0.4), 0 0 20px rgba(37, 99, 235, 0.2) !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
        padding: 0.65rem 1.4rem !important;
        transition: all 0.2s ease !important;
    }

    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 24px rgba(37, 99, 235, 0.6), 0 0 30px rgba(59, 130, 246, 0.4) !important;
    }

    div.stButton > button[kind="secondary"] {
        background: rgba(30, 41, 59, 0.8) !important;
        color: #E2E8F0 !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        border-radius: 10px !important;
        transition: all 0.2s ease !important;
    }

    div.stButton > button[kind="secondary"]:hover {
        background: rgba(51, 65, 85, 0.9) !important;
        color: #FFFFFF !important;
        border-color: rgba(255, 255, 255, 0.25) !important;
        transform: translateY(-1px) !important;
    }

    div.stDownloadButton > button[kind="primary"] {
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
        border: 1px solid rgba(52, 211, 153, 0.35) !important;
        box-shadow: 0 4px 16px rgba(16, 185, 129, 0.35) !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
    }

    div.stDownloadButton > button[kind="primary"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 24px rgba(16, 185, 129, 0.5) !important;
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
    st.markdown(
        """
        <div class="sidebar-header-box">
            <div style="background: rgba(37, 99, 235, 0.2); border: 1px solid rgba(59, 130, 246, 0.4); border-radius: 12px; width: 46px; height: 46px; display: flex; align-items: center; justify-content: center; font-size: 24px; box-shadow: 0 0 15px rgba(37, 99, 235, 0.3);">
                ⚡
            </div>
            <div>
                <h2 style="margin: 0; font-size: 1.25rem; font-weight: 800; color: #F8FAFC;">Configuration</h2>
                <span style="font-size: 0.8rem; color: #94A3B8; font-weight: 500;">Aging Parameters & Rules</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    st.markdown('<div class="sidebar-section-title">⚙️ Processing Parameters</div>', unsafe_allow_html=True)
    
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
    st.markdown('<div class="sidebar-section-title">🔄 Session Controls</div>', unsafe_allow_html=True)
    if st.button("🔄 Reset Application", use_container_width=True, type="secondary"):
        reset_application()

    st.markdown("---")
    st.markdown(
        """
        <div class="privacy-card">
            <strong style="color: #60A5FA; display: flex; align-items: center; gap: 0.45rem; font-size: 0.9rem;">
                🔒 Zero-Disk Privacy
            </strong>
            <p style="margin: 0.35rem 0 0 0; color: #CBD5E1;">
                All processing is performed strictly in-memory. Uploaded data is never persisted or stored externally.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# Main Header
st.markdown(
    f"""
    <div class="app-header-container">
        <div class="app-header-content">
            <div class="app-header-left">
                <div class="app-header-icon">📊</div>
                <div class="app-title-group">
                    <h1>{Settings.APP_NAME} <span style="font-size: 0.82rem; font-weight: 700; background: rgba(59, 130, 246, 0.25); border: 1px solid rgba(96, 165, 250, 0.4); padding: 0.2rem 0.65rem; border-radius: 9999px; vertical-align: middle; margin-left: 0.4rem; color: #93C5FD; box-shadow: 0 0 10px rgba(59, 130, 246, 0.3);">v2.0</span></h1>
                    <p>{Settings.APP_SUBTITLE}</p>
                </div>
            </div>
            <div class="app-header-badges">
                <span class="header-chip"><span class="status-dot-green"></span> In-Memory Secure</span>
                <span class="header-chip"><span class="status-dot-blue"></span> Strict Aging: &gt; {threshold_days}d</span>
                <span class="header-chip">🗓️ Ref: {ref_date_option.strftime('%d-%b-%Y')}</span>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Workflow Tabs
tab_upload, tab_results, tab_download = st.tabs([
    "📁 1. Upload & Process",
    "📊 2. Results & Preview",
    "📥 3. Download Reports",
])


# ==========================================
# TAB 1: UPLOAD & PROCESS
# ==========================================
with tab_upload:
    st.markdown(
        """
        <div class="step-banner">
            <div class="step-card">
                <div class="step-num">1</div>
                <div class="step-text">
                    <h4>Upload Documents</h4>
                    <p>Drop multi-sheet Excel (.xlsx) and CSV files</p>
                </div>
            </div>
            <div class="step-card">
                <div class="step-num">2</div>
                <div class="step-text">
                    <h4>Auto-Detection</h4>
                    <p>Instant detection of ticket IDs, names, dates</p>
                </div>
            </div>
            <div class="step-card">
                <div class="step-num">3</div>
                <div class="step-text">
                    <h4>Filter & Consolidate</h4>
                    <p>Strict aging filter with separated outputs</p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
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
            # Ingest and auto-detect columns
            with st.spinner("Ingesting and inspecting files..."):
                parsed_files: List[FileData] = []
                for uf in uploaded_files:
                    f_data = FileReader.read_file(uf, uf.name)
                    parsed_files.append(f_data)

                st.session_state.files_data = parsed_files
                st.session_state.uploaded_file_keys = current_file_keys
                # Auto-run intelligent column detection
                st.session_state.sheet_configs = ProcessingManager.analyze_uploaded_files(parsed_files)
                st.session_state.processing_result = None

    if st.session_state.files_data:
        st.markdown("### 📋 Uploaded Files Overview")
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

        cutoff_date_str = (ref_date_option - datetime.timedelta(days=threshold_days)).strftime('%d-%b-%Y')
        st.markdown(
            f"""
            <div class="process-action-card">
                <div class="process-action-info">
                    <h4>🎯 Ready to Filter Aging Records</h4>
                    <p>
                        Condition: <strong>Last Updated older than {threshold_days} calendar days</strong> 
                        (strictly updated before <strong>{cutoff_date_str}</strong> relative to reference date <strong>{ref_date_option.strftime('%d-%b-%Y')}</strong>).
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
        process_btn = st.button("🚀 Run Aging Analysis & Filter Records", type="primary", use_container_width=True)

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
                st.success(f"Processing complete! Found {res.stats['total_matching']:,} matching record(s). Proceed to **Tab 2: Results & Preview**.")
    else:
        ui.alert(
            title="Ready for File Ingestion",
            description="Drag and drop your .xlsx or .csv files above to automatically detect columns and start aging analysis.",
            variant="default",
            key="alert_ready_upload",
        )


# ==========================================
# TAB 2: RESULTS & PREVIEW
# ==========================================
with tab_results:
    res: Optional[ProcessingResult] = st.session_state.processing_result

    if not res:
        ui.alert(
            title="Awaiting Processing",
            description="Upload your files and click '🚀 Run Aging Analysis & Filter Records' in Tab 1 to generate filtered results.",
            variant="default",
            key="alert_awaiting_processing",
        )
    else:
        st.subheader("Processing Summary & KPIs")

        kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
        with kpi1:
            st.html(
                f'<div class="kpi-box kpi-box-blue">'
                f'<div class="kpi-box-header"><span class="kpi-box-label" style="color: #93C5FD;">Files Processed</span><span class="kpi-box-icon">📁</span></div>'
                f'<div class="kpi-box-value" style="color: #60A5FA;">{res.stats["total_files"]}</div>'
                f'<p class="kpi-box-desc" style="color: #94A3B8;">Uploaded files</p>'
                f'</div>'
            )
        with kpi2:
            st.html(
                f'<div class="kpi-box kpi-box-cyan">'
                f'<div class="kpi-box-header"><span class="kpi-box-label" style="color: #67E8F9;">Sheets Processed</span><span class="kpi-box-icon">📑</span></div>'
                f'<div class="kpi-box-value" style="color: #22D3EE;">{res.stats["total_sheets"]}</div>'
                f'<p class="kpi-box-desc" style="color: #94A3B8;">Worksheets analyzed</p>'
                f'</div>'
            )
        with kpi3:
            st.html(
                f'<div class="kpi-box kpi-box-purple">'
                f'<div class="kpi-box-header"><span class="kpi-box-label" style="color: #D8B4FE;">Total Examined</span><span class="kpi-box-icon">🔍</span></div>'
                f'<div class="kpi-box-value" style="color: #C084FC;">{res.stats["total_examined"]:,}</div>'
                f'<p class="kpi-box-desc" style="color: #94A3B8;">Total data rows</p>'
                f'</div>'
            )
        with kpi4:
            st.html(
                f'<div class="kpi-box kpi-box-emerald">'
                f'<div class="kpi-box-header"><span class="kpi-box-label" style="color: #6EE7B7;">Matching Records</span><span class="kpi-box-icon">🎯</span></div>'
                f'<div class="kpi-box-value" style="color: #34D399;">{res.stats["total_matching"]:,}</div>'
                f'<p class="kpi-box-desc" style="color: #A7F3D0;">&gt; {res.stats["threshold_days"]} days aging</p>'
                f'</div>'
            )
        with kpi5:
            st.html(
                f'<div class="kpi-box kpi-box-amber">'
                f'<div class="kpi-box-header"><span class="kpi-box-label" style="color: #FDE68A;">Skipped Records</span><span class="kpi-box-icon">⏭️</span></div>'
                f'<div class="kpi-box-value" style="color: #FBBF24;">{res.stats["total_skipped"]:,}</div>'
                f'<p class="kpi-box-desc" style="color: #CBD5E1;">≤ {res.stats["threshold_days"]} days aging</p>'
                f'</div>'
            )
        with kpi6:
            st.html(
                f'<div class="kpi-box kpi-box-crimson">'
                f'<div class="kpi-box-header"><span class="kpi-box-label" style="color: #FECDD3;">Issues / Errors</span><span class="kpi-box-icon">⚠️</span></div>'
                f'<div class="kpi-box-value" style="color: #F87171;">{res.stats["total_errors"]:,}</div>'
                f'<p class="kpi-box-desc" style="color: #FDA4AF;">Validation errors</p>'
                f'</div>'
            )

        # Result Details Sub-tabs
        sub_matching, sub_skipped, sub_errors, sub_breakdown = st.tabs([
            f"✅ Matching Records ({len(res.matching_records)})",
            f"⏭️ Skipped Records ({len(res.skipped_records)})",
            f"⚠️ Processing Errors ({len(res.errors)})",
            "📋 File Breakdown",
        ])

        # 1. Matching Records
        with sub_matching:
            if not res.matching_records:
                ui.alert(
                    title="No Records Met Aging Condition",
                    description=f"All evaluated records were updated within the last {res.stats['threshold_days']} calendar days. You may lower the threshold in the sidebar to include more records.",
                    variant="default",
                    key="alert_no_matching",
                )
            else:
                inc_recs = [r for r in res.matching_records if r.incident]
                sctask_recs = [r for r in res.matching_records if r.sctask]
                other_recs = [r for r in res.matching_records if not r.incident and not r.sctask]

                search_query = st.text_input("🔍 Filter by Name, Ticket ID, Description, or Source File", "")

                def filter_recs(records_list, ticket_field=""):
                    if not search_query:
                        return records_list
                    q = search_query.lower()
                    filtered = []
                    for r in records_list:
                        ticket_val = getattr(r, ticket_field, "") if ticket_field else ""
                        if (
                            q in r.name.lower()
                            or q in ticket_val.lower()
                            or q in r.short_description.lower()
                            or q in r.source_file.lower()
                        ):
                            filtered.append(r)
                    return filtered

                if inc_recs:
                    filtered_inc = filter_recs(inc_recs, "incident")
                    st.markdown(
                        f"""
                        <div class="section-banner-inc">
                            <span>🎫 INCIDENTS</span>
                            <span class="banner-count-badge">{len(filtered_inc)} of {len(inc_recs)} Records</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    df_inc = pd.DataFrame([
                        {
                            "Incident": r.incident,
                            "Short Description": r.short_description,
                            "Name": r.name,
                            "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                            "Source File": r.source_file,
                        }
                        for r in filtered_inc
                    ])
                    st.dataframe(df_inc, use_container_width=True, hide_index=True)

                if sctask_recs:
                    filtered_sctask = filter_recs(sctask_recs, "sctask")
                    st.markdown(
                        f"""
                        <div class="section-banner-sctask">
                            <span>📋 SERVICE CATALOG TASKS (SCTASKS)</span>
                            <span class="banner-count-badge">{len(filtered_sctask)} of {len(sctask_recs)} Records</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    df_sctask = pd.DataFrame([
                        {
                            "SCTask": r.sctask,
                            "Short Description": r.short_description,
                            "Name": r.name,
                            "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                            "Source File": r.source_file,
                        }
                        for r in filtered_sctask
                    ])
                    st.dataframe(df_sctask, use_container_width=True, hide_index=True)

                if other_recs:
                    filtered_other = filter_recs(other_recs)
                    st.markdown(
                        f"""
                        <div class="section-banner-other">
                            <span>📄 OTHER MATCHING RECORDS</span>
                            <span class="banner-count-badge">{len(filtered_other)} of {len(other_recs)} Records</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    has_desc = any(r.short_description for r in other_recs)
                    rows = []
                    for r in filtered_other:
                        row = {}
                        if has_desc:
                            row["Short Description"] = r.short_description
                        row["Name"] = r.name
                        row["Last Updated"] = r.last_updated.strftime("%d-%b-%Y")
                        row["Source File"] = r.source_file
                        rows.append(row)
                    df_other = pd.DataFrame(rows)
                    st.dataframe(df_other, use_container_width=True, hide_index=True)

                st.caption(f"Showing matching records exceeding the {res.stats['threshold_days']}-day aging threshold.")

        # 2. Skipped Records
        with sub_skipped:
            if not res.skipped_records:
                ui.alert(
                    title="No Skipped Records",
                    description="All examined records met the aging threshold and were included in the filtered results.",
                    variant="default",
                    key="alert_no_skipped",
                )
            else:
                inc_skip = [r for r in res.skipped_records if r.incident]
                sctask_skip = [r for r in res.skipped_records if r.sctask]
                other_skip = [r for r in res.skipped_records if not r.incident and not r.sctask]

                if inc_skip:
                    st.markdown(
                        f"""
                        <div class="section-banner-inc">
                            <span>🎫 INCIDENTS (SKIPPED)</span>
                            <span class="banner-count-badge">{len(inc_skip)} Records</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    df_inc_skip = pd.DataFrame([
                        {
                            "Incident": r.incident,
                            "Short Description": r.short_description,
                            "Name": r.name,
                            "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                            "Source File": r.source_file,
                        }
                        for r in inc_skip
                    ])
                    st.dataframe(df_inc_skip, use_container_width=True, hide_index=True)

                if sctask_skip:
                    st.markdown(
                        f"""
                        <div class="section-banner-sctask">
                            <span>📋 SERVICE CATALOG TASKS (SKIPPED)</span>
                            <span class="banner-count-badge">{len(sctask_skip)} Records</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    df_sctask_skip = pd.DataFrame([
                        {
                            "SCTask": r.sctask,
                            "Short Description": r.short_description,
                            "Name": r.name,
                            "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                            "Source File": r.source_file,
                        }
                        for r in sctask_skip
                    ])
                    st.dataframe(df_sctask_skip, use_container_width=True, hide_index=True)

                if other_skip:
                    st.markdown(
                        f"""
                        <div class="section-banner-other">
                            <span>📄 OTHER RECORDS (SKIPPED)</span>
                            <span class="banner-count-badge">{len(other_skip)} Records</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    has_desc = any(r.short_description for r in other_skip)
                    rows = []
                    for r in other_skip:
                        row = {}
                        if has_desc:
                            row["Short Description"] = r.short_description
                        row["Name"] = r.name
                        row["Last Updated"] = r.last_updated.strftime("%d-%b-%Y")
                        row["Source File"] = r.source_file
                        rows.append(row)
                    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

                st.caption("These records were updated within the threshold (Days Difference ≤ Threshold) and therefore excluded.")

        # 3. Processing Errors
        with sub_errors:
            if not res.errors:
                ui.alert(
                    title="Zero Validation Errors",
                    description="All records across uploaded sheets were cleanly validated and parsed without format issues.",
                    variant="default",
                    key="alert_clean_errors",
                )
            else:
                ui.alert(
                    title=f"{len(res.errors)} Validation Issues Encountered",
                    description="Records with missing dates, invalid date values, missing names, or future dates were safely segregated.",
                    variant="destructive",
                    key="alert_has_errors",
                )
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
    if not processing_res.matching_records:
        st.info("No records met the filtering criteria.")
    else:
        inc_recs = [r for r in processing_res.matching_records if r.incident]
        sctask_recs = [r for r in processing_res.matching_records if r.sctask]
        other_recs = [r for r in processing_res.matching_records if not r.incident and not r.sctask]

        if inc_recs:
            st.markdown(
                f"""
                <div class="section-banner-inc">
                    <span>🎫 INCIDENTS</span>
                    <span class="banner-count-badge">{len(inc_recs)} Records</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            df_inc = pd.DataFrame([
                {
                    "Incident": r.incident,
                    "Short Description": r.short_description,
                    "Name": r.name,
                    "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                    "Source File": r.source_file,
                }
                for r in inc_recs
            ])
            st.dataframe(df_inc, use_container_width=True, hide_index=True)

        if sctask_recs:
            st.markdown(
                f"""
                <div class="section-banner-sctask">
                    <span>📋 SERVICE CATALOG TASKS (SCTASKS)</span>
                    <span class="banner-count-badge">{len(sctask_recs)} Records</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            df_sctask = pd.DataFrame([
                {
                    "SCTask": r.sctask,
                    "Short Description": r.short_description,
                    "Name": r.name,
                    "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                    "Source File": r.source_file,
                }
                for r in sctask_recs
            ])
            st.dataframe(df_sctask, use_container_width=True, hide_index=True)

        if other_recs:
            st.markdown(
                f"""
                <div class="section-banner-other">
                    <span>📄 OTHER MATCHING RECORDS</span>
                    <span class="banner-count-badge">{len(other_recs)} Records</span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            has_desc = any(r.short_description for r in other_recs)
            rows = []
            for r in other_recs:
                row = {}
                if has_desc:
                    row["Short Description"] = r.short_description
                row["Name"] = r.name
                row["Last Updated"] = r.last_updated.strftime("%d-%b-%Y")
                row["Source File"] = r.source_file
                rows.append(row)
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

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


@st.dialog("📝 RTF Output Preview", width="large")
def show_rtf_modal(processing_res: ProcessingResult):
    st.subheader(f"Rich Text Report ({DEFAULT_RTF_FILENAME})")
    st.caption(
        f"Total Matching Records: **{len(processing_res.matching_records)}** | "
        f"Rule: **Last Updated older than {processing_res.stats.get('threshold_days', 5)} calendar days** | "
        f"Format: **Native Word / WordPad Tables**"
    )

    inc_records = [r for r in processing_res.matching_records if r.incident]
    sctask_records = [r for r in processing_res.matching_records if r.sctask]
    other_records = [r for r in processing_res.matching_records if not r.incident and not r.sctask]

    if inc_records:
        st.markdown(
            f"""
            <div class="section-banner-inc">
                <span>🛡️ INCIDENTS</span>
                <span style="font-size: 0.85rem; font-weight: 600; opacity: 0.9;">{len(inc_records)} Records</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        df_inc = pd.DataFrame([
            {
                "Incident": r.incident,
                "Short Description": r.short_description,
                "Name": r.name,
                "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                "Source File": r.source_file,
            }
            for r in inc_records
        ])
        st.dataframe(df_inc, use_container_width=True, hide_index=True)

    if sctask_records:
        st.markdown(
            f"""
            <div class="section-banner-sctask">
                <span>⚙️ SERVICE CATALOG TASKS (SCTASKS)</span>
                <span style="font-size: 0.85rem; font-weight: 600; opacity: 0.9;">{len(sctask_records)} Records</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        df_sctask = pd.DataFrame([
            {
                "SCTask": r.sctask,
                "Short Description": r.short_description,
                "Name": r.name,
                "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                "Source File": r.source_file,
            }
            for r in sctask_records
        ])
        st.dataframe(df_sctask, use_container_width=True, hide_index=True)

    if other_records:
        st.markdown(
            f"""
            <div class="section-banner-other">
                <span>📋 OTHER MATCHING RECORDS</span>
                <span style="font-size: 0.85rem; font-weight: 600; opacity: 0.9;">{len(other_records)} Records</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        has_desc = any(r.short_description for r in other_records)
        df_other = pd.DataFrame([
            (
                {
                    "Short Description": r.short_description,
                    "Name": r.name,
                    "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                    "Source File": r.source_file,
                }
                if has_desc
                else {
                    "Name": r.name,
                    "Last Updated": r.last_updated.strftime("%d-%b-%Y"),
                    "Source File": r.source_file,
                }
            )
            for r in other_records
        ])
        st.dataframe(df_other, use_container_width=True, hide_index=True)

    st.markdown("---")
    rtf_data = getattr(processing_res, "rtf_report", None)
    if not rtf_data and processing_res.matching_records:
        import modules.export_generator
        if not hasattr(modules.export_generator.ExportGenerator, "generate_rtf_report"):
            importlib.reload(modules.export_generator)
        from modules.export_generator import ExportGenerator
        rtf_data = ExportGenerator.generate_rtf_report(
            processing_res.matching_records,
            processing_res.stats.get("processing_date", get_current_date()),
            processing_res.stats.get("threshold_days", 5),
        )
        try:
            processing_res.rtf_report = rtf_data
        except Exception:
            pass

    st.download_button(
        label=f"📥 Download RTF ({len(processing_res.matching_records)} records)",
        data=rtf_data or "",
        file_name=DEFAULT_RTF_FILENAME,
        mime="application/rtf",
        type="secondary",
        key="modal_btn_dl_rtf",
        use_container_width=True,
    )


# ==========================================
# TAB 3: DOWNLOAD REPORTS
# ==========================================
with tab_download:
    res: Optional[ProcessingResult] = st.session_state.processing_result

    if not res:
        ui.alert(
            title="Reports Awaiting Processing",
            description="Please upload and process your files in Tab 1 before previewing or downloading reports.",
            variant="default",
            key="alert_awaiting_download",
        )
    else:
        st.subheader("Download Consolidated Reports")
        st.markdown(
            "Download your filtered data in **Excel (.xlsx)** and **Rich Text (.rtf)** formats, "
            "or click **👁️ View Output** to preview each report in a popup modal."
        )

        dl_col1, dl_col2 = st.columns(2)

        # Excel Download Card
        with dl_col1:
            st.markdown(
                f"""
                <div class="export-card export-card-excel">
                    <div>
                        <div class="export-header">
                            <div class="export-icon-box export-icon-excel">📊</div>
                            <div class="export-title-group">
                                <h3>Excel Spreadsheet</h3>
                                <span class="export-badge badge-excel">.XLSX FORMAT</span>
                            </div>
                        </div>
                        <p style="color: #64748B; font-size: 0.9rem; line-height: 1.4; margin: 0.4rem 0;">
                            Consolidated multi-table spreadsheet with formatted headers and clear separation between Incidents and SCTasks.
                        </p>
                        <ul class="export-features">
                            <li><span class="check">✓</span> Separated Incidents and SCTasks tables</li>
                            <li><span class="check">✓</span> Short Description & Source File tracked</li>
                            <li><span class="check">✓</span> Bold styled headers & auto-fit column widths</li>
                            <li><span class="check">✓</span> Formula injection sanitized (OWASP safe)</li>
                        </ul>
                        <div class="filename-pill">📁 {Settings.DEFAULT_EXCEL_FILENAME}</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
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

        # RTF Download Card
        with dl_col2:
            st.markdown(
                f"""
                <div class="export-card export-card-rtf">
                    <div>
                        <div class="export-header">
                            <div class="export-icon-box export-icon-rtf">📝</div>
                            <div class="export-title-group">
                                <h3>Rich Text Document</h3>
                                <span class="export-badge badge-rtf">.RTF FORMAT</span>
                            </div>
                        </div>
                        <p style="color: #64748B; font-size: 0.9rem; line-height: 1.4; margin: 0.4rem 0;">
                            Formatted document with native Word tables, styled headers, and clean separation between Incidents and SCTasks.
                        </p>
                        <ul class="export-features">
                            <li><span class="check">✓</span> Native Word & WordPad table grids</li>
                            <li><span class="check">✓</span> Never misaligns with editor fonts</li>
                            <li><span class="check">✓</span> Clean in-cell automatic text wrapping</li>
                            <li><span class="check">✓</span> Distinct Incidents & SCTask sections</li>
                        </ul>
                        <div class="filename-pill">📁 {DEFAULT_RTF_FILENAME}</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

            rtf_download_content = getattr(res, "rtf_report", None)
            if not rtf_download_content and res.matching_records:
                import modules.export_generator
                if not hasattr(modules.export_generator.ExportGenerator, "generate_rtf_report"):
                    importlib.reload(modules.export_generator)
                from modules.export_generator import ExportGenerator
                rtf_download_content = ExportGenerator.generate_rtf_report(
                    res.matching_records,
                    res.stats.get("processing_date", get_current_date()),
                    res.stats.get("threshold_days", 5),
                )
                try:
                    res.rtf_report = rtf_download_content
                except Exception:
                    pass

            b2_col1, b2_col2 = st.columns(2)
            with b2_col1:
                if st.button("👁️ View Output", key="btn_view_rtf", use_container_width=True):
                    show_rtf_modal(res)
            with b2_col2:
                st.download_button(
                    label=f"📥 Download RTF ({len(res.matching_records)})",
                    data=rtf_download_content or "",
                    file_name=DEFAULT_RTF_FILENAME,
                    mime="application/rtf",
                    type="secondary",
                    use_container_width=True,
                    key="btn_dl_rtf",
                )

        st.markdown("---")
        st.button("🔄 Start New Processing Session", on_click=reset_application, type="secondary")
