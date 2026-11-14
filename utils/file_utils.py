"""File manipulation, sanitization, and security utilities."""
import os
import re
from pathlib import Path
from typing import Any


def get_safe_filename(filename: str) -> str:
    """Sanitize filename to prevent directory traversal or invalid filesystem characters."""
    if not filename:
        return "unnamed_file"
    # Take only the base name (remove directory components)
    base = Path(filename).name
    # Remove potentially dangerous characters
    safe = re.sub(r'[^a-zA-Z0-9_\-\. ]', '_', base)
    return safe.strip() or "unnamed_file"


def is_valid_extension(filename: str, supported_extensions: set = None) -> bool:
    """Check if the filename has a supported extension (.xlsx or .csv)."""
    if supported_extensions is None:
        supported_extensions = {".xlsx", ".csv"}
    ext = Path(filename).suffix.lower()
    return ext in supported_extensions


def format_file_size(size_bytes: int) -> str:
    """Format file size in human-readable units (Bytes, KB, MB, GB)."""
    if size_bytes < 0:
        return "0 B"
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}" if unit != 'B' else f"{size_bytes} B"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"


def sanitize_cell_for_export(val: Any) -> Any:
    """Sanitize untrusted text to prevent Excel formula injection (CSV/Excel injection).
    
    If a string starts with '=', '+', '-', '@', '\t', or '\r', it could trigger formula execution
    when opened in spreadsheet applications. Prefix with a single quote to force text interpretation.
    """
    if isinstance(val, str):
        if val and val[0] in ('=', '+', '-', '@', '\t', '\r'):
            # Check if it's purely a negative number like -5 or -12.34
            try:
                float(val)
                return val  # Safe numeric value
            except ValueError:
                # Untrusted formula-like text; escape with single quote
                return f"'{val}"
    return val
