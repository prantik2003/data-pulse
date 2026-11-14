"""Header normalization utilities for column matching."""
import re
from typing import Set


def normalize_header(header: str) -> str:
    """Normalize a column header for semantic and fuzzy matching.
    
    Handles:
    - Case insensitivity (lowercased)
    - CamelCase splitting (e.g. employeeName -> employee Name)
    - Underscores, hyphens, slashes, periods, brackets replaced with spaces
    - Punctuation removal
    - Multi-space collapsing
    - Leading and trailing whitespace stripping
    
    Examples:
        'Employee_Name' -> 'employee name'
        'EMPLOYEE NAME' -> 'employee name'
        'employeeName' -> 'employee name'
        'Employee-Name' -> 'employee name'
        'Last_Updated_Date (UTC)' -> 'last updated date utc'
    """
    if not header or not isinstance(header, str):
        return ""
    
    text = header.strip()
    
    # Split CamelCase: 'employeeName' -> 'employee Name', 'XMLData' -> 'XML Data'
    # 1. Lowercase followed by Uppercase (e.g., 'employeeName' -> 'employee Name')
    text = re.sub(r'([a-z0-9])([A-Z])', r'\1 \2', text)
    # 2. Uppercase sequence followed by Uppercase + Lowercase (e.g., 'EMPLOYEEName' -> 'EMPLOYEE Name')
    text = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', text)
    
    # Replace separators and punctuation with space
    # Includes: _, -, ., /, \, |, [, ], (, ), {, }, :, ;, ,, #, @, *, +
    text = re.sub(r'[\_\-\.\/\\\|\[\]\(\)\{\}\:\;\,#@\*\+\?\!\<\>\=\^\~`"\']+', ' ', text)
    
    # Convert to lowercase
    text = text.lower()
    
    # Collapse multiple whitespaces into a single space
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text


def get_header_tokens(header: str) -> Set[str]:
    """Extract individual words/tokens from a normalized header."""
    normalized = normalize_header(header)
    if not normalized:
        return set()
    return set(normalized.split())
