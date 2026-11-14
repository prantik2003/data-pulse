"""Application settings and constants for DataPulse."""
from dataclasses import dataclass, field
from typing import List, Set


@dataclass(frozen=True)
class AppConfig:
    """Core configuration for DataPulse."""
    
    APP_NAME: str = "DataPulse"
    APP_SUBTITLE: str = "Intelligent Excel & CSV Processing | Automated Record Filtering"
    APP_VERSION: str = "1.0.0"
    
    # Timezone & Date Configuration
    DEFAULT_TIMEZONE: str = "Asia/Kolkata"
    SUPPORTED_TIMEZONES: List[str] = field(default_factory=lambda: [
        "Asia/Kolkata",
        "UTC",
        "US/Eastern",
        "US/Central",
        "US/Pacific",
        "Europe/London",
        "Europe/Paris",
        "Asia/Dubai",
        "Asia/Singapore",
        "Asia/Tokyo",
        "Australia/Sydney",
    ])
    
    # Threshold Configuration
    DEFAULT_THRESHOLD_DAYS: int = 5
    MIN_THRESHOLD_DAYS: int = 0
    MAX_THRESHOLD_DAYS: int = 365
    
    # File Configuration
    SUPPORTED_EXTENSIONS: Set[str] = field(default_factory=lambda: {".xlsx", ".csv"})
    MAX_FILE_SIZE_MB: int = 50
    SAMPLE_CONTENT_SIZE: int = 60
    
    # Date preference options
    DATE_PREFERENCES: List[str] = field(default_factory=lambda: [
        "Auto-detect",
        "DD/MM/YYYY",
        "MM/DD/YYYY",
        "YYYY-MM-DD",
    ])
    
    # Confidence Score Thresholds
    CONFIDENCE_HIGH: float = 0.85
    CONFIDENCE_MEDIUM: float = 0.60
    CONFIDENCE_AMBIGUITY_DELTA: float = 0.05
    
    # Comprehensive Name Keywords (normalized form)
    NAME_KEYWORDS: List[str] = field(default_factory=lambda: [
        "name",
        "full name",
        "employee name",
        "user name",
        "person name",
        "employee",
        "resource",
        "resource name",
        "assigned to",
        "owner",
        "username",
        "contact name",
        "requester",
        "modified by",
        "staff name",
        "team member",
        "member name",
        "author",
        "editor",
        "agent",
        "worker",
        "contractor",
        "assignee",
        "creator",
        "submitted by",
        "display name",
        "customer name",
        "client name",
        "individual name",
    ])
    
    # Disfavored Name Keywords (columns that may contain "name" or text but are not person names)
    DISFAVORED_NAME_KEYWORDS: List[str] = field(default_factory=lambda: [
        "company",
        "organization",
        "org",
        "file",
        "filename",
        "sheet",
        "column",
        "table",
        "database",
        "project",
        "task",
        "ticket",
        "product",
        "item",
        "status",
        "date",
        "id",
        "identifier",
        "code",
        "comment",
        "description",
        "note",
        "notes",
        "details",
        "address",
        "email",
        "phone",
        "role",
        "department",
    ])
    
    # Comprehensive Date Keywords (normalized form)
    DATE_KEYWORDS: List[str] = field(default_factory=lambda: [
        "last updated",
        "updated at",
        "updated on",
        "last modified",
        "modified date",
        "modified on",
        "modification date",
        "modification timestamp",
        "last activity",
        "last activity date",
        "updated date",
        "date modified",
        "last change",
        "last changed on",
        "last changed date",
        "record updated at",
        "update date",
        "modify date",
        "change date",
        "latest update",
        "latest modified",
        "last seen",
        "last edit",
        "last edited",
        "last edit date",
        "last updated date",
        "last modification",
        "update timestamp",
        "modification time",
        "updated time",
        "last modified timestamp",
    ])
    
    # Generic Date Keywords (fallback when no specific "last updated" header exists)
    GENERIC_DATE_KEYWORDS: List[str] = field(default_factory=lambda: [
        "date",
        "timestamp",
        "updated",
        "modified",
        "last date",
        "activity date",
    ])
    
    # Disfavored Date Keywords (creation, birth, start, due dates - prioritized LOWER than updated/modified dates)
    DISFAVORED_DATE_KEYWORDS: List[str] = field(default_factory=lambda: [
        "created at",
        "created on",
        "creation date",
        "birth date",
        "dob",
        "start date",
        "end date",
        "due date",
        "expiry date",
        "expiration date",
        "joined date",
        "hire date",
        "deleted at",
    ])
    
    # Default Export File Names
    DEFAULT_EXCEL_FILENAME: str = "filtered_records.xlsx"
    DEFAULT_TXT_FILENAME: str = "filtered_records.txt"


# Singleton instance
Settings = AppConfig()
