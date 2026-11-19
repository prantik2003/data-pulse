"""Intelligent Column Detection using fuzzy matching and content-based validation."""
import re
from dataclasses import dataclass, field
from typing import Any, List, Optional, Tuple

import pandas as pd
from rapidfuzz import fuzz

from config.settings import Settings
from modules.date_parser import DateParser
from utils.header_normalizer import get_header_tokens, normalize_header


@dataclass
class ColumnCandidate:
    """Represents an evaluated column candidate with detailed scoring."""
    
    column_name: str
    combined_score: float
    header_score: float
    content_score: float
    notes: str = ""


@dataclass
class ColumnDetectionResult:
    """Structured detection outcome for a single tabular dataset / worksheet."""
    
    detected_name_col: Optional[str] = None
    name_confidence: float = 0.0
    name_status: str = "unmatched"  # 'high', 'medium', 'low', 'unmatched'
    name_candidates: List[ColumnCandidate] = field(default_factory=list)

    detected_date_col: Optional[str] = None
    date_confidence: float = 0.0
    date_status: str = "unmatched"  # 'high', 'medium', 'low', 'unmatched'
    date_candidates: List[ColumnCandidate] = field(default_factory=list)

    detected_incident_col: Optional[str] = None
    detected_sctask_col: Optional[str] = None
    detected_short_desc_col: Optional[str] = None

    needs_confirmation: bool = False
    confirmation_reason: Optional[str] = None


class ColumnDetector:
    """Intelligently detects Name and Last Updated columns from DataFrame headers and content."""

    @classmethod
    def detect_columns(
        cls,
        df: pd.DataFrame,
        sample_size: int = 50,
        file_name: str = "",
        sheet_name: str = "",
    ) -> ColumnDetectionResult:
        """Analyze DataFrame columns using both semantic header matching and value inspection.
        
        Args:
            df: The worksheet DataFrame.
            sample_size: Number of non-empty rows to sample for content validation.
            file_name: Optional file name for contextual disambiguation.
            sheet_name: Optional sheet name for contextual disambiguation.
            
        Returns:
            ColumnDetectionResult containing top detected columns, confidence levels, and candidate rankings.
        """
        if df.empty or len(df.columns) == 0:
            return ColumnDetectionResult(
                detected_name_col=None,
                name_confidence=0.0,
                name_status="unmatched",
                detected_date_col=None,
                date_confidence=0.0,
                date_status="unmatched",
                detected_incident_col=None,
                detected_sctask_col=None,
                detected_short_desc_col=None,
                needs_confirmation=True,
                confirmation_reason="The sheet contains no data or columns.",
            )

        columns = list(df.columns)
        
        name_candidates: List[ColumnCandidate] = []
        date_candidates: List[ColumnCandidate] = []

        # Evaluate each column for both Name and Date suitability
        for col in columns:
            col_str = str(col)
            sample_values = cls._get_clean_sample(df[col], sample_size)

            # Evaluate as Name column
            name_cand = cls._evaluate_name_column(col_str, sample_values)
            name_candidates.append(name_cand)

            # Evaluate as Date column
            date_cand = cls._evaluate_date_column(col_str, sample_values)
            date_candidates.append(date_cand)

        # Sort candidates descending by combined score
        name_candidates.sort(key=lambda c: c.combined_score, reverse=True)
        date_candidates.sort(key=lambda c: c.combined_score, reverse=True)

        best_name = name_candidates[0] if name_candidates and name_candidates[0].combined_score > 0.3 else None
        best_date = date_candidates[0] if date_candidates and date_candidates[0].combined_score > 0.3 else None

        # Prevent picking the exact same column for both Name and Date
        if best_name and best_date and best_name.column_name == best_date.column_name:
            if best_date.combined_score >= best_name.combined_score:
                # Give column to Date, pick 2nd candidate for Name
                best_name = name_candidates[1] if len(name_candidates) > 1 and name_candidates[1].combined_score > 0.3 else None
            else:
                # Give column to Name, pick 2nd candidate for Date
                best_date = date_candidates[1] if len(date_candidates) > 1 and date_candidates[1].combined_score > 0.3 else None

        # Determine confidence status for Name
        name_status, name_conf = cls._get_status_and_conf(best_name, name_candidates)
        
        # Determine confidence status for Date
        date_status, date_conf = cls._get_status_and_conf(best_date, date_candidates)

        # Check if user confirmation is required
        needs_confirmation = False
        reasons: List[str] = []

        if name_status in ("medium", "low", "unmatched"):
            needs_confirmation = True
            reasons.append(f"Name column detection confidence is {name_status.upper()}.")

        if date_status in ("medium", "low", "unmatched"):
            needs_confirmation = True
            reasons.append(f"Date column detection confidence is {date_status.upper()}.")

        # Check for ambiguity between top candidates (within 0.05)
        if len(name_candidates) >= 2 and name_candidates[0].combined_score > 0.5:
            delta = name_candidates[0].combined_score - name_candidates[1].combined_score
            if delta < Settings.CONFIDENCE_AMBIGUITY_DELTA:
                needs_confirmation = True
                reasons.append(f"Multiple plausible Name columns found ('{name_candidates[0].column_name}' and '{name_candidates[1].column_name}').")

        if len(date_candidates) >= 2 and date_candidates[0].combined_score > 0.5:
            delta = date_candidates[0].combined_score - date_candidates[1].combined_score
            if delta < Settings.CONFIDENCE_AMBIGUITY_DELTA:
                needs_confirmation = True
                reasons.append(f"Multiple plausible Date columns found ('{date_candidates[0].column_name}' and '{date_candidates[1].column_name}').")

        confirmation_reason = " | ".join(reasons) if reasons else None

        # Detect Short Description Column
        detected_short_desc_col: Optional[str] = None
        best_desc_score = 0.0
        for col in columns:
            col_str = str(col)
            if best_name and col_str == best_name.column_name:
                continue
            if best_date and col_str == best_date.column_name:
                continue
            norm = normalize_header(col_str)
            score = 0.0
            if norm in Settings.SHORT_DESCRIPTION_KEYWORDS:
                score = 1.0
            elif "description" in norm or "desc" in norm or "summary" in norm or "subject" in norm:
                score = 0.85
            elif any(fuzz.partial_ratio(norm, kw) >= 85 for kw in Settings.SHORT_DESCRIPTION_KEYWORDS):
                score = 0.70

            if score > best_desc_score:
                best_desc_score = score
                detected_short_desc_col = col_str

        # Detect Incident and SCTask Columns
        detected_incident_col: Optional[str] = None
        detected_sctask_col: Optional[str] = None
        file_sheet_context = f"{file_name} {sheet_name}".lower()
        is_inc_context = any(term in file_sheet_context for term in ["incident", "inc"])
        is_sctask_context = any(term in file_sheet_context for term in ["sctask", "sc_task", "sc task", "task"])

        for col in columns:
            col_str = str(col)
            if best_name and col_str == best_name.column_name:
                continue
            if best_date and col_str == best_date.column_name:
                continue
            if detected_short_desc_col and col_str == detected_short_desc_col:
                continue

            norm = normalize_header(col_str)
            sample_vals = [str(v).strip() for v in cls._get_clean_sample(df[col], 30)]

            has_inc_values = any(v.upper().startswith("INC") for v in sample_vals)
            has_sctask_values = any(v.upper().startswith("SCTASK") or v.upper().startswith("TASK") for v in sample_vals)

            if norm in Settings.INCIDENT_KEYWORDS or has_inc_values:
                if detected_incident_col is None:
                    detected_incident_col = col_str
            elif norm in Settings.SCTASK_KEYWORDS or has_sctask_values:
                if detected_sctask_col is None:
                    detected_sctask_col = col_str
            elif norm in ("number", "ticket", "ticket number", "ticket id", "id", "ticket no", "num", "task"):
                if has_inc_values:
                    if detected_incident_col is None:
                        detected_incident_col = col_str
                elif has_sctask_values:
                    if detected_sctask_col is None:
                        detected_sctask_col = col_str
                elif is_inc_context:
                    if detected_incident_col is None:
                        detected_incident_col = col_str
                elif is_sctask_context:
                    if detected_sctask_col is None:
                        detected_sctask_col = col_str

        return ColumnDetectionResult(
            detected_name_col=best_name.column_name if best_name else None,
            name_confidence=name_conf,
            name_status=name_status,
            name_candidates=name_candidates,
            detected_date_col=best_date.column_name if best_date else None,
            date_confidence=date_conf,
            date_status=date_status,
            date_candidates=date_candidates,
            detected_incident_col=detected_incident_col,
            detected_sctask_col=detected_sctask_col,
            detected_short_desc_col=detected_short_desc_col,
            needs_confirmation=needs_confirmation,
            confirmation_reason=confirmation_reason,
        )

    @classmethod
    def _get_status_and_conf(
        cls,
        best_cand: Optional[ColumnCandidate],
        candidates: List[ColumnCandidate],
    ) -> Tuple[str, float]:
        """Classify candidate into High, Medium, Low, or Unmatched."""
        if not best_cand or best_cand.combined_score < 0.35:
            return "unmatched", 0.0

        score = best_cand.combined_score
        if score >= Settings.CONFIDENCE_HIGH:
            return "high", round(score, 2)
        elif score >= Settings.CONFIDENCE_MEDIUM:
            return "medium", round(score, 2)
        else:
            return "low", round(score, 2)

    @classmethod
    def _get_clean_sample(cls, series: pd.Series, sample_size: int) -> List[Any]:
        """Extract a sample of non-empty, non-null values from a pandas series."""
        clean_series = series.dropna()
        if clean_series.empty:
            return []
        
        # Take up to sample_size values
        if len(clean_series) > sample_size:
            # Sample evenly across the column
            indices = [int(i * len(clean_series) / sample_size) for i in range(sample_size)]
            return list(clean_series.iloc[indices])
        return list(clean_series)

    @classmethod
    def _evaluate_name_column(cls, col_name: str, sample_values: List[Any]) -> ColumnCandidate:
        """Score a column as a potential Name column."""
        normalized = normalize_header(col_name)
        header_score = cls._score_name_header(normalized)
        content_score, notes = cls._score_name_content(sample_values)

        # If content cannot be names at all (e.g. 0% text, pure dates, or pure numbers), invalidate
        if content_score <= 0.05:
            combined_score = 0.0
        elif header_score >= 0.95 and content_score >= 0.5:
            combined_score = 0.55 * header_score + 0.45 * content_score
        else:
            combined_score = 0.50 * header_score + 0.50 * content_score

        return ColumnCandidate(
            column_name=col_name,
            combined_score=round(combined_score, 4),
            header_score=round(header_score, 4),
            content_score=round(content_score, 4),
            notes=notes,
        )

    @classmethod
    def _score_name_header(cls, normalized: str) -> float:
        """Calculate header match score for name keywords."""
        if not normalized:
            return 0.0

        # Exact match with primary name keywords
        if normalized in Settings.NAME_KEYWORDS:
            return 1.0

        # Check disfavored keywords (e.g. company, file, sheet, comment, date)
        tokens = set(normalized.split())
        for disfavored in Settings.DISFAVORED_NAME_KEYWORDS:
            if disfavored in tokens:
                # If it's something like "file name" or "company name", significantly penalize
                return 0.15

        # Check partial / token similarity against NAME_KEYWORDS
        best_fuzzy = 0.0
        for kw in Settings.NAME_KEYWORDS:
            ratio = fuzz.ratio(normalized, kw) / 100.0
            token_sort = fuzz.token_sort_ratio(normalized, kw) / 100.0
            token_set = fuzz.token_set_ratio(normalized, kw) / 100.0
            kw_score = max(ratio, token_sort, token_set)
            if kw_score > best_fuzzy:
                best_fuzzy = kw_score

        return best_fuzzy

    @classmethod
    def _score_name_content(cls, sample_values: List[Any]) -> Tuple[float, str]:
        """Validate sample values for human name characteristics."""
        if not sample_values:
            return 0.3, "Empty sample"

        valid_name_count = 0
        total = len(sample_values)
        long_comment_count = 0
        date_like_count = 0
        numeric_count = 0

        for val in sample_values:
            # Check if value is a date
            date_res = DateParser.parse_date(val)
            if date_res.success:
                date_like_count += 1
                continue

            # Check if numeric
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                numeric_count += 1
                continue

            s = str(val).strip()
            if not s:
                continue

            # Check if numeric string
            if re.match(r'^\d+(\.\d+)?$', s):
                numeric_count += 1
                continue

            # Check if excessively long (e.g. comment/description)
            if len(s) > 70 or s.count(" ") > 6:
                long_comment_count += 1
                continue

            # Check if reasonable person name (e.g. 2-50 chars, letters, maybe space, hyphen, period)
            if re.search(r'[a-zA-Z]', s):
                valid_name_count += 1

        # If majority are dates, this cannot be a name column
        if date_like_count / total > 0.4:
            return 0.0, "Contains date values"

        # If majority are numbers, this is not a name column
        if numeric_count / total > 0.5:
            return 0.0, "Contains numerical IDs"

        # If majority are long comments/sentences
        if long_comment_count / total > 0.3:
            return 0.1, "Contains long text/comments"

        ratio = valid_name_count / total
        return ratio, f"Valid name text in {valid_name_count}/{total} sample values"

    @classmethod
    def _evaluate_date_column(cls, col_name: str, sample_values: List[Any]) -> ColumnCandidate:
        """Score a column as a potential Last Updated Date column."""
        normalized = normalize_header(col_name)
        header_score = cls._score_date_header(normalized)
        allow_serial = (header_score >= 0.60)
        content_score, notes = cls._score_date_content(sample_values, allow_numeric_serial=allow_serial)

        # If content has 0% valid dates, it CANNOT be a date column even if header contains "date"
        if content_score == 0.0:
            combined_score = 0.0
        elif header_score >= 0.90 and content_score >= 0.6:
            combined_score = 0.55 * header_score + 0.45 * content_score
        else:
            combined_score = 0.45 * header_score + 0.55 * content_score

        return ColumnCandidate(
            column_name=col_name,
            combined_score=round(combined_score, 4),
            header_score=round(header_score, 4),
            content_score=round(content_score, 4),
            notes=notes,
        )

    @classmethod
    def _score_date_header(cls, normalized: str) -> float:
        """Calculate header match score for date keywords."""
        if not normalized:
            return 0.0

        # Disfavored financial, numeric, and ID keywords
        tokens = set(normalized.split())
        disfavored_numeric_terms = {
            "salary", "amount", "price", "cost", "total", "revenue", "income",
            "age", "phone", "mobile", "zip", "postal", "id", "identifier",
            "code", "count", "quantity", "qty", "num", "number", "rate",
            "balance", "credit", "debit", "currency",
        }
        if any(term in tokens for term in disfavored_numeric_terms):
            return 0.05

        # Exact match with primary "last updated" keywords
        if normalized in Settings.DATE_KEYWORDS:
            return 1.0

        # Exact match with generic date keywords (e.g. 'date', 'updated', 'modified')
        if normalized in Settings.GENERIC_DATE_KEYWORDS:
            return 0.80

        # Check disfavored date keywords (e.g. created at, birth date, start date, due date)
        for disfavored in Settings.DISFAVORED_DATE_KEYWORDS:
            if disfavored in normalized:
                # Heavily downweight so 'last updated' is preferred over 'created at'
                return 0.35

        # Check fuzzy similarity with DATE_KEYWORDS
        best_fuzzy = 0.0
        for kw in Settings.DATE_KEYWORDS:
            ratio = fuzz.ratio(normalized, kw) / 100.0
            token_sort = fuzz.token_sort_ratio(normalized, kw) / 100.0
            token_set = fuzz.token_set_ratio(normalized, kw) / 100.0
            kw_score = max(ratio, token_sort, token_set)
            if kw_score > best_fuzzy:
                best_fuzzy = kw_score

        # Also check fuzzy with generic keywords
        for kw in Settings.GENERIC_DATE_KEYWORDS:
            ratio = fuzz.ratio(normalized, kw) / 100.0
            token_set = fuzz.token_set_ratio(normalized, kw) / 100.0
            kw_score = max(ratio, token_set) * 0.80
            if kw_score > best_fuzzy:
                best_fuzzy = kw_score

        return best_fuzzy

    @classmethod
    def _score_date_content(
        cls,
        sample_values: List[Any],
        allow_numeric_serial: bool = False,
    ) -> Tuple[float, str]:
        """Validate sample values to test date parseability."""
        if not sample_values:
            return 0.0, "Empty sample"

        valid_ratio, valid_count, total = DateParser.is_likely_date_column(
            sample_values, allow_numeric_serial=allow_numeric_serial
        )
        if total == 0:
            return 0.0, "No non-empty values"

        if valid_ratio >= 0.80:
            return 1.0, f"High date parseability ({valid_count}/{total} valid)"
        elif valid_ratio >= 0.50:
            return valid_ratio, f"Moderate date parseability ({valid_count}/{total} valid)"
        elif valid_ratio > 0.0:
            return valid_ratio * 0.5, f"Low date parseability ({valid_count}/{total} valid)"
        else:
            return 0.0, "No parseable date values found"
