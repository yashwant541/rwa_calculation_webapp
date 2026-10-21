"""Split an input dataset into filtered sub-datasets (one workbook / one sheet each).

A split rule is ``{"name": "COLLATERIZED", "column": "TRANCHE_COLLATERAL_TYPE",
"value": "COLLATERIZED"}``: keep the rows whose ``column`` equals ``value``. Matching is
case-insensitive on whitespace-collapsed text, and is *equality* (so ``COLLATERIZED`` does
not pick up ``NOT COLLATERIZED``). The rules are independent filters, so a row can land in
more than one dataset (e.g. collateralised and guaranteed).

Column names are expected to be standardised already (the backend's
``standardize_column_name`` upper-snake form); this module has no Dataiku dependency.
"""

from __future__ import annotations

import io
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

DEFAULT_SPLIT_RULES: List[Dict[str, str]] = [
    {"name": "COLLATERIZED", "column": "TRANCHE_COLLATERAL_TYPE", "value": "COLLATERIZED"},
    {"name": "NOT COLLATERIZED", "column": "TRANCHE_COLLATERAL_TYPE", "value": "NOT COLLATERIZED"},
    {"name": "GUARANTEED", "column": "TRANCHE_GUARANTEE_TYPE", "value": "GUARANTEED"},
]


def norm_text(value: Any) -> Optional[str]:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text.upper() if text else None


def clean_rules(raw_rules: Any, std_col=lambda c: str(c).strip()) -> List[Dict[str, str]]:
    """Validate/normalise user rules. Raises ValueError with a readable message."""
    rules = []
    for i, raw in enumerate(raw_rules or [], 1):
        column = std_col(raw.get("column", ""))
        value = str(raw.get("value", "") if raw.get("value") is not None else "").strip()
        name = str(raw.get("name") or value or "SPLIT_%d" % i).strip()
        if not column:
            raise ValueError("Split rule %d (%s): pick a column." % (i, name))
        if not value:
            raise ValueError("Split rule %d (%s): enter the value to keep." % (i, name))
        rules.append({"name": name, "column": column, "value": value})
    if not rules:
        raise ValueError("No split rules defined.")
    names = [safe_name(r["name"]) for r in rules]
    if len(set(names)) != len(names):
        raise ValueError("Split rule names must be unique.")
    return rules


def safe_name(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", str(text)).strip("_").upper() or "SPLIT"


def filter_frame(frame: pd.DataFrame, rule: Dict[str, str]) -> pd.DataFrame:
    target = norm_text(rule["value"])
    mask = frame[rule["column"]].map(norm_text) == target
    return frame[mask.fillna(False)]


def missing_columns(frame: pd.DataFrame, rules: List[Dict[str, str]]) -> List[str]:
    return sorted({r["column"] for r in rules if r["column"] not in frame.columns})


def split_frame(frame: pd.DataFrame, rules: List[Dict[str, str]]
                ) -> List[Tuple[Dict[str, str], pd.DataFrame]]:
    gone = missing_columns(frame, rules)
    if gone:
        raise ValueError("Split column(s) not found in the file: %s." % ", ".join(gone))
    return [(rule, filter_frame(frame, rule)) for rule in rules]


def frame_to_xlsx_bytes(frame: pd.DataFrame, sheet_name: str) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name=(safe_name(sheet_name)[:31] or "DATA"), index=False)
    return buf.getvalue()


def split_file_name(source_filename: str, rule: Dict[str, str]) -> str:
    stem = re.sub(r"\.[A-Za-z0-9]+$", "", str(source_filename))
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", stem).strip("_") or "input"
    return "%s__%s.xlsx" % (stem, safe_name(rule["name"]))
