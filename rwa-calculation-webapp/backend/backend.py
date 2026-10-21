# -*- coding: utf-8 -*-

"""
RWA Calculation Webapp
Dataiku Python backend

Main capabilities
-----------------
1. Category configuration
2. List managed-folder files
3. Upload multiple source data files
4. Upload or replace one mapping file
5. Download files from managed folders
6. Delete files
7. Validate source and mapping files
8. Validate inputs and run the configured category recipe
9. Compare multiple columns on the same row
10. Trigger category-specific Dataiku Python recipes
11. Maintain run history and audit logs

Required Dataiku managed folders
--------------------------------
For every category:
    1 data input folder
    1 mapping folder
    1 output folder
    1 template folder (secondary recipe output)

Configure the actual managed folder IDs in CATEGORY_CONFIG.
"""

import os
import re
import io
import json
import uuid
import math
import traceback
import logging
from copy import copy
from datetime import datetime, timezone

import numpy as np
import pandas as pd

import dataiku
from flask import request, jsonify, Response, stream_with_context

from excel_deep_trace import ExcelDeepTraceEngine
from webapp_core import split as split_core


# =============================================================================
# GENERAL CONFIGURATION
# =============================================================================

APP_TITLE = "RWA Calculation"

LOG = logging.getLogger("rwa_calculation_webapp")
LOG.setLevel(logging.INFO)

RECIPE_POLLING_MODE = "synchronous"

ALLOWED_DATA_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".csv"
}

ALLOWED_MAPPING_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".csv"
}

ALLOWED_OUTPUT_EXTENSIONS = {
    ".xlsx",
    ".csv",
    ".json"
}
ALLOWED_TEMPLATE_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".csv",
    ".json"
}

MAX_UPLOAD_SIZE_MB = 250
MAX_UPLOAD_SIZE_BYTES = MAX_UPLOAD_SIZE_MB * 1024 * 1024

METADATA_DIRECTORY = "_webapp_metadata"
RUN_HISTORY_FILE = METADATA_DIRECTORY + "/run_history.json"
AUDIT_LOG_FILE = METADATA_DIRECTORY + "/audit_log.json"
SELECTION_MANIFEST_FILE = METADATA_DIRECTORY + "/selected_inputs.json"

DEFAULT_SHEET_NAME = 0

NULL_TEXT_VALUES = {
    "",
    "NONE",
    "NULL",
    "N/A",
    "NA",
    "NAN",
    "<NA>"
}


# =============================================================================
# CATEGORY CONFIGURATION
# =============================================================================

"""
Replace the managed folder IDs below with your actual Dataiku managed folder IDs.

Do not expose these folder IDs in JavaScript.
The browser only sends the category key to the backend.
"""

CATEGORY_CONFIG = {
    "government": {
        "display_name": "Government",
        "description": "Government exposure RWA calculation",
        "icon": "landmark",
        "data_folder_id": "GOVERNMENT_DATA_FOLDER_ID",
        "mapping_folder_id": "GOVERNMENT_MAPPING_FOLDER_ID",
        "output_folder_id": "GOVERNMENT_OUTPUT_FOLDER_ID",

        "template_folder_id": "REPLACE_GOVERNMENT_TEMPLATE_FOLDER_ID",
        "recipe_id": "REPLACE_GOVERNMENT_RECIPE_ID",
        "calculation_handler": "calculate_government",
        "required_input_columns": [],
        "required_mapping_columns": [],
        "comparison_rules": [
            {
                "left_column": "SOURCE_RWA",
                "right_column": "CALCULATED_RWA",
                "comparison_type": "numeric",
                "absolute_tolerance": 0.01,
                "relative_tolerance": 0.000001
            },
            {
                "left_column": "SOURCE_RISK_WEIGHT",
                "right_column": "CALCULATED_RISK_WEIGHT",
                "comparison_type": "numeric",
                "absolute_tolerance": 0.000001,
                "relative_tolerance": 0.000001
            }
        ]
    },

    "corporate": {
        "display_name": "Corporate",
        "description": "Corporate exposure RWA calculation",
        "icon": "building",
        "data_folder_id": "mReNlOzQ",
        "mapping_folder_id": "gCxSFlpK",
        "output_folder_id": "5I5ZSrWc",

        "template_folder_id": "iK1IJYr9",
        "recipe_id": "compute_QvDsUspv",
        "calculation_handler": "calculate_corporate",
        "required_input_columns": [],
        "required_mapping_columns": [],
        "comparison_rules": [
            {
                "left_column": "SOURCE_RWA",
                "right_column": "CALCULATED_RWA",
                "comparison_type": "numeric",
                "absolute_tolerance": 0.01,
                "relative_tolerance": 0.000001
            }
        ]
    },

    "institutions": {
        "display_name": "Institutions",
        "description": "Institutions exposure RWA calculation",
        "icon": "bank",
        "data_folder_id": "WB28JAfi",
        "mapping_folder_id": "shveMrqs",
        "output_folder_id": "cXkZJwZy",

        "template_folder_id": "DAVd0l9d",
        "recipe_id": "compute_RDrI9CkK",
        "calculation_handler": "calculate_institutions",
        "required_input_columns": [],
        "required_mapping_columns": [],
        "comparison_rules": [
            {
                "left_column": "SOURCE_RWA",
                "right_column": "CALCULATED_RWA",
                "comparison_type": "numeric",
                "absolute_tolerance": 0.01,
                "relative_tolerance": 0.000001
            }
        ]
    },

    "cgcb": {
        "display_name": "CGCB",
        "description": "CGCB exposure RWA calculation",
        "icon": "briefcase",
        "data_folder_id": "THQmkw7R",
        "mapping_folder_id": "QeVMzeaL",
        "output_folder_id": "GKZFG5b1",

        "template_folder_id": "ZpDfHeqb",
        "recipe_id": "compute_lCchXNoe",
        "calculation_handler": "calculate_cgcb",
        "required_input_columns": [],
        "required_mapping_columns": [],
        "comparison_rules": [
            {
                "left_column": "SOURCE_RWA",
                "right_column": "CALCULATED_RWA",
                "comparison_type": "numeric",
                "absolute_tolerance": 0.01,
                "relative_tolerance": 0.000001
            }
        ]
    },
    "retail": {
        "display_name": "Retail",
        "description": "RWA calculation for Retail",
        "icon": "layers",
        "data_folder_id": "UxEG1uQL",
        "mapping_folder_id": "OvwVxyAX",
        "output_folder_id": "cTUpQ03o",

        "template_folder_id": "mE37wT3m",
        "recipe_id": "compute_cTUpQ03o",
        "calculation_handler": "calculate_retail",
        "required_input_columns": [],
        "required_mapping_columns": [],
        "comparison_rules": [
            {
                "left_column": "SOURCE_RWA",
                "right_column": "CALCULATED_RWA",
                "comparison_type": "numeric",
                "absolute_tolerance": 0.01,
                "relative_tolerance": 0.000001
            }
        ]
    },
    
    "category_6": {
        "display_name": "Category 6",
        "description": "RWA calculation for Category 6",
        "icon": "layers",
        "data_folder_id": "CATEGORY_6_DATA_FOLDER_ID",
        "mapping_folder_id": "CATEGORY_6_MAPPING_FOLDER_ID",
        "output_folder_id": "CATEGORY_6_OUTPUT_FOLDER_ID",

        "template_folder_id": "REPLACE_CATEGORY_6_TEMPLATE_FOLDER_ID",
        "recipe_id": "REPLACE_CATEGORY_6_RECIPE_ID",
        "calculation_handler": "calculate_generic",
        "required_input_columns": [],
        "required_mapping_columns": [],
        "comparison_rules": []
    },

    "category_7": {
        "display_name": "Category 7",
        "description": "RWA calculation for Category 7",
        "icon": "layers",
        "data_folder_id": "CATEGORY_7_DATA_FOLDER_ID",
        "mapping_folder_id": "CATEGORY_7_MAPPING_FOLDER_ID",
        "output_folder_id": "CATEGORY_7_OUTPUT_FOLDER_ID",

        "template_folder_id": "REPLACE_CATEGORY_7_TEMPLATE_FOLDER_ID",
        "recipe_id": "REPLACE_CATEGORY_7_RECIPE_ID",
        "calculation_handler": "calculate_generic",
        "required_input_columns": [],
        "required_mapping_columns": [],
        "comparison_rules": []
    },

    "category_8": {
        "display_name": "Category 8",
        "description": "RWA calculation for Category 8",
        "icon": "layers",
        "data_folder_id": "CATEGORY_8_DATA_FOLDER_ID",
        "mapping_folder_id": "CATEGORY_8_MAPPING_FOLDER_ID",
        "output_folder_id": "CATEGORY_8_OUTPUT_FOLDER_ID",

        "template_folder_id": "REPLACE_CATEGORY_8_TEMPLATE_FOLDER_ID",
        "recipe_id": "REPLACE_CATEGORY_8_RECIPE_ID",
        "calculation_handler": "calculate_generic",
        "required_input_columns": [],
        "required_mapping_columns": [],
        "comparison_rules": []
    },

    "category_9": {
        "display_name": "Category 9",
        "description": "RWA calculation for Category 9",
        "icon": "layers",
        "data_folder_id": "CATEGORY_9_DATA_FOLDER_ID",
        "mapping_folder_id": "CATEGORY_9_MAPPING_FOLDER_ID",
        "output_folder_id": "CATEGORY_9_OUTPUT_FOLDER_ID",

        "template_folder_id": "REPLACE_CATEGORY_9_TEMPLATE_FOLDER_ID",
        "recipe_id": "REPLACE_CATEGORY_9_RECIPE_ID",
        "calculation_handler": "calculate_generic",
        "required_input_columns": [],
        "required_mapping_columns": [],
        "comparison_rules": []
    }
}


# =============================================================================
# STANDARD API RESPONSE HELPERS
# =============================================================================

def success_response(data=None, message=None, status_code=200):
    payload = {
        "success": True,
        "message": message,
        "data": data
    }
    return jsonify(payload), status_code


def error_response(message, details=None, status_code=400):
    payload = {
        "success": False,
        "message": message,
        "details": details
    }
    return jsonify(payload), status_code


def utc_now_iso():
    return datetime.now(timezone.utc).isoformat()


def generate_run_id():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = uuid.uuid4().hex[:8].upper()
    return "RUN_{}_{}".format(timestamp, suffix)


# =============================================================================
# SECURITY AND VALIDATION HELPERS
# =============================================================================

def normalize_category_key(value):
    if value is None:
        raise ValueError("Category is required.")

    key = str(value).strip().lower()

    if key not in CATEGORY_CONFIG:
        raise ValueError(
            "Unsupported category '{}'. Supported categories are {}."
            .format(value, sorted(CATEGORY_CONFIG.keys()))
        )

    return key


def sanitize_filename(filename):
    """
    Removes path traversal and unsupported characters.
    """

    if filename is None:
        raise ValueError("Filename is required.")

    filename = os.path.basename(str(filename).strip())

    if not filename:
        raise ValueError("Filename cannot be blank.")

    filename = re.sub(r"[\x00-\x1f\x7f]", "", filename)
    filename = re.sub(r'[<>:"/\\|?*]', "_", filename)
    filename = re.sub(r"\s+", " ", filename).strip()

    if filename in {".", ".."}:
        raise ValueError("Invalid filename.")

    return filename


def normalize_folder_path(path):
    if path is None:
        raise ValueError("File path is required.")

    path = str(path).replace("\\", "/").strip()
    path = path.lstrip("/")

    if not path:
        raise ValueError("File path cannot be blank.")

    parts = [part for part in path.split("/") if part]

    if any(part in {".", ".."} for part in parts):
        raise ValueError("Invalid file path.")

    return "/" + "/".join(parts)


def get_extension(filename):
    return os.path.splitext(str(filename))[1].lower()


def validate_extension(filename, allowed_extensions):
    extension = get_extension(filename)

    if extension not in allowed_extensions:
        raise ValueError(
            "Unsupported file type '{}'. Allowed types are {}."
            .format(extension or "<none>", sorted(allowed_extensions))
        )


def get_category_config(category):
    category_key = normalize_category_key(category)
    return category_key, CATEGORY_CONFIG[category_key]


def get_folder(category, folder_type):
    category_key, config = get_category_config(category)

    folder_field_map = {
        "data": "data_folder_id",
        "mapping": "mapping_folder_id",
        "output": "output_folder_id",
        "template": "template_folder_id"
    }

    if folder_type not in folder_field_map:
        raise ValueError(
            "Unsupported folder type '{}'."
            .format(folder_type)
        )

    folder_id = config[folder_field_map[folder_type]]

    placeholder_folder = (
        not folder_id
        or str(folder_id).upper().endswith("_FOLDER_ID")
    )
    if placeholder_folder:
        raise ValueError(
            "{} folder is not configured for category '{}'. Replace the placeholder managed folder ID in CATEGORY_CONFIG."
            .format(folder_type.title(), category_key)
        )

    return dataiku.Folder(folder_id)


def get_allowed_extensions(folder_type):
    if folder_type == "data":
        return ALLOWED_DATA_EXTENSIONS

    if folder_type == "mapping":
        return ALLOWED_MAPPING_EXTENSIONS

    if folder_type == "output":
        return ALLOWED_OUTPUT_EXTENSIONS
    if folder_type == "template":
        return ALLOWED_TEMPLATE_EXTENSIONS

    raise ValueError("Unsupported folder type '{}'.".format(folder_type))


# =============================================================================
# MANAGED FOLDER HELPERS
# =============================================================================

def safe_folder_details(folder):
    try:
        return folder.get_info()
    except Exception:
        return {}


def list_folder_files(folder, allowed_extensions=None, include_metadata=False):
    paths = folder.list_paths_in_partition()

    files = []

    for raw_path in paths:
        normalized_path = normalize_folder_path(raw_path)
        filename = os.path.basename(normalized_path)

        if filename.startswith("."):
            continue

        if normalized_path.startswith("/" + METADATA_DIRECTORY + "/"):
            continue

        extension = get_extension(filename)

        if allowed_extensions and extension not in allowed_extensions:
            continue

        file_record = {
            "path": normalized_path,
            "filename": filename,
            "extension": extension,
            "size": None,
            "last_modified": None
        }

        if include_metadata:
            try:
                details = folder.get_path_details(normalized_path)

                if isinstance(details, dict):
                    file_record["size"] = details.get("size")
                    file_record["last_modified"] = (
                        details.get("lastModified")
                        or details.get("last_modified")
                    )
            except Exception:
                pass

        files.append(file_record)

    files.sort(
        key=lambda item: (
            item.get("last_modified") or "",
            item["filename"].lower()
        ),
        reverse=True
    )

    return files


def folder_file_exists(folder, path):
    normalized_path = normalize_folder_path(path)

    try:
        paths = folder.list_paths_in_partition()
        normalized_paths = {
            normalize_folder_path(item)
            for item in paths
        }
        return normalized_path in normalized_paths

    except Exception:
        return False


def read_json_from_folder(folder, path, default_value):
    normalized_path = normalize_folder_path(path)

    if not folder_file_exists(folder, normalized_path):
        return copy(default_value)

    try:
        with folder.get_download_stream(normalized_path) as stream:
            content = stream.read()

        if not content:
            return copy(default_value)

        if isinstance(content, bytes):
            content = content.decode("utf-8")

        return json.loads(content)

    except Exception:
        return copy(default_value)


def write_json_to_folder(folder, path, payload):
    normalized_path = normalize_folder_path(path)

    encoded = json.dumps(
        payload,
        ensure_ascii=False,
        indent=2,
        default=str
    ).encode("utf-8")

    folder.upload_stream(normalized_path, io.BytesIO(encoded))


def append_audit_log(category, action, details=None, status="SUCCESS"):
    try:
        category_key, config = get_category_config(category)
        output_folder = get_folder(category_key, "output")

        audit_log = read_json_from_folder(
            output_folder,
            AUDIT_LOG_FILE,
            []
        )

        audit_log.append({
            "timestamp": utc_now_iso(),
            "category": category_key,
            "action": action,
            "status": status,
            "details": details or {}
        })

        audit_log = audit_log[-2000:]

        write_json_to_folder(
            output_folder,
            AUDIT_LOG_FILE,
            audit_log
        )

    except Exception:
        pass


def append_run_history(category, run_record):
    category_key, config = get_category_config(category)
    output_folder = get_folder(category_key, "output")

    history = read_json_from_folder(
        output_folder,
        RUN_HISTORY_FILE,
        []
    )

    history.append(run_record)
    history = history[-500:]

    write_json_to_folder(
        output_folder,
        RUN_HISTORY_FILE,
        history
    )


# =============================================================================
# FILE READING
# =============================================================================

def read_managed_folder_file(folder, path):
    normalized_path = normalize_folder_path(path)

    with folder.get_download_stream(normalized_path) as stream:
        content = stream.read()

    if not content:
        raise ValueError(
            "File '{}' is empty."
            .format(normalized_path)
        )

    if not isinstance(content, bytes):
        content = bytes(content)

    return content


def standardize_column_name(value):
    text = str(value).strip()
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^A-Za-z0-9_]+", "_", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("_").upper()


def standardize_dataframe_columns(frame):
    result = frame.copy()

    new_columns = []
    observed = {}

    for raw_column in result.columns:
        base_column = standardize_column_name(raw_column)

        if not base_column:
            base_column = "UNNAMED_COLUMN"

        count = observed.get(base_column, 0)
        observed[base_column] = count + 1

        if count == 0:
            unique_column = base_column
        else:
            unique_column = "{}_{}".format(base_column, count + 1)

        new_columns.append(unique_column)

    result.columns = new_columns
    return result


def _normalize_header_candidate(value):
    if pd.isna(value):
        return ""
    return standardize_column_name(value)


def detect_header_row(raw_frame, expected_columns=None, max_scan_rows=100):
    """
    Detect the real header row instead of assuming row 1.
    Priority is exact/partial coverage of configured expected columns. Otherwise,
    choose the strongest unique, non-empty row that has populated data beneath it.
    """
    if raw_frame is None or raw_frame.empty:
        raise ValueError("The file does not contain readable rows.")
    expected = {
        standardize_column_name(column)
        for column in (expected_columns or [])
        if str(column).strip()
    }
    scan_limit = min(len(raw_frame), max_scan_rows)
    best = None
    for row_index in range(scan_limit):
        values = list(raw_frame.iloc[row_index])
        normalized = [_normalize_header_candidate(value) for value in values]
        non_empty = [value for value in normalized if value]
        if not non_empty:
            continue
        unique_count = len(set(non_empty))
        duplicate_count = len(non_empty) - unique_count
        expected_hits = len(expected.intersection(non_empty))
        # A header normally has unique labels and at least two populated cells.
        # Data beneath the candidate must populate comparable positions.
        lookahead = raw_frame.iloc[row_index + 1:min(row_index + 6, len(raw_frame))]
        populated_below = 0
        if not lookahead.empty:
            candidate_positions = [i for i, value in enumerate(normalized) if value]
            populated_below = int(sum(
                1 for i in candidate_positions
                if i < lookahead.shape[1] and lookahead.iloc[:, i].notna().any()
            ))
        text_like = sum(
            1 for value in values
            if isinstance(value, str) and str(value).strip()
        )
        score = (
            expected_hits * 10000
            + min(unique_count, 50) * 100
            + min(populated_below, 50) * 20
            + min(text_like, 50) * 5
            - duplicate_count * 200
        )
        # Reject likely title/note rows when no configured expected columns exist.
        if not expected and len(non_empty) < 2:
            score -= 5000
        candidate = (score, expected_hits, unique_count, -row_index, row_index)
        if best is None or candidate > best:
            best = candidate
    if best is None:
        raise ValueError("Unable to detect a valid header row in the first {} rows.".format(scan_limit))
    header_row = best[-1]
    if expected and best[1] == 0:
        raise ValueError(
            "Header detection could not find any configured expected columns in the first {} rows."
            .format(scan_limit)
        )
    return header_row


def _frame_from_detected_header(raw_frame, header_row):
    header_values = list(raw_frame.iloc[header_row])
    data = raw_frame.iloc[header_row + 1:].copy()
    data.columns = header_values
    data = data.dropna(how="all").dropna(axis=1, how="all")
    data = standardize_dataframe_columns(data)
    data.attrs["detected_header_row"] = int(header_row + 1)
    return data


def read_tabular_content(
    content,
    filename,
    sheet_name=DEFAULT_SHEET_NAME,
    expected_columns=None,
    max_header_scan_rows=100
):
    extension = get_extension(filename)
    buffer = io.BytesIO(content)
    if extension in {".xlsx", ".xlsm"}:
        raw = pd.read_excel(buffer, sheet_name=sheet_name, header=None, engine="openpyxl")
    elif extension == ".xlsb":
        raw = pd.read_excel(buffer, sheet_name=sheet_name, header=None, engine="pyxlsb")
    elif extension == ".csv":
        try:
            raw = pd.read_csv(buffer, header=None, encoding="utf-8-sig", low_memory=False)
        except UnicodeDecodeError:
            buffer.seek(0)
            raw = pd.read_csv(buffer, header=None, encoding="latin-1", low_memory=False)
    else:
        raise ValueError("Unsupported tabular file type '{}'.".format(extension))
    raw = raw.dropna(how="all").reset_index(drop=True)
    header_row = detect_header_row(raw, expected_columns, max_header_scan_rows)
    return _frame_from_detected_header(raw, header_row)


def read_input_files(category, selected_paths):
    category_key, config = get_category_config(category)
    data_folder = dataiku.Folder(config["data_folder_id"])

    frames = []
    file_summaries = []

    for selected_path in selected_paths:
        normalized_path = normalize_folder_path(selected_path)
        filename = os.path.basename(normalized_path)

        validate_extension(
            filename,
            ALLOWED_DATA_EXTENSIONS
        )

        content = read_managed_folder_file(
            data_folder,
            normalized_path
        )

        frame = read_tabular_content(
            content,
            filename,
            expected_columns=config.get("required_input_columns", [])
        )

        detected_header_row = int(frame.attrs.get("detected_header_row", 1))
        frame["SOURCE_FILE"] = filename
        frame["SOURCE_ROW_NUMBER"] = np.arange(
            detected_header_row + 1,
            detected_header_row + 1 + len(frame)
        )

        frames.append(frame)

        file_summaries.append({
            "path": normalized_path,
            "filename": filename,
            "row_count": int(len(frame)),
            "column_count": int(len(frame.columns)),
            "detected_header_row": detected_header_row
        })

    if not frames:
        raise ValueError(
            "No input data files were selected."
        )

    combined = pd.concat(
        frames,
        ignore_index=True,
        sort=False
    )

    return combined, file_summaries


def get_active_mapping_file(category):
    category_key, config = get_category_config(category)
    mapping_folder = dataiku.Folder(config["mapping_folder_id"])

    mapping_files = list_folder_files(
        mapping_folder,
        allowed_extensions=ALLOWED_MAPPING_EXTENSIONS,
        include_metadata=True
    )

    if not mapping_files:
        raise ValueError(
            "No mapping file is available for category '{}'."
            .format(config["display_name"])
        )

    if len(mapping_files) > 1:
        raise ValueError(
            "More than one mapping file exists for category '{}'. "
            "Keep only one active mapping file."
            .format(config["display_name"])
        )

    return mapping_files[0]


def is_mapping_sheet_name(sheet_name):
    return bool(
        re.match(
            r"^mapping(?:\s+\d+)?$",
            str(sheet_name).strip(),
            flags=re.IGNORECASE
        )
    )

def is_toggle_sheet_name(sheet_name):
    return bool(
        re.match(
            r"^toggle(?:\s+\d+)?$",
            str(sheet_name).strip(),
            flags=re.IGNORECASE
        )
    )

def detect_mapping_sheet(content):

    excel = pd.ExcelFile(
        io.BytesIO(content),
        engine="openpyxl"
    )

    mapping_sheets = [
        sheet
        for sheet in excel.sheet_names
        if is_mapping_sheet_name(sheet)
    ]

    if not mapping_sheets:
        raise ValueError(
            "No Mapping sheet found. "
            "Expected Mapping, Mapping 1, Mapping 2..."
        )

    return mapping_sheets[0]

def detect_toggle_sheets(content):

    excel = pd.ExcelFile(
        io.BytesIO(content),
        engine="openpyxl"
    )

    return [
        sheet
        for sheet in excel.sheet_names
        if is_toggle_sheet_name(sheet)
    ]

def read_active_mapping(category):

    category_key, config = get_category_config(category)

    mapping_folder = dataiku.Folder(
        config["mapping_folder_id"]
    )

    mapping_file = get_active_mapping_file(
        category_key
    )

    content = read_managed_folder_file(
        mapping_folder,
        mapping_file["path"]
    )

    mapping_sheet = detect_mapping_sheet(
        content
    )

    toggle_sheets = detect_toggle_sheets(
        content
    )

    mapping_frame = read_tabular_content(
        content,
        mapping_file["filename"],
        sheet_name=mapping_sheet,
        expected_columns=config.get(
            "required_mapping_columns",
            []
        )
    )

    mapping_file["mapping_sheet"] = mapping_sheet
    mapping_file["toggle_sheets"] = toggle_sheets

    return (
        mapping_frame,
        mapping_file
    )


# =============================================================================
# INPUT VALIDATION
# =============================================================================

def validate_required_columns(frame, required_columns, source_label):
    required_columns = [
        standardize_column_name(column)
        for column in required_columns
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in frame.columns
    ]

    return {
        "source": source_label,
        "required_columns": required_columns,
        "missing_columns": missing_columns,
        "valid": len(missing_columns) == 0
    }


def validate_mapping_uniqueness(mapping_frame, key_columns=None):
    if not key_columns:
        return {
            "valid": True,
            "duplicate_count": 0,
            "key_columns": []
        }

    normalized_keys = [
        standardize_column_name(column)
        for column in key_columns
    ]

    missing_keys = [
        column
        for column in normalized_keys
        if column not in mapping_frame.columns
    ]

    if missing_keys:
        return {
            "valid": False,
            "duplicate_count": None,
            "key_columns": normalized_keys,
            "missing_key_columns": missing_keys
        }

    duplicate_mask = mapping_frame.duplicated(
        subset=normalized_keys,
        keep=False
    )

    return {
        "valid": not duplicate_mask.any(),
        "duplicate_count": int(duplicate_mask.sum()),
        "key_columns": normalized_keys,
        "missing_key_columns": []
    }


def validate_category_inputs(category, selected_paths):
    category_key, config = get_category_config(category)

    input_frame, file_summaries = read_input_files(
        category_key,
        selected_paths
    )

    mapping_frame, mapping_file = read_active_mapping(
        category_key
    )

    input_column_validation = validate_required_columns(
        input_frame,
        config.get("required_input_columns", []),
        "input"
    )

    mapping_column_validation = validate_required_columns(
        mapping_frame,
        config.get("required_mapping_columns", []),
        "mapping"
    )

    validation_errors = []

    if input_frame.empty:
        validation_errors.append(
            "The combined input data is empty."
        )

    if mapping_frame.empty:
        validation_errors.append(
            "The Mapping sheet contains no mapping rows."
        )

    if not input_column_validation["valid"]:
        validation_errors.append(
            "Missing input column(s): {}."
            .format(
                ", ".join(
                    input_column_validation["missing_columns"]
                )
            )
        )

    if not mapping_column_validation["valid"]:
        validation_errors.append(
            "Missing mapping column(s): {}."
            .format(
                ", ".join(
                    mapping_column_validation["missing_columns"]
                )
            )
        )

    return {
        "valid": len(validation_errors) == 0,
        "category": category_key,
        "input_row_count": int(len(input_frame)),
        "input_column_count": int(len(input_frame.columns)),
        "mapping_row_count": int(len(mapping_frame)),
        "mapping_column_count": int(len(mapping_frame.columns)),
        "input_files": file_summaries,
        "mapping_file": mapping_file,
        "input_column_validation": input_column_validation,
        "mapping_column_validation": mapping_column_validation,
        "errors": validation_errors
    }

import os
import uuid
import shutil
import tempfile
import threading

from flask import request, jsonify, send_file

from excel_deep_trace import ExcelDeepTraceEngine


# ============================================================
# Configuration
# ============================================================

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

MAX_UPLOAD_SIZE_BYTES = 100 * 1024 * 1024

DEFAULT_MAX_DEPTH = 12
DEFAULT_MAX_NODES = 500
DEFAULT_MAX_EXPANDED_RANGE = 10000


# ============================================================
# Runtime session store
# ============================================================

TRACE_SESSIONS = {}
TRACE_SESSION_LOCK = threading.RLock()


# ============================================================
# Helpers
# ============================================================

def json_error(message, status=400, details=None):
    payload = {
        "success": False,
        "error": message
    }

    if details is not None:
        payload["details"] = details

    return jsonify(payload), status


def get_file_extension(filename):
    return os.path.splitext(filename or "")[1].lower()


def require_session(session_id):
    with TRACE_SESSION_LOCK:
        session = TRACE_SESSIONS.get(session_id)

    if session is None:
        raise KeyError(
            "Trace session does not exist or the webapp backend "
            "has been restarted."
        )

    return session


def safe_int(value, default, minimum=None, maximum=None):
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = default

    if minimum is not None:
        number = max(number, minimum)

    if maximum is not None:
        number = min(number, maximum)

    return number


# ============================================================
# Analyze workbook
# ============================================================

@app.route("/excel-trace/analyze", methods=["POST"])
def excel_trace_analyze():

    uploaded = request.files.get("file")

    if uploaded is None:
        return json_error("No Excel workbook was uploaded.")

    if not uploaded.filename:
        return json_error("The uploaded file does not have a filename.")

    extension = get_file_extension(uploaded.filename)

    if extension not in ALLOWED_EXTENSIONS:
        return json_error(
            "Unsupported workbook type.",
            details={
                "supported_extensions": sorted(ALLOWED_EXTENSIONS),
                "received_extension": extension
            }
        )

    uploaded.stream.seek(0, os.SEEK_END)
    file_size = uploaded.stream.tell()
    uploaded.stream.seek(0)

    if file_size > MAX_UPLOAD_SIZE_BYTES:
        return json_error(
            "Workbook exceeds the configured upload limit.",
            status=413,
            details={
                "file_size_bytes": file_size,
                "maximum_size_bytes": MAX_UPLOAD_SIZE_BYTES
            }
        )

    session_id = uuid.uuid4().hex

    session_directory = tempfile.mkdtemp(
        prefix="excel_deep_trace_"
    )

    safe_filename = os.path.basename(uploaded.filename)

    workbook_path = os.path.join(
        session_directory,
        safe_filename
    )

    try:
        uploaded.save(workbook_path)

        engine = ExcelDeepTraceEngine(
            workbook_path,
            max_expanded_range=DEFAULT_MAX_EXPANDED_RANGE
        )

        summary = engine.scan()

        with TRACE_SESSION_LOCK:
            TRACE_SESSIONS[session_id] = {
                "session_id": session_id,
                "directory": session_directory,
                "workbook_path": workbook_path,
                "filename": safe_filename,
                "engine": engine
            }

        return jsonify({
            "success": True,
            "session_id": session_id,
            "filename": safe_filename,
            "summary": summary
        })

    except Exception as exc:
        shutil.rmtree(
            session_directory,
            ignore_errors=True
        )

        return json_error(
            "Workbook analysis failed.",
            status=500,
            details=str(exc)
        )

@app.route("/excel-trace/explain-cell", methods=["POST"])
def excel_trace_explain_cell():

    payload = request.get_json(force=True)

    session_id = payload.get("session_id")
    cell_reference = payload.get("cell")

    if not session_id:
        return json_error("session_id is required.")

    if not cell_reference:
        return json_error(
            "cell is required in Sheet!Cell format."
        )

    max_depth = safe_int(
        payload.get("max_depth"),
        DEFAULT_MAX_DEPTH,
        minimum=1,
        maximum=50
    )

    max_nodes = safe_int(
        payload.get("max_nodes"),
        DEFAULT_MAX_NODES,
        minimum=10,
        maximum=5000
    )

    try:
        session = require_session(session_id)
        engine = session["engine"]

        key = engine._normalize_key(cell_reference)
        record = engine.records.get(key)

        result = {
            "success": True,
            "cell": key,
            "value": engine.value(key),
            "formula": (
                record.formula
                if record is not None
                else None
            ),
            "formula_metadata": {
                "functions": (
                    record.functions
                    if record is not None
                    else []
                ),
                "formula_family": (
                    record.family
                    if record is not None
                    else None
                ),
                "formula_length": (
                    record.formula_length
                    if record is not None
                    else 0
                ),
                "dependency_count": (
                    record.dependency_count
                    if record is not None
                    else 0
                ),
                "is_conditional": (
                    record.is_conditional
                    if record is not None
                    else False
                ),
                "is_lookup": (
                    record.is_lookup
                    if record is not None
                    else False
                ),
                "is_volatile": (
                    record.is_volatile
                    if record is not None
                    else False
                )
            },
            "branch_analysis": (
                engine.branch_analysis(key)
                if record is not None
                else None
            ),
            "lookup_analysis": (
                engine.lookup_analysis(key)
                if record is not None
                else None
            ),
            "precedent_graph": engine.trace(
                key,
                direction="precedents",
                max_depth=max_depth,
                max_nodes=max_nodes
            ),
            "dependent_graph": engine.trace(
                key,
                direction="dependents",
                max_depth=max_depth,
                max_nodes=max_nodes
            )
        }

        return jsonify(result)

    except KeyError as exc:
        return json_error(str(exc), status=404)

    except Exception as exc:
        return json_error(
            "Cell explanation failed.",
            status=500,
            details=str(exc)
        )
    
    

@app.route("/excel-trace/row-journey", methods=["POST"])
def excel_trace_row_journey():

    payload = request.get_json(force=True)

    session_id = payload.get("session_id")
    sheet_name = payload.get("sheet")
    row_number = payload.get("row")

    if not session_id:
        return json_error("session_id is required.")

    if not sheet_name:
        return json_error("sheet is required.")

    row_number = safe_int(
        row_number,
        default=0,
        minimum=1
    )

    if row_number < 1:
        return json_error(
            "row must be a positive integer."
        )

    try:
        session = require_session(session_id)
        engine = session["engine"]

        if sheet_name not in engine.wb_formula.sheetnames:
            return json_error(
                "Worksheet does not exist.",
                status=404,
                details={
                    "requested_sheet": sheet_name,
                    "available_sheets":
                        engine.wb_formula.sheetnames
                }
            )

        journey = engine.row_journey(
            sheet_name,
            row_number
        )

        return jsonify({
            "success": True,
            "journey": journey
        })

    except Exception as exc:
        return json_error(
            "Row journey analysis failed.",
            status=500,
            details=str(exc)
        )


@app.route("/excel-trace/compare-cells", methods=["POST"])
def excel_trace_compare_cells():

    payload = request.get_json(force=True)

    session_id = payload.get("session_id")
    left_cell = payload.get("left")
    right_cell = payload.get("right")

    if not all([
        session_id,
        left_cell,
        right_cell
    ]):
        return json_error(
            "session_id, left and right are required."
        )

    try:
        session = require_session(session_id)
        engine = session["engine"]

        comparison = engine.compare_cells(
            left_cell,
            right_cell
        )

        comparison["left_branch"] = (
            engine.branch_analysis(left_cell)
        )

        comparison["right_branch"] = (
            engine.branch_analysis(right_cell)
        )

        comparison["left_lookup"] = (
            engine.lookup_analysis(left_cell)
        )

        comparison["right_lookup"] = (
            engine.lookup_analysis(right_cell)
        )

        return jsonify({
            "success": True,
            "comparison": comparison
        })

    except Exception as exc:
        return json_error(
            "Cell comparison failed.",
            status=500,
            details=str(exc)
        )
    
@app.route("/excel-trace/findings", methods=["GET"])
def excel_trace_findings():

    session_id = request.args.get("session_id")
    severity = request.args.get("severity")
    finding_code = request.args.get("code")

    if not session_id:
        return json_error("session_id is required.")

    try:
        session = require_session(session_id)
        engine = session["engine"]

        findings = []

        for finding in engine.findings:

            if severity:
                if finding.severity.upper() != severity.upper():
                    continue

            if finding_code:
                if finding.code.upper() != finding_code.upper():
                    continue

            findings.append({
                "severity": finding.severity,
                "code": finding.code,
                "cell": finding.cell,
                "message": finding.message,
                "evidence": finding.evidence
            })

        return jsonify({
            "success": True,
            "total": len(findings),
            "findings": findings
        })

    except Exception as exc:
        return json_error(
            "Unable to retrieve findings.",
            status=500,
            details=str(exc)
        )


@app.route("/excel-trace/formulas", methods=["GET"])
def excel_trace_formulas():

    session_id = request.args.get("session_id")
    sheet_filter = request.args.get("sheet")
    function_filter = request.args.get("function")

    if not session_id:
        return json_error("session_id is required.")

    try:
        session = require_session(session_id)
        engine = session["engine"]

        formulas = []

        for record in engine.records.values():

            if sheet_filter:
                if record.sheet != sheet_filter:
                    continue

            if function_filter:
                requested_function = function_filter.upper()

                if requested_function not in record.functions:
                    continue

            formulas.append({
                "cell": record.cell,
                "sheet": record.sheet,
                "coordinate": record.coordinate,
                "formula": record.formula,
                "value": record.cached_value,
                "functions": record.functions,
                "dependency_count":
                    record.dependency_count,
                "is_conditional":
                    record.is_conditional,
                "is_lookup":
                    record.is_lookup,
                "is_volatile":
                    record.is_volatile
            })

        return jsonify({
            "success": True,
            "total": len(formulas),
            "formulas": formulas
        })

    except Exception as exc:
        return json_error(
            "Unable to retrieve formulas.",
            status=500,
            details=str(exc)
        )


    
    


# =============================================================================
# RWA CALCULATION FUNCTIONS
# =============================================================================

def apply_mapping(
    input_frame,
    mapping_frame,
    input_key_columns=None,
    mapping_key_columns=None,
    mapping_columns=None
):
    """
    Generic left-join helper.

    Set the key columns inside the category-specific calculation functions.
    """

    if not input_key_columns or not mapping_key_columns:
        return input_frame.copy()

    input_keys = [
        standardize_column_name(column)
        for column in input_key_columns
    ]

    mapping_keys = [
        standardize_column_name(column)
        for column in mapping_key_columns
    ]

    missing_input_keys = [
        column
        for column in input_keys
        if column not in input_frame.columns
    ]

    missing_mapping_keys = [
        column
        for column in mapping_keys
        if column not in mapping_frame.columns
    ]

    if missing_input_keys:
        raise ValueError(
            "Input mapping key column(s) missing: {}."
            .format(", ".join(missing_input_keys))
        )

    if missing_mapping_keys:
        raise ValueError(
            "Mapping key column(s) missing: {}."
            .format(", ".join(missing_mapping_keys))
        )

    selected_mapping = mapping_frame.copy()

    if mapping_columns:
        normalized_mapping_columns = [
            standardize_column_name(column)
            for column in mapping_columns
        ]

        required_mapping_selection = list(
            dict.fromkeys(
                mapping_keys + normalized_mapping_columns
            )
        )

        selected_mapping = selected_mapping[
            required_mapping_selection
        ].copy()

    selected_mapping = selected_mapping.drop_duplicates(
        subset=mapping_keys,
        keep="last"
    )

    rename_mapping_keys = {
        mapping_key: input_key
        for input_key, mapping_key in zip(
            input_keys,
            mapping_keys
        )
        if input_key != mapping_key
    }

    selected_mapping = selected_mapping.rename(
        columns=rename_mapping_keys
    )

    result = input_frame.merge(
        selected_mapping,
        how="left",
        on=input_keys,
        suffixes=("", "_MAPPING"),
        validate="m:1"
    )

    return result


def calculate_government(input_frame, mapping_frame, context):
    """
    Insert the Government calculation logic here.

    This fallback implementation preserves the input data and mapping context.
    """

    result = input_frame.copy()

    if (
        "EXPOSURE_AMOUNT" in result.columns
        and "CALCULATED_RISK_WEIGHT" in result.columns
    ):
        exposure = pd.to_numeric(
            result["EXPOSURE_AMOUNT"],
            errors="coerce"
        )

        risk_weight = pd.to_numeric(
            result["CALCULATED_RISK_WEIGHT"],
            errors="coerce"
        )

        result["CALCULATED_RWA"] = exposure * risk_weight

    return result


def calculate_corporate(input_frame, mapping_frame, context):
    """
    Replace or extend with the Corporate calculation formula.
    """

    result = input_frame.copy()

    if (
        "EXPOSURE_AMOUNT" in result.columns
        and "CALCULATED_RISK_WEIGHT" in result.columns
    ):
        exposure = pd.to_numeric(
            result["EXPOSURE_AMOUNT"],
            errors="coerce"
        )

        risk_weight = pd.to_numeric(
            result["CALCULATED_RISK_WEIGHT"],
            errors="coerce"
        )

        result["CALCULATED_RWA"] = exposure * risk_weight

    return result


def calculate_institutions(input_frame, mapping_frame, context):
    """
    Replace or extend with the Institutions calculation formula.
    """

    result = input_frame.copy()

    if (
        "EXPOSURE_AMOUNT" in result.columns
        and "CALCULATED_RISK_WEIGHT" in result.columns
    ):
        exposure = pd.to_numeric(
            result["EXPOSURE_AMOUNT"],
            errors="coerce"
        )

        risk_weight = pd.to_numeric(
            result["CALCULATED_RISK_WEIGHT"],
            errors="coerce"
        )

        result["CALCULATED_RWA"] = exposure * risk_weight

    return result


def calculate_cgcb(input_frame, mapping_frame, context):
    """
    Replace or extend with the CGCB calculation formula.
    """

    result = input_frame.copy()

    if (
        "EXPOSURE_AMOUNT" in result.columns
        and "CALCULATED_RISK_WEIGHT" in result.columns
    ):
        exposure = pd.to_numeric(
            result["EXPOSURE_AMOUNT"],
            errors="coerce"
        )

        risk_weight = pd.to_numeric(
            result["CALCULATED_RISK_WEIGHT"],
            errors="coerce"
        )

        result["CALCULATED_RWA"] = exposure * risk_weight

    return result


def calculate_generic(input_frame, mapping_frame, context):
    """
    Generic fallback for categories whose formulas have not yet been inserted.
    """

    result = input_frame.copy()

    if (
        "EXPOSURE_AMOUNT" in result.columns
        and "CALCULATED_RISK_WEIGHT" in result.columns
    ):
        exposure = pd.to_numeric(
            result["EXPOSURE_AMOUNT"],
            errors="coerce"
        )

        risk_weight = pd.to_numeric(
            result["CALCULATED_RISK_WEIGHT"],
            errors="coerce"
        )

        result["CALCULATED_RWA"] = exposure * risk_weight

    return result


CALCULATION_HANDLERS = {
    "calculate_government": calculate_government,
    "calculate_corporate": calculate_corporate,
    "calculate_institutions": calculate_institutions,
    "calculate_cgcb": calculate_cgcb,
    "calculate_generic": calculate_generic
}


def execute_category_calculation(
    category,
    input_frame,
    mapping_frame,
    context
):
    category_key, config = get_category_config(category)

    handler_name = config.get(
        "calculation_handler",
        "calculate_generic"
    )

    handler = CALCULATION_HANDLERS.get(handler_name)

    if handler is None:
        raise ValueError(
            "Calculation handler '{}' is not registered."
            .format(handler_name)
        )

    result = handler(
        input_frame.copy(),
        mapping_frame.copy(),
        context
    )

    if not isinstance(result, pd.DataFrame):
        raise TypeError(
            "Calculation handler '{}' did not return a pandas DataFrame."
            .format(handler_name)
        )

    return standardize_dataframe_columns(result)


# =============================================================================
# SAME-ROW COLUMN COMPARISON ENGINE
# =============================================================================

def normalize_text_value(value):
    if pd.isna(value):
        return None

    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)

    if text.upper() in NULL_TEXT_VALUES:
        return None

    return text.upper()


def normalize_date_series(series):
    return pd.to_datetime(
        series,
        errors="coerce"
    ).dt.normalize()


def numeric_comparison(
    left_series,
    right_series,
    absolute_tolerance,
    relative_tolerance
):
    left_numeric = pd.to_numeric(
        left_series,
        errors="coerce"
    )

    right_numeric = pd.to_numeric(
        right_series,
        errors="coerce"
    )

    both_null = (
        left_numeric.isna()
        & right_numeric.isna()
    )

    one_null = (
        left_numeric.isna()
        ^ right_numeric.isna()
    )

    difference = left_numeric - right_numeric
    absolute_difference = difference.abs()

    denominator = pd.concat(
        [
            left_numeric.abs(),
            right_numeric.abs()
        ],
        axis=1
    ).max(axis=1)

    denominator = denominator.replace(0, np.nan)

    relative_difference = (
        absolute_difference / denominator
    )

    match_mask = (
        both_null
        | (
            ~one_null
            & (
                absolute_difference <= absolute_tolerance
            )
        )
        | (
            ~one_null
            & (
                relative_difference <= relative_tolerance
            )
        )
    )

    match_mask = match_mask.fillna(False)

    return {
        "match": match_mask,
        "difference": difference,
        "absolute_difference": absolute_difference,
        "relative_difference": relative_difference
    }


def text_comparison(left_series, right_series):
    left_normalized = left_series.map(
        normalize_text_value
    )

    right_normalized = right_series.map(
        normalize_text_value
    )

    both_null = (
        left_normalized.isna()
        & right_normalized.isna()
    )

    match_mask = (
        both_null
        | (
            left_normalized == right_normalized
        )
    )

    return {
        "match": match_mask.fillna(False)
    }


def exact_comparison(left_series, right_series):
    both_null = (
        left_series.isna()
        & right_series.isna()
    )

    match_mask = (
        both_null
        | (
            left_series == right_series
        )
    )

    return {
        "match": match_mask.fillna(False)
    }


def date_comparison(left_series, right_series):
    left_dates = normalize_date_series(left_series)
    right_dates = normalize_date_series(right_series)

    both_null = (
        left_dates.isna()
        & right_dates.isna()
    )

    match_mask = (
        both_null
        | (
            left_dates == right_dates
        )
    )

    return {
        "match": match_mask.fillna(False)
    }


def apply_comparison_rules(frame, comparison_rules):
    result = frame.copy()

    result["MISMATCH_COLUMNS"] = ""
    result["MISMATCH_REASONS"] = ""
    result["MISMATCH_COUNT"] = 0

    rule_summaries = []

    for rule_index, rule in enumerate(comparison_rules, start=1):
        left_column = standardize_column_name(
            rule.get("left_column")
        )

        right_column = standardize_column_name(
            rule.get("right_column")
        )

        comparison_type = str(
            rule.get("comparison_type", "exact")
        ).strip().lower()

        status_column = rule.get(
            "status_column",
            "{}_VS_{}_STATUS".format(
                left_column,
                right_column
            )
        )

        status_column = standardize_column_name(
            status_column
        )

        if (
            left_column not in result.columns
            or right_column not in result.columns
        ):
            missing_columns = [
                column
                for column in [
                    left_column,
                    right_column
                ]
                if column not in result.columns
            ]

            rule_summaries.append({
                "rule_number": rule_index,
                "left_column": left_column,
                "right_column": right_column,
                "comparison_type": comparison_type,
                "status": "SKIPPED",
                "reason": (
                    "Missing column(s): {}"
                    .format(", ".join(missing_columns))
                )
            })

            continue

        if comparison_type == "numeric":
            absolute_tolerance = float(
                rule.get(
                    "absolute_tolerance",
                    0.0
                )
            )

            relative_tolerance = float(
                rule.get(
                    "relative_tolerance",
                    0.0
                )
            )

            comparison = numeric_comparison(
                result[left_column],
                result[right_column],
                absolute_tolerance,
                relative_tolerance
            )

            difference_column = standardize_column_name(
                rule.get(
                    "difference_column",
                    "{}_VS_{}_DIFFERENCE".format(
                        left_column,
                        right_column
                    )
                )
            )

            percentage_column = standardize_column_name(
                rule.get(
                    "percentage_difference_column",
                    "{}_VS_{}_DIFFERENCE_PERCENT".format(
                        left_column,
                        right_column
                    )
                )
            )

            result[difference_column] = comparison[
                "difference"
            ]

            result[percentage_column] = (
                comparison["relative_difference"] * 100
            )

        elif comparison_type == "text":
            comparison = text_comparison(
                result[left_column],
                result[right_column]
            )

        elif comparison_type == "date":
            comparison = date_comparison(
                result[left_column],
                result[right_column]
            )

        elif comparison_type == "exact":
            comparison = exact_comparison(
                result[left_column],
                result[right_column]
            )

        else:
            raise ValueError(
                "Unsupported comparison type '{}'."
                .format(comparison_type)
            )

        match_mask = comparison["match"]
        mismatch_mask = ~match_mask

        result[status_column] = np.where(
            match_mask,
            "MATCH",
            "MISMATCH"
        )

        mismatch_label = "{} <> {}".format(
            left_column,
            right_column
        )

        result.loc[
            mismatch_mask,
            "MISMATCH_COLUMNS"
        ] = result.loc[
            mismatch_mask,
            "MISMATCH_COLUMNS"
        ].apply(
            lambda existing: (
                mismatch_label
                if not existing
                else existing + "; " + mismatch_label
            )
        )

        reason_label = "{} comparison failed".format(
            comparison_type.title()
        )

        result.loc[
            mismatch_mask,
            "MISMATCH_REASONS"
        ] = result.loc[
            mismatch_mask,
            "MISMATCH_REASONS"
        ].apply(
            lambda existing: (
                reason_label
                if not existing
                else existing + "; " + reason_label
            )
        )

        result.loc[
            mismatch_mask,
            "MISMATCH_COUNT"
        ] += 1

        rule_summaries.append({
            "rule_number": rule_index,
            "left_column": left_column,
            "right_column": right_column,
            "comparison_type": comparison_type,
            "status": "APPLIED",
            "match_count": int(match_mask.sum()),
            "mismatch_count": int(mismatch_mask.sum())
        })

    result["OVERALL_VALIDATION_STATUS"] = np.where(
        result["MISMATCH_COUNT"] > 0,
        "MISMATCH",
        "MATCH"
    )

    summary = {
        "total_rows": int(len(result)),
        "matched_rows": int(
            (
                result["OVERALL_VALIDATION_STATUS"]
                == "MATCH"
            ).sum()
        ),
        "mismatched_rows": int(
            (
                result["OVERALL_VALIDATION_STATUS"]
                == "MISMATCH"
            ).sum()
        ),
        "rules": rule_summaries
    }

    if summary["total_rows"] > 0:
        summary["match_rate"] = round(
            (
                summary["matched_rows"]
                / summary["total_rows"]
            ) * 100,
            4
        )
    else:
        summary["match_rate"] = None

    return result, summary


# =============================================================================
# OUTPUT WRITING
# =============================================================================

def make_excel_safe(value):
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None

    return value


def autosize_worksheet(worksheet, dataframe):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="EAF2FF"
    )

    header_font = Font(
        bold=True,
        color="172033"
    )

    for header_cell in worksheet[1]:
        header_cell.fill = copy(header_fill)
        header_cell.font = copy(header_font)
        header_cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_index, column_name in enumerate(
        dataframe.columns,
        start=1
    ):
        values = dataframe[column_name].astype(str)
        content_width = values.map(len).max() if len(values) else 0
        width = max(len(str(column_name)), content_width)
        width = min(max(width + 2, 10), 50)

        worksheet.column_dimensions[
            get_column_letter(column_index)
        ].width = width


def dataframe_to_excel_bytes(result_frame, mismatch_frame=None, run_summary=None, input_file_summary=None):
    """One clean output sheet containing only calculated data."""
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        result_frame.to_excel(writer, sheet_name=0, index=False)
        autosize_worksheet(writer.book["RWA Output"], result_frame)
    output_buffer.seek(0)
    return output_buffer.getvalue()


def is_primary_rwa_output(file_record):
    name = str(file_record.get("filename", ""))
    upper = name.upper()
    return get_extension(name) == ".xlsx" and "_RWA_" in upper and "_COMPARISON_" not in upper


def read_calculated_output(
    output_folder,
    file_record,
    expected_columns=None
):
    content = read_managed_folder_file(
        output_folder,
        file_record["path"]
    )

    return read_tabular_content(
        content,
        file_record["filename"],
        sheet_name=0,
        expected_columns=expected_columns,
        max_header_scan_rows=100
    )


def get_category_output_schema(category):
    category_key, config = get_category_config(category)
    folder = get_folder(category_key, "output")
    files = [f for f in list_folder_files(folder, {".xlsx"}, True) if is_primary_rwa_output(f)]
    if not files:
        return {"category": category_key, "columns": [], "file_count": 0, "detected_header_row": None}
    frame = read_calculated_output(folder, files[0])
    return {
        "category": category_key,
        "columns": list(frame.columns),
        "file_count": len(files),
        "detected_header_row": frame.attrs.get("detected_header_row")
    }


def validate_user_comparison_rules(raw_rules, columns, threshold=0.0):
    available=set(columns); rules=[]; threshold=float(threshold or 0)
    for i, raw in enumerate(raw_rules or [], 1):
        left=standardize_column_name(raw.get("left_column", "")); right=standardize_column_name(raw.get("right_column", ""))
        kind=str(raw.get("comparison_type", "exact")).lower().strip()
        if not left or not right: raise ValueError("Comparison {} requires two columns.".format(i))
        if left == right: raise ValueError("Comparison {} must use different columns.".format(i))
        missing=[x for x in (left,right) if x not in available]
        if missing: raise ValueError("Comparison {} column(s) not found: {}.".format(i, ", ".join(missing)))
        if kind not in {"exact","text","numeric","date"}: raise ValueError("Unsupported comparison type '{}'.".format(kind))
        rule={"left_column":left,"right_column":right,"comparison_type":kind,"status_column":"COMPARE_{}_STATUS".format(i)}
        if kind=="numeric": rule.update({"absolute_tolerance":threshold,"relative_tolerance":0.0})
        rules.append(rule)
    if not rules: raise ValueError("Select at least one column pair.")
    return rules


def compare_all_category_outputs(category, raw_rules, threshold=0.0):
    category_key, config=get_category_config(category)
    folder=get_folder(category_key, "output")
    files=[f for f in list_folder_files(folder,{".xlsx"},True) if is_primary_rwa_output(f)]
    if not files: raise ValueError("No calculated output files are available for this category.")
    first=read_calculated_output(folder,files[0]); expected=list(first.columns)
    rules=validate_user_comparison_rules(raw_rules,expected,threshold); frames=[]; header_rows=[]
    for file_record in files:
        frame=read_calculated_output(folder,file_record,expected_columns=expected)
        actual = list(frame.columns)
        if actual != expected:
            missing = [column for column in expected if column not in actual]
            extra = [column for column in actual if column not in expected]
            raise ValueError(
                "Output schema mismatch in '{}'. Missing: {}. Extra: {}."
                .format(file_record["filename"], missing, extra)
            )
        header_rows.append({"file":file_record["filename"],"header_row":frame.attrs.get("detected_header_row")})
        compared,_=apply_comparison_rules(frame,rules)
        compared.insert(0,"SOURCE_OUTPUT_FILE",file_record["filename"]); frames.append(compared)
    result=pd.concat(frames,ignore_index=True,sort=False)
    total=len(result); mismatches=int((result["OVERALL_VALIDATION_STATUS"]=="MISMATCH").sum()); matches=total-mismatches
    safe=re.sub(r"[^A-Za-z0-9]+","_",config["display_name"]).strip("_")
    filename="{}_COMPARISON_{}.xlsx".format(safe,datetime.now().strftime("%Y%m%d_%H%M%S"))
    buffer=io.BytesIO()
    with pd.ExcelWriter(buffer,engine="openpyxl") as writer:
        result.to_excel(writer,sheet_name="Comparison",index=False)
        autosize_worksheet(writer.book["Comparison"],result)
    buffer.seek(0); folder.upload_stream(normalize_folder_path(filename),buffer)
    return {"category":category_key,"files_compared":len(files),"total_rows":int(total),"matched_rows":int(matches),"mismatched_rows":mismatches,"match_rate":round(matches/total*100,4) if total else None,"comparison_file":filename,"detected_headers":header_rows}


# =============================================================================
# RECIPE SELECTION MANIFEST
# =============================================================================

def write_selection_manifest(category, selected_paths, run_id, split_result=None):
    """Persist the exact UI selection so the Python Recipe processes only those files.

    When the data was split, ``selected_paths`` lists the split datasets (the files the
    recipe should process), ``source_paths`` the original uploads, and ``datasets`` tags
    every split file with its rule so the recipe can process each type accordingly.
    """
    category_key, config = get_category_config(category)
    data_folder = get_folder(category_key, "data")
    existing_paths = {
        item["path"]
        for item in list_folder_files(data_folder, ALLOWED_DATA_EXTENSIONS)
    }
    normalized = []
    for path in selected_paths:
        normalized_path = normalize_folder_path(path)
        if normalized_path not in existing_paths:
            raise ValueError(
                "Selected input file '{}' no longer exists."
                .format(normalized_path)
            )
        normalized.append(normalized_path)
    manifest = {
        "run_id": run_id,
        "category": category_key,
        "selected_paths": normalized,
        "created_at": utc_now_iso()
    }
    if split_result:
        manifest["source_paths"] = normalized
        manifest["selected_paths"] = [d["path"] for d in split_result["datasets"]]
        manifest["split"] = {
            "enabled": True,
            "rules": split_result["rules"],
            "datasets": split_result["datasets"]
        }
    write_json_to_folder(data_folder, SELECTION_MANIFEST_FILE, manifest)
    return manifest


# =============================================================================
# DATA SPLIT (optional step between upload and validation)
# =============================================================================

SPLIT_DIRECTORY = METADATA_DIRECTORY + "/split"


def get_split_defaults(config):
    return [dict(rule) for rule in (config.get("split_rules") or split_core.DEFAULT_SPLIT_RULES)]


def parse_split_request(category, split_payload):
    """Return validated rules when the UI asked for a split, else None."""
    if not split_payload or not split_payload.get("enabled"):
        return None
    category_key, config = get_category_config(category)
    return split_core.clean_rules(
        split_payload.get("rules") or get_split_defaults(config),
        standardize_column_name
    )


def read_split_source(category, path):
    category_key, config = get_category_config(category)
    folder = get_folder(category_key, "data")
    normalized_path = normalize_folder_path(path)
    filename = os.path.basename(normalized_path)
    validate_extension(filename, ALLOWED_DATA_EXTENSIONS)
    frame = read_tabular_content(
        read_managed_folder_file(folder, normalized_path),
        filename,
        expected_columns=config.get("required_input_columns", [])
    )
    return normalized_path, filename, frame


def preview_split(category, selected_paths, rules):
    """Columns available for splitting plus the row count each rule would produce."""
    category_key, config = get_category_config(category)
    folder = get_folder(category_key, "data")
    paths = list(selected_paths or [])
    if not paths:
        paths = [
            item["path"]
            for item in list_folder_files(folder, ALLOWED_DATA_EXTENSIONS)
        ]

    columns = []
    files = []
    missing = set()

    for path in paths:
        try:
            normalized_path, filename, frame = read_split_source(category_key, path)
        except Exception as exc:
            files.append({"path": path, "filename": os.path.basename(str(path)), "error": str(exc)})
            continue

        for column in frame.columns:
            if column not in columns:
                columns.append(column)

        record = {"path": normalized_path, "filename": filename, "rows": int(len(frame))}
        if rules:
            gone = split_core.missing_columns(frame, rules)
            missing.update(gone)
            record["datasets"] = [
                {
                    "name": rule["name"],
                    "rows": (
                        None if rule["column"] in gone
                        else int(len(split_core.filter_frame(frame, rule)))
                    )
                }
                for rule in rules
            ]
        files.append(record)

    return {
        "columns": columns,
        "files": files,
        "missing_columns": sorted(missing)
    }


def clear_split_files(category):
    folder = get_folder(category, "data")
    prefix = "/" + SPLIT_DIRECTORY + "/"
    for raw_path in folder.list_paths_in_partition():
        if normalize_folder_path(raw_path).startswith(prefix):
            folder.delete_path(raw_path)


def run_split(category, selected_paths, rules, run_id):
    """Write one single-sheet workbook per (input file, rule) and describe them."""
    category_key, config = get_category_config(category)
    folder = get_folder(category_key, "data")
    clear_split_files(category_key)

    datasets = []
    skipped = []

    for path in selected_paths:
        source_path, source_name, frame = read_split_source(category_key, path)

        for rule, subset in split_core.split_frame(frame, rules):
            if subset.empty:
                skipped.append(
                    "{}: no rows where {} = '{}' ({} dataset skipped).".format(
                        source_name, rule["column"], rule["value"], rule["name"]
                    )
                )
                continue

            split_name = split_core.split_file_name(source_name, rule)
            split_path = "/{}/{}".format(SPLIT_DIRECTORY, split_name)
            folder.upload_stream(
                split_path,
                io.BytesIO(split_core.frame_to_xlsx_bytes(subset, rule["name"]))
            )
            datasets.append({
                "path": split_path,
                "filename": split_name,
                "source_path": source_path,
                "source_filename": source_name,
                "split_name": rule["name"],
                "column": rule["column"],
                "value": rule["value"],
                "row_count": int(len(subset))
            })

    if not datasets:
        raise ValueError(
            "The split produced no rows for any dataset. " + " ".join(skipped)
        )

    return {"rules": rules, "datasets": datasets, "skipped": skipped, "run_id": run_id}

# =============================================================================
# DATAIKU RECIPE ORCHESTRATION
# =============================================================================

def is_placeholder(value):
    text = str(value or "").strip().upper()
    return not text or text.startswith("REPLACE_") or text.endswith("_FOLDER_ID")


def get_configured_recipe(category):
    category_key, config = get_category_config(category)
    recipe_id = config.get("recipe_id")
    if is_placeholder(recipe_id):
        raise ValueError(
            "Recipe ID is not configured for category '{}'."
            .format(config["display_name"])
        )
    client = dataiku.api_client()
    project = client.get_default_project()
    return category_key, config, project.get_recipe(recipe_id)


def serialize_recipe_result(result):
    if result is None:
        return {"state": "COMPLETED"}
    if isinstance(result, dict):
        safe = {}
        for key, value in result.items():
            try:
                json.dumps(value)
                safe[key] = value
            except Exception:
                safe[key] = str(value)
        safe.setdefault("state", "COMPLETED")
        return safe
    safe = {"state": "COMPLETED"}
    for attribute in ("job_id", "id", "state"):
        try:
            value = getattr(result, attribute, None)
            if callable(value):
                value = value()
            if value is not None:
                safe[attribute] = str(value)
        except Exception:
            continue
    return safe


def snapshot_recipe_artifacts(category):
    """Snapshot both recipe output folders without mixing their schemas."""
    category_key, config = get_category_config(category)
    snapshots = {}
    for folder_type in ("output", "template"):
        folder = get_folder(category_key, folder_type)
        files = list_folder_files(
            folder,
            allowed_extensions=get_allowed_extensions(folder_type),
            include_metadata=True
        )
        snapshots[folder_type] = {item["path"]: item for item in files}
    return snapshots

def changed_artifacts(before, after):
    produced = []
    for folder_type in ("output", "template"):
        old_files = before.get(folder_type, {})
        new_files = after.get(folder_type, {})
        for path, record in new_files.items():
            old = old_files.get(path)
            if old is None or (
                record.get("size") != old.get("size")
                or record.get("last_modified") != old.get("last_modified")
            ):
                enriched = dict(record)
                enriched["folder_type"] = folder_type
                produced.append(enriched)
    return produced


def run_category_recipe(category, selected_paths, run_id, split_rules=None):
    category_key, config, recipe = get_configured_recipe(category)

    # Validate only the files selected in the UI before starting the category recipe.
    validation = validate_category_inputs(category_key, selected_paths)
    if not validation["valid"]:
        raise ValueError(
            "Input validation failed: {}"
            .format(" ".join(validation["errors"]))
        )

    split_result = None
    if split_rules:
        split_result = run_split(
            category_key,
            selected_paths,
            split_rules,
            run_id
        )

    selection_manifest = write_selection_manifest(
        category_key,
        selected_paths,
        run_id,
        split_result
    )
    before = snapshot_recipe_artifacts(category_key)
    started_at = utc_now_iso()
    LOG.info(
        "Starting recipe %s for category %s.",
        config["recipe_id"],
        category_key
    )
    raw_result = recipe.run()
    execution = serialize_recipe_result(raw_result)
    completed_at = utc_now_iso()
    after = snapshot_recipe_artifacts(category_key)
    produced_files = changed_artifacts(before, after)

    return {
        "category": category_key,
        "category_display_name": config["display_name"],
        "recipe_id": config["recipe_id"],
        "recipe_name": config.get("recipe_name", config["recipe_id"]),
        "started_at": started_at,
        "completed_at": completed_at,
        "execution": execution,
        "validation": validation,
        "selection_manifest": selection_manifest,
        "split": split_result,
        "produced_files": produced_files,
        "all_output_files": list(after.get("output", {}).values()),
        "all_template_files": list(after.get("template", {}).values())
    }

# =============================================================================
# API ROUTES
# =============================================================================

@app.route("/api/categories", methods=["GET"])
def api_categories():
    try:
        categories = []

        for category_key, config in CATEGORY_CONFIG.items():
            category_record = {
                "key": category_key,
                "display_name": config["display_name"],
                "description": config.get(
                    "description",
                    ""
                ),
                "icon": config.get(
                    "icon",
                    "layers"
                ),
                "data_file_count": 0,
                "mapping_file_count": 0,
                "output_file_count": 0,
                "template_file_count": 0,
                "mapping_status": "MISSING",
                "recipe_configured": not is_placeholder(config.get("recipe_id")),
                "split_enabled": bool(config.get("split_enabled", False)),
                "split_rules": get_split_defaults(config),
                "status": "NOT_READY"
            }

            try:
                data_folder = dataiku.Folder(
                    config["data_folder_id"]
                )

                data_files = list_folder_files(
                    data_folder,
                    ALLOWED_DATA_EXTENSIONS
                )

                category_record[
                    "data_file_count"
                ] = len(data_files)

            except Exception:
                category_record[
                    "data_folder_error"
                ] = True

            try:
                mapping_folder = dataiku.Folder(
                    config["mapping_folder_id"]
                )

                mapping_files = list_folder_files(
                    mapping_folder,
                    ALLOWED_MAPPING_EXTENSIONS
                )

                mapping_count = len(mapping_files)

                category_record[
                    "mapping_file_count"
                ] = mapping_count

                if mapping_count == 1:
                    category_record[
                        "mapping_status"
                    ] = "AVAILABLE"

                elif mapping_count > 1:
                    category_record[
                        "mapping_status"
                    ] = "MULTIPLE_FILES"

            except Exception:
                category_record[
                    "mapping_folder_error"
                ] = True

            try:
                output_folder = dataiku.Folder(
                    config["output_folder_id"]
                )

                output_files = list_folder_files(
                    output_folder,
                    ALLOWED_OUTPUT_EXTENSIONS
                )

                category_record[
                    "output_file_count"
                ] = len(output_files)

            except Exception:
                category_record[
                    "output_folder_error"
                ] = True


            try:
                template_folder = get_folder(category_key, "template")
                template_files = list_folder_files(
                    template_folder,
                    ALLOWED_TEMPLATE_EXTENSIONS
                )
                category_record[
                    "template_file_count"
                ] = len(template_files)
            except Exception:
                category_record[
                    "template_folder_error"
                ] = True

            if (
                category_record["data_file_count"] > 0
                and category_record["mapping_file_count"] == 1
                and category_record["recipe_configured"]
            ):
                category_record["status"] = "READY"

            elif not category_record["recipe_configured"]:
                category_record["status"] = "RECIPE_REQUIRED"

            elif category_record["mapping_file_count"] != 1:
                category_record["status"] = "MAPPING_REQUIRED"

            else:
                category_record["status"] = "DATA_REQUIRED"

            categories.append(category_record)

        return success_response(
            data={
                "app_title": APP_TITLE,
                "categories": categories
            }
        )

    except Exception as exc:
        return error_response(
            "Unable to load categories.",
            details=str(exc),
            status_code=500
        )


@app.route("/api/files", methods=["GET"])
def api_list_files():
    try:
        category = request.args.get("category")
        folder_type = request.args.get("folder_type")

        category_key, config = get_category_config(category)
        folder = get_folder(category_key, folder_type)

        files = list_folder_files(
            folder,
            get_allowed_extensions(folder_type),
            include_metadata=True
        )

        return success_response(
            data={
                "category": category_key,
                "folder_type": folder_type,
                "files": files,
                "count": len(files)
            }
        )

    except Exception as exc:
        return error_response(
            "Unable to list folder files.",
            details=str(exc),
            status_code=400
        )


@app.route("/api/upload", methods=["POST"])
def api_upload_files():
    category = request.form.get("category")
    folder_type = request.form.get("folder_type")
    replace_mapping = (
        request.form.get("replace_mapping", "false")
        .strip()
        .lower()
        == "true"
    )

    try:
        category_key, config = get_category_config(category)

        if folder_type not in {"data", "mapping"}:
            raise ValueError(
                "Uploads are only allowed for data or mapping folders."
            )

        uploaded_files = request.files.getlist("files")

        if not uploaded_files:
            raise ValueError(
                "No files were supplied."
            )

        if folder_type == "mapping" and len(uploaded_files) != 1:
            raise ValueError(
                "Exactly one mapping file must be uploaded."
            )

        destination_folder = get_folder(
            category_key,
            folder_type
        )

        allowed_extensions = get_allowed_extensions(
            folder_type
        )

        if folder_type == "mapping":
            existing_mapping_files = list_folder_files(
                destination_folder,
                ALLOWED_MAPPING_EXTENSIONS
            )

            if existing_mapping_files and not replace_mapping:
                return error_response(
                    "A mapping file already exists. "
                    "Set replace_mapping to true to replace it.",
                    details={
                        "existing_files": existing_mapping_files
                    },
                    status_code=409
                )

            if replace_mapping:
                for existing_file in existing_mapping_files:
                    destination_folder.delete_path(
                        existing_file["path"]
                    )

        uploaded_records = []

        for uploaded_file in uploaded_files:
            filename = sanitize_filename(
                uploaded_file.filename
            )

            validate_extension(
                filename,
                allowed_extensions
            )

            content = uploaded_file.read()

            if not content:
                raise ValueError(
                    "Uploaded file '{}' is empty."
                    .format(filename)
                )

            if len(content) > MAX_UPLOAD_SIZE_BYTES:
                raise ValueError(
                    "File '{}' exceeds the maximum upload size of {} MB."
                    .format(filename, MAX_UPLOAD_SIZE_MB)
                )

            destination_path = normalize_folder_path(
                filename
            )

            destination_folder.upload_stream(
                destination_path,
                io.BytesIO(content)
            )

            uploaded_records.append({
                "filename": filename,
                "path": destination_path,
                "size": len(content)
            })

        append_audit_log(
            category_key,
            "UPLOAD_{}".format(folder_type.upper()),
            details={
                "files": uploaded_records,
                "replace_mapping": replace_mapping
            }
        )

        return success_response(
            data={
                "uploaded_files": uploaded_records
            },
            message="Files uploaded successfully.",
            status_code=201
        )

    except Exception as exc:
        append_audit_log(
            category or "unknown",
            "UPLOAD_FAILED",
            details={
                "folder_type": folder_type,
                "error": str(exc)
            },
            status="FAILED"
        )

        return error_response(
            "File upload failed.",
            details=str(exc),
            status_code=400
        )


@app.route("/api/delete", methods=["POST"])
def api_delete_file():
    try:
        payload = request.get_json(force=True)

        category = payload.get("category")
        folder_type = payload.get("folder_type")
        path = payload.get("path")

        category_key, config = get_category_config(category)

        if folder_type not in {"data", "mapping", "output", "template"}:
            raise ValueError(
                "Unsupported folder type."
            )

        normalized_path = normalize_folder_path(path)
        folder = get_folder(category_key, folder_type)

        if not folder_file_exists(folder, normalized_path):
            raise ValueError(
                "The requested file does not exist."
            )

        folder.delete_path(normalized_path)

        append_audit_log(
            category_key,
            "DELETE_FILE",
            details={
                "folder_type": folder_type,
                "path": normalized_path
            }
        )

        return success_response(
            message="File deleted successfully."
        )

    except Exception as exc:
        return error_response(
            "Unable to delete file.",
            details=str(exc),
            status_code=400
        )


@app.route("/api/download", methods=["GET"])
def api_download_file():
    try:
        category = request.args.get("category")
        folder_type = request.args.get("folder_type")
        path = request.args.get("path")

        category_key, config = get_category_config(category)
        folder = get_folder(category_key, folder_type)

        normalized_path = normalize_folder_path(path)
        filename = os.path.basename(normalized_path)

        if not folder_file_exists(folder, normalized_path):
            return error_response(
                "The requested file does not exist.",
                status_code=404
            )

        download_stream = folder.get_download_stream(
            normalized_path
        )

        def generate():
            try:
                while True:
                    chunk = download_stream.read(
                        1024 * 1024
                    )

                    if not chunk:
                        break

                    yield chunk

            finally:
                try:
                    download_stream.close()
                except Exception:
                    pass

        append_audit_log(
            category_key,
            "DOWNLOAD_FILE",
            details={
                "folder_type": folder_type,
                "path": normalized_path
            }
        )

        response = Response(
            stream_with_context(generate()),
            content_type="application/octet-stream"
        )

        response.headers[
            "Content-Disposition"
        ] = 'attachment; filename="{}"'.format(
            filename.replace('"', "")
        )

        return response

    except Exception as exc:
        return error_response(
            "Unable to download file.",
            details=str(exc),
            status_code=400
        )


@app.route("/api/split-preview", methods=["POST"])
def api_split_preview():
    try:
        payload = request.get_json(force=True) or {}
        category_key, config = get_category_config(payload.get("category"))

        rules = None
        rules_error = None
        if payload.get("rules"):
            try:
                rules = split_core.clean_rules(payload["rules"], standardize_column_name)
            except ValueError as exc:
                rules_error = str(exc)

        preview = preview_split(
            category_key,
            payload.get("selected_paths") or [],
            rules
        )
        preview["rules_error"] = rules_error

        return success_response(data=preview)

    except Exception as exc:
        return error_response(
            "Unable to preview the split.",
            details=str(exc),
            status_code=400
        )


@app.route("/api/validate-inputs", methods=["POST"])
def api_validate_inputs():
    try:
        payload = request.get_json(force=True)

        category = payload.get("category")
        selected_paths = payload.get(
            "selected_paths",
            []
        )

        category_key = normalize_category_key(category)

        validation = validate_category_inputs(
            category_key,
            selected_paths
        )

        if (payload.get("split") or {}).get("enabled"):
            try:
                split_rules = parse_split_request(category_key, payload.get("split"))
                preview = preview_split(category_key, selected_paths, split_rules)
                validation["split_preview"] = preview
                if preview["missing_columns"]:
                    validation["errors"].append(
                        "Split column(s) not found in the selected files: {}."
                        .format(", ".join(preview["missing_columns"]))
                    )
            except ValueError as split_error:
                validation["errors"].append(str(split_error))
            validation["valid"] = len(validation["errors"]) == 0

        append_audit_log(
            category_key,
            "VALIDATE_INPUTS",
            details={
                "selected_paths": selected_paths,
                "validation": validation
            },
            status=(
                "SUCCESS"
                if validation["valid"]
                else "FAILED"
            )
        )

        return success_response(
            data=validation,
            message=(
                "Validation completed successfully."
                if validation["valid"]
                else "Validation completed with errors."
            )
        )

    except Exception as exc:
        return error_response(
            "Input validation failed.",
            details=str(exc),
            status_code=400
        )


@app.route("/api/run", methods=["POST"])
def api_run_calculation():
    run_id = generate_run_id()
    started_at = utc_now_iso()
    category = None

    try:
        payload = request.get_json(force=True) or {}
        category = payload.get("category")
        selected_paths = payload.get("selected_paths", [])
        category_key, config = get_category_config(category)

        if not selected_paths:
            raise ValueError("Select at least one input data file.")

        split_rules = parse_split_request(category_key, payload.get("split"))

        result = run_category_recipe(
            category_key, selected_paths, run_id, split_rules
        )
        produced_files = result.get("produced_files", [])
        primary_outputs = [
            item for item in produced_files
            if item.get("folder_type", "output") == "output" and is_primary_rwa_output(item)
        ]
        output_file = (
            primary_outputs[0]["filename"]
            if primary_outputs
            else (
                produced_files[0]["filename"]
                if produced_files else None
            )
        )

        output_folder_type = "output" if output_file else None

        run_summary = {
            "run_id": run_id,
            "category": category_key,
            "category_display_name": config["display_name"],
            "recipe_id": config["recipe_id"],
            "started_at": result["started_at"],
            "completed_at": result["completed_at"],
            "status": "COMPLETED",
            "selected_paths": selected_paths,
            "total_rows": None,
            "matched_rows": None,
            "mismatched_rows": None,
            "match_rate": None,
            "output_file": output_file,
            "output_folder_type": output_folder_type,
            "produced_files": produced_files,
            "execution": result["execution"],
            "selection_manifest": result["selection_manifest"],
            "split": result.get("split")
        }

        append_run_history(category_key, run_summary)
        append_audit_log(
            category_key,
            "RUN_RECIPE",
            details=run_summary
        )

        return success_response(
            data=run_summary,
            message="Dataiku recipe completed successfully."
        )

    except Exception as exc:
        LOG.exception("Category recipe execution failed.")
        failure_record = {
            "run_id": run_id,
            "started_at": started_at,
            "completed_at": utc_now_iso(),
            "status": "FAILED",
            "error": str(exc)
        }
        try:
            if category:
                category_key = normalize_category_key(category)
                append_run_history(category_key, failure_record)
                append_audit_log(
                    category_key,
                    "RUN_RECIPE",
                    details=failure_record,
                    status="FAILED"
                )
        except Exception:
            pass

        return error_response(
            "Dataiku recipe execution failed.",
            details={"run_id": run_id, "error": str(exc)},
            status_code=500
        )

@app.route("/api/run-history", methods=["GET"])
def api_run_history():
    try:
        category = request.args.get("category")
        category_key, config = get_category_config(category)

        output_folder = dataiku.Folder(
            config["output_folder_id"]
        )

        history = read_json_from_folder(
            output_folder,
            RUN_HISTORY_FILE,
            []
        )

        history = sorted(
            history,
            key=lambda item: item.get(
                "started_at",
                ""
            ),
            reverse=True
        )

        return success_response(
            data={
                "category": category_key,
                "history": history
            }
        )

    except Exception as exc:
        return error_response(
            "Unable to load run history.",
            details=str(exc),
            status_code=400
        )


@app.route("/api/mismatch-preview", methods=["GET"])
def api_mismatch_preview():
    try:
        category = request.args.get("category")
        path = request.args.get("path")
        limit = int(
            request.args.get(
                "limit",
                100
            )
        )

        limit = max(1, min(limit, 500))

        category_key, config = get_category_config(category)
        output_folder = dataiku.Folder(
            config["output_folder_id"]
        )

        normalized_path = normalize_folder_path(path)

        if get_extension(normalized_path) != ".csv":
            raise ValueError(
                "Mismatch preview requires a CSV file."
            )

        content = read_managed_folder_file(
            output_folder,
            normalized_path
        )

        frame = pd.read_csv(
            io.BytesIO(content),
            encoding="utf-8-sig",
            low_memory=False
        )

        preview_frame = frame.head(limit).copy()
        preview_frame = preview_frame.replace(
            [np.nan, np.inf, -np.inf],
            None
        )

        records = preview_frame.to_dict(
            orient="records"
        )

        return success_response(
            data={
                "columns": list(preview_frame.columns),
                "records": records,
                "preview_count": len(records),
                "total_count": int(len(frame))
            }
        )

    except Exception as exc:
        return error_response(
            "Unable to load mismatch preview.",
            details=str(exc),
            status_code=400
        )


@app.route("/api/output-schema", methods=["GET"])
def api_output_schema():
    try:
        return success_response(data=get_category_output_schema(request.args.get("category")))
    except Exception as exc:
        return error_response("Unable to detect output schema.", details=str(exc), status_code=400)


@app.route("/api/compare-all-outputs", methods=["POST"])
def api_compare_all_outputs():
    try:
        payload=request.get_json(force=True) or {}
        result=compare_all_category_outputs(payload.get("category"),payload.get("comparisons",[]),payload.get("threshold",0))
        append_audit_log(result["category"],"COMPARE_OUTPUTS",details=result)
        return success_response(data=result,message="Consolidated comparison completed.")
    except Exception as exc:
        return error_response("Output comparison failed.",details=str(exc),status_code=400)


@app.route("/api/health", methods=["GET"])
def api_health():
    return success_response(
        data={
            "application": APP_TITLE,
            "status": "UP",
            "timestamp": utc_now_iso(),
            "categories": {
                key: {
                    "recipe_configured": not is_placeholder(config.get("recipe_id")),
                    "data_folder_configured": not is_placeholder(config.get("data_folder_id")),
                    "mapping_folder_configured": not is_placeholder(config.get("mapping_folder_id")),
                    "output_folder_configured": not is_placeholder(config.get("output_folder_id")),
                    "template_folder_configured": not is_placeholder(config.get("template_folder_id"))
                }
                for key, config in CATEGORY_CONFIG.items()
            }
        }
    )


# =============================================================================
# RWA webapp additions - admin + deep-trace routes + comparison on the OUTPUT
# folder with every formula evaluated.
# =============================================================================

from webapp_core import auth as _auth

_ADMIN_FILE = "admin.json"
_OVERRIDES_FILE = "category_overrides.json"
_EDITABLE_KEYS = (
    "display_name", "description", "icon", "recipe_id", "recipe_name",
    "data_folder_id", "mapping_folder_id", "output_folder_id", "template_folder_id",
)


def _rwa_cfg_folder():
    folder_id = None
    try:
        folder_id = dataiku.get_custom_variables().get("RWA_WEBAPP_CONFIG_FOLDER")
    except Exception:  # noqa: BLE001
        pass
    folder_id = folder_id or os.environ.get("RWA_WEBAPP_CONFIG_FOLDER")
    if not folder_id:
        raise RuntimeError("project variable RWA_WEBAPP_CONFIG_FOLDER is not set")
    return dataiku.Folder(folder_id)


def _rwa_read_json(name, default):
    try:
        folder = _rwa_cfg_folder()
        present = {normalize_folder_path(p) for p in folder.list_paths_in_partition()}
        if ("/" + name) not in present:
            return default
        with folder.get_download_stream("/" + name) as stream:
            raw = stream.read()
        if not raw:
            return default
        return json.loads(raw.decode("utf-8") if isinstance(raw, bytes) else raw)
    except Exception:  # noqa: BLE001
        return default


def _rwa_write_json(name, obj):
    folder = _rwa_cfg_folder()
    folder.upload_stream(
        "/" + name,
        io.BytesIO(json.dumps(obj, ensure_ascii=False, indent=2, default=str).encode("utf-8")),
    )


def _rwa_apply_overrides():
    overrides = _rwa_read_json(_OVERRIDES_FILE, {}) or {}
    for key, patch in overrides.items():
        if key in CATEGORY_CONFIG and isinstance(patch, dict):
            CATEGORY_CONFIG[key].update(
                {k: v for k, v in patch.items() if k in _EDITABLE_KEYS})
    return overrides


try:
    _rwa_apply_overrides()
except Exception as _exc:  # noqa: BLE001
    LOG.warning("RWA admin overrides not applied: %s", _exc)


def _rwa_admin_record():
    rec = _rwa_read_json(_ADMIN_FILE, None)
    if not rec:
        salt, digest = _auth.hash_password("changeme")
        rec = {"username": "admin", "salt": salt, "hash": digest,
               "must_change": True, "secret": _auth.new_secret()}
        try:
            _rwa_write_json(_ADMIN_FILE, rec)
        except Exception:  # noqa: BLE001
            pass
    return rec


def _rwa_require_admin():
    rec = _rwa_admin_record()
    token = request.headers.get("X-Admin-Token", "")
    return rec if _auth.verify_token(token, rec.get("secret", "")) else None


# --------------------------------------------------------------------------- admin API
@app.route("/api/whoami", methods=["GET"])
def api_whoami():
    return success_response(data={
        "app_title": APP_TITLE,
        "user": os.environ.get("DKU_CURRENT_USER") or None,
    })


@app.route("/api/login", methods=["POST"])
def api_admin_login():
    body = request.get_json(force=True, silent=True) or {}
    rec = _rwa_admin_record()
    ok = (body.get("username") == rec["username"]
          and _auth.verify_password(body.get("password", ""), rec["salt"], rec["hash"]))
    if not ok:
        return error_response("Invalid credentials.", status_code=401)
    return success_response(data={
        "token": _auth.make_token(rec["username"], rec["secret"]),
        "must_change": rec.get("must_change", False),
    })


@app.route("/api/admin/change-credentials", methods=["POST"])
def api_admin_change_credentials():
    rec = _rwa_require_admin()
    if not rec:
        return error_response("Admin authentication required.", status_code=401)
    body = request.get_json(force=True, silent=True) or {}
    if not _auth.verify_password(body.get("current_password", ""), rec["salt"], rec["hash"]):
        return error_response("Current password is wrong.", status_code=403)
    new_pw = body.get("new_password", "")
    if len(new_pw) < 6:
        return error_response("New password must be at least 6 characters.")
    salt, digest = _auth.hash_password(new_pw)
    _rwa_write_json(_ADMIN_FILE, {
        "username": body.get("new_username") or rec["username"],
        "salt": salt, "hash": digest, "must_change": False, "secret": rec["secret"],
    })
    return success_response(message="Credentials updated - sign in again.")


@app.route("/api/admin/config", methods=["GET"])
def api_admin_config():
    if not _rwa_require_admin():
        return error_response("Admin authentication required.", status_code=401)
    out = []
    for key, cfg in CATEGORY_CONFIG.items():
        def clean(field):
            val = cfg.get(field, "")
            return "" if is_placeholder(val) else val
        out.append({
            "key": key,
            "display_name": cfg.get("display_name", key),
            "description": cfg.get("description", ""),
            "icon": cfg.get("icon", "layers"),
            "recipe_id": clean("recipe_id"),
            "recipe_name": cfg.get("recipe_name", ""),
            "folders": {
                "data": clean("data_folder_id"),
                "mapping": clean("mapping_folder_id"),
                "output": clean("output_folder_id"),
                "template": clean("template_folder_id"),
            },
        })
    return success_response(data={"app_title": APP_TITLE, "categories": out})


@app.route("/api/admin/category/<key>", methods=["POST"])
def api_admin_save_category(key):
    if not _rwa_require_admin():
        return error_response("Admin authentication required.", status_code=401)
    if key not in CATEGORY_CONFIG:
        return error_response("Unknown category.", status_code=404)
    body = request.get_json(force=True, silent=True) or {}
    folders = body.get("folders", {}) or {}
    current = CATEGORY_CONFIG[key]
    patch = {
        "display_name": (body.get("display_name") or current.get("display_name") or key).strip(),
        "description": (body.get("description") or current.get("description", "")).strip(),
        "recipe_id": (body.get("recipe_id") or "").strip(),
        "recipe_name": (body.get("recipe_name") or "").strip(),
        "data_folder_id": (folders.get("data") or "").strip(),
        "mapping_folder_id": (folders.get("mapping") or "").strip(),
        "output_folder_id": (folders.get("output") or "").strip(),
        "template_folder_id": (folders.get("template") or "").strip(),
    }
    overrides = _rwa_read_json(_OVERRIDES_FILE, {}) or {}
    overrides[key] = dict(overrides.get(key, {}), **patch)
    _rwa_write_json(_OVERRIDES_FILE, overrides)
    CATEGORY_CONFIG[key].update(patch)
    append_audit_log(key, "ADMIN_UPDATE_CATEGORY", details=patch)
    return success_response(message="Category configuration saved.", data={"key": key})


# --------------------------------------------------------------- deep-trace extra routes
@app.route("/excel-trace/open-output", methods=["POST"])
def excel_trace_open_output():
    body = request.get_json(force=True, silent=True) or {}
    category = body.get("category")
    path = body.get("path")
    folder_type = body.get("folder_type", "output")

    if not category or not path:
        return json_error("category and path are required.")

    session_directory = None
    try:
        category_key, _ = get_category_config(category)
        folder = get_folder(category_key, folder_type)
        content = read_managed_folder_file(folder, normalize_folder_path(path))

        # write the workbook to a temp file and open it by path - same as
        # /excel-trace/analyze, and works with any ExcelDeepTraceEngine build.
        import tempfile
        safe_name = os.path.basename(path).replace("/", "_").replace("\\", "_") or "workbook.xlsx"
        session_directory = tempfile.mkdtemp(prefix="excel_deep_trace_")
        workbook_path = os.path.join(session_directory, safe_name)
        with open(workbook_path, "wb") as fh:
            fh.write(content if isinstance(content, bytes) else bytes(content))

        try:
            engine = ExcelDeepTraceEngine(workbook_path,
                                         max_expanded_range=DEFAULT_MAX_EXPANDED_RANGE)
        except TypeError:
            engine = ExcelDeepTraceEngine(workbook_path)
        summary = engine.scan()

        session_id = uuid.uuid4().hex
        with TRACE_SESSION_LOCK:
            TRACE_SESSIONS[session_id] = {
                "session_id": session_id, "engine": engine, "directory": session_directory,
                "workbook_path": workbook_path, "filename": safe_name, "category": category_key,
            }
        return jsonify({"success": True, "session_id": session_id,
                        "filename": safe_name, "summary": summary})
    except Exception as exc:  # noqa: BLE001
        if session_directory:
            import shutil
            shutil.rmtree(session_directory, ignore_errors=True)
        return json_error("Could not open the output workbook.", status=500, details=str(exc))


@app.route("/excel-trace/columns", methods=["POST"])
def excel_trace_columns():
    body = request.get_json(force=True, silent=True) or {}
    try:
        session = require_session(body.get("session_id"))
        return jsonify({"success": True, **session["engine"].sheet_columns(body.get("sheet"))})
    except KeyError as exc:
        return json_error(str(exc), status=404)
    except Exception as exc:  # noqa: BLE001
        return json_error("Could not read the column list.", status=500, details=str(exc))


@app.route("/excel-trace/mismatch-rows", methods=["POST"])
def excel_trace_mismatch_rows():
    body = request.get_json(force=True, silent=True) or {}
    try:
        session = require_session(body.get("session_id"))
        result = session["engine"].mismatch_rows(
            body.get("sheet"), body.get("left"), body.get("right"),
            float(body.get("tolerance", 0.01)))
        return jsonify({"success": True, **result})
    except KeyError as exc:
        return json_error(str(exc), status=404)
    except Exception as exc:  # noqa: BLE001
        return json_error("Mismatch scan failed.", status=500, details=str(exc))


@app.route("/excel-trace/expression-tree", methods=["POST"])
def excel_trace_expression_tree():
    body = request.get_json(force=True, silent=True) or {}
    try:
        session = require_session(body.get("session_id"))
        tree = session["engine"].expression_tree(
            body.get("cell"), int(body.get("max_depth", 8)))
        return jsonify({"success": True, "tree": tree})
    except KeyError as exc:
        return json_error(str(exc), status=404)
    except Exception as exc:  # noqa: BLE001
        return json_error("Expression tree failed.", status=500, details=str(exc))


# ----------------------------------------------- comparison: OUTPUT folder, formulas evaluated
# The comparison + trace both run on the OUTPUT folder.  Those workbooks carry live Excel
# formulas, so every computed column is EVALUATED here (via ExcelDeepTraceEngine) - the
# comparison never sees an empty "=A2*B2" column even when the recipe stored no cached
# values.  A mismatched row can then be walked node-by-node in the Trace page.
# These redefinitions take precedence over the originals earlier in the file.

_RWA_COMPARE_FOLDER = "output"


def _rwa_is_source_workbook(file_record):
    name = str(file_record.get("filename", ""))
    if name.startswith("~$") or name.startswith("."):
        return False
    return get_extension(name) in {".xlsx", ".xlsm"} and "_COMPARISON_" not in name.upper()


def _rwa_compare_source(category_key):
    """(folder, [file_records], source_name) - the OUTPUT folder, falling back to the
    TEMPLATE folder if the output folder has no usable workbooks."""
    for kind in (_RWA_COMPARE_FOLDER, "template"):
        try:
            folder = get_folder(category_key, kind)
            files = [f for f in list_folder_files(folder, {".xlsx", ".xlsm"}, True)
                     if _rwa_is_source_workbook(f)]
            if files:
                return folder, files, kind
        except Exception:  # noqa: BLE001 - folder not configured
            continue
    folder = get_folder(category_key, "output")
    return folder, [], "output"


def _rwa_values_frame(content, filename, expected_columns=None):
    """Read an .xlsx/.xlsm as a table with **every formula evaluated** (input cells kept
    as-is), through the shared ExcelDeepTraceEngine, then the backend's column
    standardisation. Falls back to a plain read for csv/xlsb."""
    if get_extension(filename) not in {".xlsx", ".xlsm"}:
        return read_tabular_content(content, filename, sheet_name=0,
                                    expected_columns=expected_columns, max_header_scan_rows=100)
    try:
        engine = ExcelDeepTraceEngine(workbook_bytes=content, filename=filename,
                                      max_expanded_range=DEFAULT_MAX_EXPANDED_RANGE)
    except TypeError:
        import tempfile
        _d = tempfile.mkdtemp(prefix="rwa_cmp_")
        _p = os.path.join(_d, os.path.basename(filename) or "workbook.xlsx")
        with open(_p, "wb") as _fh:
            _fh.write(content if isinstance(content, bytes) else bytes(content))
        engine = ExcelDeepTraceEngine(_p, max_expanded_range=DEFAULT_MAX_EXPANDED_RANGE)
    engine.scan()
    # trace the sheet with the most formula cells, else the first sheet
    best = None
    for name in engine.wb_formula.sheetnames:
        n = sum(1 for k in engine.records if k.startswith(name + "!"))
        if best is None or n > best[1]:
            best = (name, n)
    sheet = best[0] if best else engine.wb_formula.sheetnames[0]
    ev = engine.evaluate_sheet(sheet)
    data = pd.DataFrame(ev["rows"], columns=list(ev["headers"]))
    data = standardize_dataframe_columns(data)
    data.attrs["detected_header_row"] = int(ev["header_row"] + 1)
    data.attrs["traced_sheet"] = ev["sheet"]
    return data


def get_category_output_schema(category):
    category_key, config = get_category_config(category)
    folder, files, source = _rwa_compare_source(category_key)
    if not files:
        return {"category": category_key, "columns": [], "file_count": 0,
                "detected_header_row": None, "source_folder": source}
    content = read_managed_folder_file(folder, files[0]["path"])
    frame = _rwa_values_frame(content, files[0]["filename"])
    return {"category": category_key, "columns": list(frame.columns),
            "file_count": len(files),
            "detected_header_row": frame.attrs.get("detected_header_row"),
            "source_folder": source}


def compare_all_category_outputs(category, raw_rules, threshold=0.0):
    category_key, config = get_category_config(category)
    folder, files, source = _rwa_compare_source(category_key)
    if not files:
        raise ValueError("No produced workbooks for this category (looked in the %s "
                         "folder). Run the calculation first." % source)

    def read_frame(fr, expected=None):
        content = read_managed_folder_file(folder, fr["path"])
        return _rwa_values_frame(content, fr["filename"], expected)

    expected = list(read_frame(files[0]).columns)
    rules = validate_user_comparison_rules(raw_rules, expected, threshold)
    frames, header_rows, warnings = [], [], []
    for fr in files:
        frame = read_frame(fr, expected)
        actual = list(frame.columns)
        if actual != expected:
            raise ValueError("Schema mismatch in '%s'. Missing: %s. Extra: %s." % (
                fr["filename"], [c for c in expected if c not in actual],
                [c for c in actual if c not in expected]))
        for rule in rules:
            for col in (rule["left_column"], rule["right_column"]):
                if col in frame.columns and frame[col].isna().all():
                    warnings.append("'%s' in %s is still empty after evaluating its "
                                    "formulas (unsupported function, or genuinely blank)."
                                    % (col, fr["filename"]))
        hr = int(frame.attrs.get("detected_header_row", 1))
        header_rows.append({"file": fr["filename"], "header_row": hr})
        compared, _ = apply_comparison_rules(frame, rules)
        compared.insert(0, "SOURCE_ROW_NUMBER", range(hr + 1, hr + 1 + len(compared)))
        compared.insert(0, "SOURCE_OUTPUT_FILE", fr["filename"])
        frames.append(compared)

    result = pd.concat(frames, ignore_index=True, sort=False)
    total = len(result)
    mismatches = int((result["OVERALL_VALIDATION_STATUS"] == "MISMATCH").sum())
    matches = total - mismatches
    safe = re.sub(r"[^A-Za-z0-9]+", "_", config["display_name"]).strip("_")
    filename = "%s_COMPARISON_%s.xlsx" % (safe, datetime.now().strftime("%Y%m%d_%H%M%S"))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="Comparison", index=False)
        autosize_worksheet(writer.book["Comparison"], result)
    buf.seek(0)
    get_folder(category_key, "output").upload_stream(normalize_folder_path(filename), buf)
    return {"category": category_key, "source_folder": source,
            "files_compared": len(files), "total_rows": int(total),
            "matched_rows": int(matches), "mismatched_rows": mismatches,
            "match_rate": round(matches / total * 100, 4) if total else None,
            "comparison_file": filename, "detected_headers": header_rows,
            "warnings": sorted(set(warnings))}
