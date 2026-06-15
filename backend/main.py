from fastapi import FastAPI, File, UploadFile, Form, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from collections import Counter
from datetime import datetime
from difflib import SequenceMatcher
import json
import csv
import shutil
import os
import re
import uuid
from reconcile_engine import dynamic_vertical_reconciliation # Import your engine

app = FastAPI()

# Allow your frontend to talk to this backend (CORS)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # In production, change this to your frontend domain
    allow_methods=["*"],
    allow_headers=["*"],
)

def cleanup_files(*file_paths):
    """Background task to delete files after sending them to the user"""
    for path in file_paths:
        if os.path.exists(path):
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)


def normalize_text(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def detect_value_type(value):
    text = str(value).strip()
    if text == "":
        return "blank"

    lowered = text.lower()
    if lowered in {"true", "false", "yes", "no", "y", "n"}:
        return "boolean"

    if re.fullmatch(r"[-+]?\d+", text):
        return "integer"

    if re.fullmatch(r"[-+]?\d*\.\d+", text):
        return "decimal"

    date_formats = [
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y/%m/%d",
        "%d-%b-%Y",
        "%d %b %Y",
    ]
    for fmt in date_formats:
        try:
            datetime.strptime(text, fmt)
            return "date"
        except ValueError:
            continue

    return "text"


def profile_csv_file(file_path):
    with open(file_path, mode="r", encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.reader(handle))

    if not rows:
        raise HTTPException(status_code=400, detail=f"File '{os.path.basename(file_path)}' is empty.")

    headers = rows[0]
    data_rows = rows[1:]
    columns = []

    for index, header in enumerate(headers):
        values = [row[index].strip() for row in data_rows if len(row) > index]
        non_empty_values = [value for value in values if value != ""]
        type_counts = Counter(detect_value_type(value) for value in values)
        filtered_counts = {key: count for key, count in type_counts.items() if key != "blank"}
        dominant_type = max(filtered_counts, key=filtered_counts.get) if filtered_counts else "blank"

        sample_values = []
        for value in non_empty_values:
            if value not in sample_values:
                sample_values.append(value)
            if len(sample_values) == 3:
                break

        columns.append({
            "name": header,
            "normalized_name": normalize_text(header),
            "value_type": dominant_type,
            "type_counts": dict(type_counts),
            "non_empty_count": len(non_empty_values),
            "empty_count": len(values) - len(non_empty_values),
            "unique_count": len(set(non_empty_values)),
            "sample_values": sample_values,
        })

    return {
        "file_name": os.path.basename(file_path),
        "row_count": max(len(rows) - 1, 0),
        "column_count": len(headers),
        "headers": headers,
        "columns": columns,
    }


def score_column_pair(core_column, swi_column):
    header_similarity = SequenceMatcher(
        None,
        core_column["normalized_name"],
        swi_column["normalized_name"]
    ).ratio()
    type_similarity = 1.0 if core_column["value_type"] == swi_column["value_type"] else 0.0
    exact_header_bonus = 1.0 if core_column["normalized_name"] and core_column["normalized_name"] == swi_column["normalized_name"] else 0.0

    unique_bonus = 0.0
    if core_column["non_empty_count"] and swi_column["non_empty_count"]:
        unique_ratio_a = core_column["unique_count"] / core_column["non_empty_count"]
        unique_ratio_b = swi_column["unique_count"] / swi_column["non_empty_count"]
        unique_bonus = 1.0 - min(abs(unique_ratio_a - unique_ratio_b), 1.0)

    score = (header_similarity * 0.6) + (type_similarity * 0.25) + (unique_bonus * 0.1) + (exact_header_bonus * 0.5)
    return round(score, 4)


def build_match_suggestions(core_profile, swi_profile):
    suggestions = []

    for core_column in core_profile["columns"]:
        for swi_column in swi_profile["columns"]:
            if core_column["value_type"] != swi_column["value_type"]:
                continue

            score = score_column_pair(core_column, swi_column)
            if score < 0.4:
                continue

            reasons = []
            if core_column["normalized_name"] == swi_column["normalized_name"]:
                reasons.append("matching column name")
            if core_column["value_type"] == swi_column["value_type"]:
                reasons.append(f"same value type ({core_column['value_type']})")
            if core_column["unique_count"] and swi_column["unique_count"]:
                reasons.append("stable unique values")

            suggestions.append({
                "core_column": core_column["name"],
                "swi_column": swi_column["name"],
                "score": score,
                "reason": ", ".join(reasons) if reasons else "similar shape and content",
            })

    suggestions.sort(key=lambda item: item["score"], reverse=True)
    return suggestions[:5]


def get_column_profile(profile, column_name):
    for column in profile["columns"]:
        if column["name"] == column_name:
            return column
    raise HTTPException(status_code=400, detail=f"Column '{column_name}' was not found in the uploaded file.")


def normalize_match_rules(match_rules, core_profile, swi_profile):
    normalized_rules = []
    text_modes = {
        "contains",
        "starts_with",
        "ends_with",
        "core_contains_swi",
        "swi_contains_core",
        "core_starts_with_swi",
        "swi_starts_with_core",
        "core_ends_with_swi",
        "swi_ends_with_core",
    }
    allowed_value_types = {"text", "integer", "date"}

    for index, rule in enumerate(match_rules, start=1):
        core_column_name = rule.get("core_column", "").strip()
        swi_column_name = rule.get("swi_column", "").strip()
        match_mode = rule.get("match_mode", "exact")
        tolerance_value = rule.get("tolerance", 0)

        if not core_column_name or not swi_column_name:
            raise HTTPException(status_code=400, detail=f"Match rule {index} needs both columns selected.")

        core_column = get_column_profile(core_profile, core_column_name)
        swi_column = get_column_profile(swi_profile, swi_column_name)
        core_value_type = rule.get("core_value_type", rule.get("value_type", core_column["value_type"]))
        swi_value_type = rule.get("swi_value_type", rule.get("value_type", swi_column["value_type"]))

        if core_value_type not in allowed_value_types:
            raise HTTPException(status_code=400, detail=f"Match rule {index} uses an unsupported File A data type override.")

        if swi_value_type not in allowed_value_types:
            raise HTTPException(status_code=400, detail=f"Match rule {index} uses an unsupported File B data type override.")

        if core_value_type != swi_value_type:
            raise HTTPException(
                status_code=400,
                detail=f"Match rule {index} must use the same data type on both sides."
            )

        if match_mode in text_modes and core_value_type != "text":
            raise HTTPException(status_code=400, detail=f"Match rule {index} text rule only works for text columns.")

        if match_mode == "integer_tolerance" and core_value_type != "integer":
            raise HTTPException(status_code=400, detail=f"Match rule {index} tolerance mode only works for integer columns.")

        if match_mode == "date_tolerance" and core_value_type != "date":
            raise HTTPException(status_code=400, detail=f"Match rule {index} date interval mode only works for date columns.")

        if match_mode == "contains":
            match_mode = rule.get("match_direction", "core_contains_swi")

        if match_mode not in {
            "exact",
            "integer_tolerance",
            "date_tolerance",
            "core_contains_swi",
            "swi_contains_core",
            "core_starts_with_swi",
            "swi_starts_with_core",
            "core_ends_with_swi",
            "swi_ends_with_core",
        }:
            raise HTTPException(status_code=400, detail=f"Match rule {index} uses an unsupported match mode.")

        try:
            tolerance = int(tolerance_value)
        except (TypeError, ValueError):
            raise HTTPException(status_code=400, detail=f"Match rule {index} interval must be a number.")

        if tolerance < 0:
            raise HTTPException(status_code=400, detail=f"Match rule {index} interval must be zero or greater.")

        normalized_rules.append({
            "core_column": core_column_name,
            "swi_column": swi_column_name,
            "match_mode": match_mode,
            "tolerance": tolerance,
            "value_type": core_value_type,
            "swi_value_type": swi_value_type,
        })

    if not normalized_rules:
        raise HTTPException(status_code=400, detail="Add at least one match column pair.")

    return normalized_rules


def save_uploaded_file(upload_file, folder_path):
    os.makedirs(folder_path, exist_ok=True)
    file_path = os.path.join(folder_path, os.path.basename(upload_file.filename))
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)
    return file_path


def prepare_upload_pair(core_file, swi_file, base_folder):
    request_id = uuid.uuid4().hex
    request_folder = os.path.join(base_folder, request_id)
    folder_a = os.path.join(request_folder, "folder_a")
    folder_b = os.path.join(request_folder, "folder_b")
    core_path = save_uploaded_file(core_file, folder_a)
    swi_path = save_uploaded_file(swi_file, folder_b)
    return request_id, request_folder, folder_a, folder_b, core_path, swi_path


@app.post("/api/analyze")
async def analyze_uploaded_files(
    background_tasks: BackgroundTasks,
    core_file: UploadFile = File(...),
    swi_file: UploadFile = File(...)
):
    temp_root = "temp/analyze"
    request_folder = None

    try:
        request_id, request_folder, folder_a, folder_b, core_path, swi_path = prepare_upload_pair(core_file, swi_file, temp_root)
        core_profile = profile_csv_file(core_path)
        swi_profile = profile_csv_file(swi_path)
        return JSONResponse({
            "request_id": request_id,
            "files": {
                "core": core_profile,
                "swi": swi_profile,
            },
            "suggestions": build_match_suggestions(core_profile, swi_profile),
        })
    except Exception:
        if request_folder:
            cleanup_files(request_folder)
        raise
    finally:
        if request_folder:
            background_tasks.add_task(cleanup_files, request_folder)

@app.post("/api/reconcile")
async def process_reconciliation(
    background_tasks: BackgroundTasks,
    core_file: UploadFile = File(...),
    swi_file: UploadFile = File(...),
    match_col_core: str = Form(""),
    match_col_swi: str = Form(""),
    match_rules: str = Form(""),
    format_type: str = Form("split_tables")
):
    # 1. Create temporary directories
    temp_root = "temp/reconcile"
    request_folder = None

    try:
        _, request_folder, folder_a, folder_b, _, _ = prepare_upload_pair(core_file, swi_file, temp_root)

        core_profile = profile_csv_file(next(path for path in [os.path.join(folder_a, name) for name in os.listdir(folder_a)] if path.endswith(".csv")))
        swi_profile = profile_csv_file(next(path for path in [os.path.join(folder_b, name) for name in os.listdir(folder_b)] if path.endswith(".csv")))

        parsed_rules = []
        if match_rules.strip():
            parsed_rules = normalize_match_rules(json.loads(match_rules), core_profile, swi_profile)
        elif match_col_core.strip() and match_col_swi.strip():
            parsed_rules = normalize_match_rules([
                {
                    "core_column": match_col_core,
                    "swi_column": match_col_swi,
                    "match_mode": "exact",
                    "tolerance": 0,
                }
            ], core_profile, swi_profile)
        else:
            raise HTTPException(status_code=400, detail="Add at least one match column pair.")

        # 2. Run the Reconciliation Engine
        output_excel = "temp/Reconciliation_Report.xlsx"
        dynamic_vertical_reconciliation(
            folder_a=folder_a,
            folder_b=folder_b,
            match_rules=parsed_rules,
            output_file=output_excel,
            output_format=format_type
        )

        if not os.path.exists(output_excel):
            raise HTTPException(
                status_code=400,
                detail="Reconciliation failed. Please make sure the selected match columns exist in both files."
            )

        # 3. Schedule cleanup of temp files AFTER the file is downloaded
        background_tasks.add_task(cleanup_files, request_folder, output_excel)

        # 4. Return the Excel file to the frontend
        return FileResponse(
            path=output_excel, 
            filename="Reconciliation_Report.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception:
        if request_folder:
            cleanup_files(request_folder)
        raise