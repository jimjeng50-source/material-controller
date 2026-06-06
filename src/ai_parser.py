"""
AI-powered column detection using Claude API.
Sends a sample of the sheet (headers + first rows) to Claude,
asks it to identify which column index maps to which field,
then uses that mapping to extract records.
"""
from __future__ import annotations

import json
import re
from datetime import datetime, date
from pathlib import Path
from typing import Any

import openpyxl

from parser import (
    _to_date, _str, _find_header_end,
    ESP_SHEETS, parse_esp,
    _is_submission_control,
)

# ── fields we want Claude to identify ────────────────────────────────────────
FIELD_DESCRIPTIONS = {
    "item_name":   "材料 / 設備名稱 (material or equipment name)",
    "po_no":       "PO 編號 / 文件編碼 (purchase order number or document code)",
    "tag_no":      "設備 Tag 號 / 設備編號 (equipment tag number)",
    "sub_vendor":  "次廠商 / 供應商名稱 (sub-vendor or supplier name)",
    "ros":         "需求到料日期 / 預定進場日期 (required on site date)",
    "eta":         "預計到料日期 / 預定廠驗日期 (expected delivery or factory inspection date)",
    "ata":         "實際到料日期 (actual delivery date)",
    "qty":         "數量 (quantity)",
}

SYSTEM_PROMPT = """\
You are a data extraction assistant. Your job is to analyse a spreadsheet table
(given as JSON rows) and identify which column index (0-based) contains each of
the requested fields. Some fields may not exist in the table — return null for
those. Only return a JSON object, no other text.

Rules:
- Prefer columns whose header text clearly matches the field description.
- If two columns are candidates, pick the one whose header is a closer match.
- For date fields (ros, eta, ata): the column must contain actual calendar dates,
  NOT document reference numbers or text codes.
- po_no must contain alphanumeric codes (like ED2024-001), NOT dates.
- Return only: {"item_name": <int|null>, "po_no": <int|null>, "tag_no": <int|null>,
  "sub_vendor": <int|null>, "ros": <int|null>, "eta": <int|null>,
  "ata": <int|null>, "qty": <int|null>}
"""


def _sheet_to_text_sample(rows: list[tuple], max_rows: int = 8) -> str:
    """Serialise header rows + first data rows to JSON for the AI prompt."""
    sample = []
    for row in rows[:max_rows]:
        sample.append([_str(c).split("\n")[0] for c in row])
    return json.dumps(sample, ensure_ascii=False)


def _call_claude(api_key: str, table_sample: str, sheet_name: str) -> dict | None:
    """Ask Claude to return a column-index mapping as JSON."""
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)

    user_msg = (
        f"Sheet name: {sheet_name}\n\n"
        f"Table sample (rows as JSON arrays, index 0-based):\n{table_sample}\n\n"
        "Fields to identify:\n"
        + "\n".join(f"- {k}: {v}" for k, v in FIELD_DESCRIPTIONS.items())
        + "\n\nReturn only a JSON object with the column indices."
    )

    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",   # fast + cheap for structured extraction
        max_tokens=256,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}],
    )
    raw = msg.content[0].text.strip()
    # Strip markdown code fences if present
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None


def _extract_records_with_mapping(
    rows: list[tuple],
    col_map: dict,
    sheet_name: str,
    file_name: str,
    record_type: str = "AI解析",
) -> list[dict]:
    """Use the AI-detected column mapping to extract records from rows."""
    item_col = col_map.get("item_name")
    if item_col is None:
        return []

    tag_col    = col_map.get("tag_no")
    po_col     = col_map.get("po_no")
    vendor_col = col_map.get("sub_vendor")
    ros_col    = col_map.get("ros")
    eta_col    = col_map.get("eta")
    ata_col    = col_map.get("ata")
    qty_col    = col_map.get("qty")

    # Skip rows that look like headers
    from parser import _find_header_end
    data_start = _find_header_end(rows, {"item": item_col})

    records = []
    seen: set = set()
    current_tag = ""

    for row in rows[data_start:]:
        def g(ci):
            return row[ci] if ci is not None and len(row) > ci else None

        raw_tag = _str(g(tag_col))
        if raw_tag:
            current_tag = raw_tag.split("\n")[0]

        item_name = _str(g(item_col)).split("\n")[0].strip()
        if not item_name:
            continue
        if re.search(r"合計|小計|total|subtotal|^(文件|設備名稱|item|description|材料名稱|編碼)$",
                     item_name, re.I):
            continue

        # De-duplicate multi-revision rows
        key = (sheet_name, item_name)
        if key in seen:
            continue
        seen.add(key)

        eta = _to_date(g(eta_col))
        ata = _to_date(g(ata_col))
        ros = _to_date(g(ros_col))

        records.append({
            "source_file":       file_name,
            "sheet":             sheet_name,
            "mr_no":             "",
            "mr_name":           "",
            "po_no":             _str(g(po_col)).split("\n")[0],
            "tag_no":            current_tag,
            "vendor":            "",
            "delivery_location": "",
            "ros":               ros,
            "item_name":         item_name,
            "qty":               _str(g(qty_col)),
            "sub_vendor":        _str(g(vendor_col)).split("\n")[0],
            "eta":               eta,
            "ata":               ata,
            "effective_delivery": ata or eta or ros,
            "sub_order_received": None,
            "record_type":        record_type,
        })

    return records


# ── Public entry point ────────────────────────────────────────────────────────

def parse_with_ai(path: Path, api_key: str) -> tuple[list[dict], list[dict]]:
    """
    Parse any Excel file using Claude for column detection.

    Returns:
        (records, mapping_log)
        mapping_log: list of dicts describing what Claude detected per sheet
    """
    import pandas as pd

    # Skip if it's a known ESP format (already handled perfectly)
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xls"):
        wb = openpyxl.load_workbook(str(path), data_only=True)
        sheets = set(wb.sheetnames)
        if sheets & ESP_SHEETS:
            wb.close()
            return parse_esp(path), []
    else:
        return [], []

    all_records = []
    mapping_log = []

    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows or ws.max_row < 3:
            continue

        # Build sample: up to 5 header rows + first 5 data rows
        non_empty = [r for r in rows[:12] if any(c for c in r)]
        sample_text = _sheet_to_text_sample(non_empty, max_rows=10)

        col_map_raw = _call_claude(api_key, sample_text, ws.title)
        if not col_map_raw:
            mapping_log.append({"sheet": ws.title, "status": "AI無法解析", "mapping": {}})
            continue

        # Filter out null values
        col_map = {k: v for k, v in col_map_raw.items() if v is not None}
        mapping_log.append({
            "sheet":   ws.title,
            "status":  "成功" if col_map.get("item_name") is not None else "找不到材料名稱欄",
            "mapping": col_map,
        })

        if col_map.get("item_name") is None:
            continue

        records = _extract_records_with_mapping(rows, col_map, ws.title, path.name)
        all_records.extend(records)

    wb.close()
    return all_records, mapping_log


def parse_uploaded_with_ai(file_bytes: bytes, file_name: str, api_key: str) -> tuple[list[dict], list[dict]]:
    """Parse an in-memory uploaded file."""
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=Path(file_name).suffix, delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = Path(tmp.name)
    try:
        return parse_with_ai(tmp_path, api_key)
    finally:
        tmp_path.unlink(missing_ok=True)
