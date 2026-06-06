"""
ESP / Sub-order Excel parser.
Supports the multi-sheet ESP format and simple sub-order list format.
"""
import re
from datetime import datetime, date
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd


# ── helpers ──────────────────────────────────────────────────────────────────

def _to_date(val: Any) -> date | None:
    if isinstance(val, (datetime, date)):
        return val.date() if isinstance(val, datetime) else val
    if isinstance(val, str):
        for fmt in ("%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _str(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip()


# ── ESP format parser ─────────────────────────────────────────────────────────

ESP_SHEETS = {"Stationary", "Rotating", "Piping", "Pipeline", "Electrical", "Instrument"}

# Row offsets within each PO block (0-indexed from block start row)
_PO_LABEL_COL = 1   # column B
_PO_VAL_COL   = 2   # column C
_PO_LABEL_MAP = {
    "M/R No.":          "mr_no",
    "M/R Name":         "mr_name",
    "P/O No.":          "po_no",
    "Vendor Name":      "vendor",
    "Delivery Date":    "delivery_date",
    "Delivery Location":"delivery_location",
}
_ITEM_NAME_COL     = 7   # H  – Sub-Order Item
_QTY_COL           = 8   # I
_SUB_ORDER_REC_ACT = 14  # O  – Sub-Order Received Actual
_SUB_VENDOR_COL    = 15  # P
_ETA_COL           = 35  # AJ – ETA at destination
_ATA_COL           = 36  # AK – Actual time of arrival
_ROS_COL           = 37  # AL – Required on site


def _parse_esp_sheet(ws, sheet_name: str, file_name: str) -> list[dict]:
    rows = list(ws.iter_rows(values_only=True))
    records = []

    # Collect PO header info + material rows.
    # A new PO block starts when column-A (index 0) contains an integer.
    po_info: dict = {}
    in_block = False

    for r_idx, row in enumerate(rows):
        # detect new PO block
        if isinstance(row[0], (int, float)) and row[0] not in (None,):
            # flush previous block items already collected (nothing to flush here,
            # items are emitted row-by-row below with po_info captured)
            po_info = {}
            in_block = True
            _update_po_info(row, po_info)
            # First material row is on the same row
            _maybe_add_item(row, po_info, sheet_name, file_name, records)
            continue

        if in_block:
            # Accumulate PO label-value pairs (cols B/C)
            if row[_PO_LABEL_COL] in _PO_LABEL_MAP:
                key = _PO_LABEL_MAP[row[_PO_LABEL_COL]]
                val = row[_PO_VAL_COL]
                if key in ("delivery_date",):
                    po_info[key] = _to_date(val)
                else:
                    po_info[key] = _str(val)

            # Every row in the block that has an item name contributes a record
            _maybe_add_item(row, po_info, sheet_name, file_name, records)

    return records


def _update_po_info(row, po_info: dict):
    if row[_PO_LABEL_COL] in _PO_LABEL_MAP:
        key = _PO_LABEL_MAP[row[_PO_LABEL_COL]]
        po_info[key] = row[_PO_VAL_COL]


def _maybe_add_item(row, po_info: dict, sheet: str, file_name: str, out: list):
    item_name = _str(row[_ITEM_NAME_COL]) if len(row) > _ITEM_NAME_COL else ""
    if not item_name:
        return

    eta  = _to_date(row[_ETA_COL])  if len(row) > _ETA_COL  else None
    ata  = _to_date(row[_ATA_COL])  if len(row) > _ATA_COL  else None
    ros  = _to_date(row[_ROS_COL])  if len(row) > _ROS_COL  else None
    sub_ord_received = _to_date(row[_SUB_ORDER_REC_ACT]) if len(row) > _SUB_ORDER_REC_ACT else None

    # Effective delivery date: ATA > ETA > sub_order_received_actual
    effective_delivery = ata or eta

    out.append({
        "source_file":   file_name,
        "sheet":         sheet,
        "mr_no":         po_info.get("mr_no", ""),
        "mr_name":       po_info.get("mr_name", ""),
        "po_no":         _str(po_info.get("po_no", "")),
        "vendor":        po_info.get("vendor", ""),
        "delivery_location": po_info.get("delivery_location", ""),
        "ros":           ros or po_info.get("delivery_date"),   # required on site
        "item_name":     item_name,
        "qty":           _str(row[_QTY_COL]) if len(row) > _QTY_COL else "",
        "sub_vendor":    _str(row[_SUB_VENDOR_COL]) if len(row) > _SUB_VENDOR_COL else "",
        "eta":           eta,
        "ata":           ata,
        "effective_delivery": effective_delivery,
        "sub_order_received": sub_ord_received,
        "record_type":   "ESP",
    })


def parse_esp(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(str(path), data_only=True)
    records = []
    for sname in wb.sheetnames:
        if sname in ESP_SHEETS:
            records.extend(_parse_esp_sheet(wb[sname], sname, path.name))
    return records


# ── Sub-order list parser (vendor-prepared, simpler format) ───────────────────
# Expected columns (flexible detection):
#   No | Item | Description/Material | Qty | Unit | Sub-Order PO | Sub-Vendor |
#   Expected PO issue | PO Deadline | Actual PO issue |
#   Required Date | Expected Date | Actual Date | Delivered Qty | Note

_SUBORDER_DATE_PATTERNS = [
    re.compile(r"required|ros|need", re.I),
    re.compile(r"expected.*date|exp.*date|eta", re.I),
    re.compile(r"actual.*date|ata|delivered.*date", re.I),
]

def _detect_suborder_columns(header_rows: list[tuple]) -> dict:
    """Best-effort column index detection from header rows."""
    flat = {}
    for hr in header_rows:
        for ci, cell in enumerate(hr):
            s = _str(cell).lower()
            if not s:
                continue
            if re.search(r"item|description|material|品名", s) and "item" not in flat:
                flat["item"] = ci
            if re.search(r"po.?no|purchase.*order|sub.?order.?po", s) and "po_no" not in flat:
                flat["po_no"] = ci
            if re.search(r"sub.?vendor|vendor|supplier|廠商", s) and "sub_vendor" not in flat:
                flat["sub_vendor"] = ci
            if re.search(r"required|ros|need.*date|需求", s) and "ros" not in flat:
                flat["ros"] = ci
            if re.search(r"expected.*date|exp.*date|eta|預計.*到", s) and "eta" not in flat:
                flat["eta"] = ci
            if re.search(r"actual.*date|ata|actual.*deliver|實際", s) and "ata" not in flat:
                flat["ata"] = ci
            if re.search(r"qty|quantity|數量", s) and "qty" not in flat:
                flat["qty"] = ci
            if re.search(r"mr.?no|requisition|請購", s) and "mr_no" not in flat:
                flat["mr_no"] = ci
            if re.search(r"po.?issue|po.*deadline", s) and "po_issue" not in flat:
                flat["po_issue"] = ci
    return flat


def parse_suborder_list(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(str(path), data_only=True)
    records = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        # Find header rows (usually first 1-3 non-empty rows)
        header_rows = [r for r in rows[:5] if any(c for c in r)]
        col_map = _detect_suborder_columns(header_rows)

        if "item" not in col_map:
            continue  # not a recognisable sub-order sheet

        n_header = len(header_rows)
        for row in rows[n_header:]:
            item = _str(row[col_map["item"]]) if col_map.get("item") is not None and len(row) > col_map["item"] else ""
            if not item:
                continue
            get = lambda key: row[col_map[key]] if key in col_map and len(row) > col_map[key] else None
            eta = _to_date(get("eta"))
            ata = _to_date(get("ata"))
            ros = _to_date(get("ros"))
            records.append({
                "source_file":        path.name,
                "sheet":              ws.title,
                "mr_no":              _str(get("mr_no")),
                "mr_name":            "",
                "po_no":              _str(get("po_no")),
                "vendor":             "",
                "delivery_location":  "",
                "ros":                ros,
                "item_name":          item,
                "qty":                _str(get("qty")),
                "sub_vendor":         _str(get("sub_vendor")),
                "eta":                eta,
                "ata":                ata,
                "effective_delivery": ata or eta,
                "sub_order_received": None,
                "record_type":        "SubOrder",
            })
    return records


# ── Auto-detect file type and dispatch ───────────────────────────────────────

def _is_esp_file(path: Path) -> bool:
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        sheets = set(wb.sheetnames)
        wb.close()
        return bool(sheets & ESP_SHEETS)
    except Exception:
        return False


def load_all_files(folder: str | Path) -> pd.DataFrame:
    folder = Path(folder)
    all_records: list[dict] = []
    for f in sorted(folder.glob("*.xlsx")) + sorted(folder.glob("*.xls")):
        try:
            if _is_esp_file(f):
                all_records.extend(parse_esp(f))
            else:
                all_records.extend(parse_suborder_list(f))
        except Exception as e:
            print(f"[WARN] Could not parse {f.name}: {e}")

    df = pd.DataFrame(all_records) if all_records else pd.DataFrame(columns=[
        "source_file", "sheet", "mr_no", "mr_name", "po_no", "vendor",
        "delivery_location", "ros", "item_name", "qty", "sub_vendor",
        "eta", "ata", "effective_delivery", "sub_order_received", "record_type",
    ])
    # normalise date columns
    for col in ("ros", "eta", "ata", "effective_delivery", "sub_order_received"):
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df
