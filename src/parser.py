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
        val = val.strip().split("\n")[0].strip()
        for fmt in ("%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(val, fmt).date()
            except ValueError:
                pass
    return None


def _str(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip()


# ── ESP format parser ─────────────────────────────────────────────────────────

ESP_SHEETS = {"Stationary", "Rotating", "Piping", "Pipeline", "Electrical", "Instrument"}

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
    po_info: dict = {}
    in_block = False

    for row in rows:
        if isinstance(row[0], (int, float)) and row[0] not in (None,):
            po_info = {}
            in_block = True
            _update_po_info(row, po_info)
            _maybe_add_esp_item(row, po_info, sheet_name, file_name, records)
            continue

        if in_block:
            if row[_PO_LABEL_COL] in _PO_LABEL_MAP:
                key = _PO_LABEL_MAP[row[_PO_LABEL_COL]]
                val = row[_PO_VAL_COL]
                po_info[key] = _to_date(val) if key == "delivery_date" else _str(val)

            _maybe_add_esp_item(row, po_info, sheet_name, file_name, records)

    return records


def _update_po_info(row, po_info: dict):
    if row[_PO_LABEL_COL] in _PO_LABEL_MAP:
        key = _PO_LABEL_MAP[row[_PO_LABEL_COL]]
        po_info[key] = row[_PO_VAL_COL]


def _maybe_add_esp_item(row, po_info: dict, sheet: str, file_name: str, out: list):
    item_name = _str(row[_ITEM_NAME_COL]) if len(row) > _ITEM_NAME_COL else ""
    if not item_name:
        return

    eta  = _to_date(row[_ETA_COL])  if len(row) > _ETA_COL  else None
    ata  = _to_date(row[_ATA_COL])  if len(row) > _ATA_COL  else None
    ros  = _to_date(row[_ROS_COL])  if len(row) > _ROS_COL  else None

    out.append({
        "source_file":       file_name,
        "sheet":             sheet,
        "mr_no":             po_info.get("mr_no", ""),
        "mr_name":           po_info.get("mr_name", ""),
        "po_no":             _str(po_info.get("po_no", "")),
        "tag_no":            "",
        "vendor":            po_info.get("vendor", ""),
        "delivery_location": po_info.get("delivery_location", ""),
        "ros":               ros or po_info.get("delivery_date"),
        "item_name":         item_name,
        "qty":               _str(row[_QTY_COL]) if len(row) > _QTY_COL else "",
        "sub_vendor":        _str(row[_SUB_VENDOR_COL]) if len(row) > _SUB_VENDOR_COL else "",
        "eta":               eta,
        "ata":               ata,
        "effective_delivery": ata or eta,
        "sub_order_received": _to_date(row[_SUB_ORDER_REC_ACT]) if len(row) > _SUB_ORDER_REC_ACT else None,
        "record_type":       "ESP",
    })


def parse_esp(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(str(path), data_only=True)
    records = []
    for sname in wb.sheetnames:
        if sname in ESP_SHEETS:
            records.extend(_parse_esp_sheet(wb[sname], sname, path.name))
    return records


# ── Sub-order list parser ─────────────────────────────────────────────────────
#
# Handles vendor-prepared sub-order lists.  The column layout is detected
# from the first 3 header rows using keyword matching.
#
# Typical layout (from the screenshot):
#   No. | ITEM(tag) | Exp RFP | Act RFP | DESCRIPTION(item_name) | MATERIAL |
#   Q'ty | UNIT | SUB-ORDER PO NO. | SUB-VENDOR | MFG COUNTRY |
#   Expected PO issue | PO Deadline | Actual PO issue |
#   REQUIRED Date(ros) | EXPECTED Date(eta) | ACTUAL Date(ata) | Delivered Q'ty | NOTE

def _detect_suborder_columns(header_rows: list[tuple]) -> dict:
    """
    Scan up to 5 header rows, match each cell against keyword patterns,
    and return a best-guess column-index map.
    Supports both English ESP/SubOrder formats and Chinese 送審管制 formats.
    Higher priority wins when two rows both match the same key.
    """
    col: dict = {}

    patterns = [
        # ── item / material name ──────────────────────────────────────────────
        ("item", re.compile(r"^item$", re.I),                                             1),
        ("item", re.compile(r"description|desc\b|品名|材料名稱?|文件$|設備名稱", re.I),    2),
        ("item", re.compile(r"^設備名稱|^material\s*name|^document\s*name", re.I),         3),

        # ── tag / equipment number ────────────────────────────────────────────
        ("tag_no", re.compile(r"^item$|tag.?no|equip(ment)?(\s*no)?$", re.I),            1),

        # ── document / PO number ─────────────────────────────────────────────
        # Must NOT match RFP/issue/deadline/date columns
        ("po_no", re.compile(r"sub.?order.?po|^po[\s#\.]*no\b|po\s*number|po\s*num\b",
                              re.I),                                                       1),
        ("po_no", re.compile(r"^purchase\s*order\s*no|^編碼$|^文件編碼|^doc(ument)?\s*(no|code)",
                              re.I),                                                       2),

        # ── sub-vendor ───────────────────────────────────────────────────────
        ("sub_vendor", re.compile(r"sub.?vendor|vendor|supplier|廠商|統包商", re.I),       1),

        # ── quantity / unit ───────────────────────────────────────────────────
        ("qty",  re.compile(r"q'?ty|quantity|數量", re.I),                                 1),
        ("unit", re.compile(r"^unit$|單位", re.I),                                         1),

        # ── Required On Site (ROS) ────────────────────────────────────────────
        # English
        ("ros", re.compile(r"required|ros\b|need.*date", re.I),                           1),
        ("ros", re.compile(r"required.*date|need.*on.*site", re.I),                        2),
        # Chinese 送審管制格式
        ("ros", re.compile(r"預定進場|需求.*日期|預計進場", re.I),                           3),

        # ── Expected delivery / ETA ───────────────────────────────────────────
        # Exclude RFP, issue, deadline, 送審 columns
        ("eta", re.compile(
            r"(?!.*rfp)(?!.*issue)(?!.*deadline)(?!.*送審)expected.*date|^eta$|預計.*到",
            re.I),                                                                          1),
        ("eta", re.compile(r"^expected\s*date$", re.I),                                    2),
        ("eta", re.compile(r"預定廠驗|預計.*交貨|預計.*到料", re.I),                         3),

        # ── Actual delivery / ATA ─────────────────────────────────────────────
        ("ata", re.compile(
            r"(?!.*rfp)(?!.*issue)actual.*date|^ata$|實際.*到",
            re.I),                                                                          1),
        ("ata", re.compile(r"^actual\s*date$", re.I),                                      2),
        ("ata", re.compile(r"實際.*進場|實際.*交貨|實際.*到料", re.I),                        3),

        # ── misc ─────────────────────────────────────────────────────────────
        ("delivered_qty", re.compile(r"delivered.*q|交付.*數", re.I),                      1),
        ("mr_no",         re.compile(r"mr.?no|requisition|請購", re.I),                    1),
        ("po_issue_exp",  re.compile(r"expected.*po|exp.*po.*issue|預定送審", re.I),        1),
        ("po_deadline",   re.compile(r"po.*deadline|deadline", re.I),                      1),
        ("po_issue_act",  re.compile(r"actual.*po|po.*actual|act.*po", re.I),              1),
        ("spec",          re.compile(r"規範|spec(ification)?", re.I),                      1),
        ("warning_days",  re.compile(r"預警|warning.*day|送審.*延遲", re.I),               1),
    ]

    priority: dict[str, int] = {}

    for hr in header_rows:
        for ci, cell in enumerate(hr):
            s = _str(cell).split("\n")[0]   # only first line of merged headers
            if not s:
                continue
            sl = s.lower()
            for key, pat, pri in patterns:
                if pat.search(sl):
                    if priority.get(key, 0) < pri:
                        col[key] = ci
                        priority[key] = pri

    return col


def _validate_columns(col: dict, data_rows: list[tuple]) -> dict:
    """
    Post-detection sanity check: if a column that should contain text (po_no,
    sub_vendor, item_name) instead contains mostly date-like values, discard it.
    """
    _date_pat = re.compile(
        r"^\d{1,2}[/-]\w{2,3}[/-]\d{2,4}$|^\d{4}[/-]\d{2}[/-]\d{2}$|"
        r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$", re.I
    )
    sample = data_rows[:10]

    def _is_date_col(ci: int) -> bool:
        vals = [_str(r[ci]) for r in sample if len(r) > ci and _str(r[ci])]
        if not vals:
            return False
        date_hits = sum(1 for v in vals if _date_pat.match(v) or isinstance(v, (datetime, date)))
        return date_hits / len(vals) >= 0.6

    for key in ("po_no", "mr_no"):
        if key in col and _is_date_col(col[key]):
            del col[key]

    return col


def _find_header_end(rows: list[tuple], col_map: dict) -> int:
    """Return index of first data row (skip header rows)."""
    # A data row typically has a numeric value in col 0 or a non-empty item cell
    item_col = col_map.get("item", col_map.get("tag_no", 4))
    for i, row in enumerate(rows[:6]):
        if i == 0:
            continue
        item_val = _str(row[item_col]) if len(row) > item_col else ""
        no_val   = row[0] if row else None
        # first row where col-0 is a number  OR item cell is clearly a material name
        if isinstance(no_val, (int, float)):
            return i
        if item_val and not re.search(r"date|date|vendor|item|qty|required|expected|actual|no\.|unit", item_val, re.I):
            return i
    return min(3, len(rows))


def parse_suborder_list(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(str(path), data_only=True)
    records = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        header_rows = [r for r in rows[:4] if any(c for c in r)]
        col_map = _detect_suborder_columns(header_rows)

        # Need at least a description/item column
        if "item" not in col_map and "tag_no" not in col_map:
            continue

        item_col   = col_map.get("item", col_map.get("tag_no"))
        tag_col    = col_map.get("tag_no") if col_map.get("tag_no") != item_col else None
        data_start = _find_header_end(rows, col_map)
        col_map    = _validate_columns(col_map, rows[data_start:])

        # Carry-forward for ITEM (tag) column which may span multiple rows
        current_tag = ""

        for row in rows[data_start:]:
            if not any(c for c in row):
                continue

            # Update carry-forward tag
            raw_tag = _str(row[tag_col]) if tag_col is not None and len(row) > tag_col else ""
            if raw_tag:
                current_tag = " / ".join(raw_tag.splitlines())

            item_name = _str(row[item_col]) if len(row) > item_col else ""
            if not item_name:
                continue
            # Skip summary / status rows
            if re.search(r"sub.?order.*status|delivery status|total|合計", item_name, re.I):
                continue

            def gcol(key):
                ci = col_map.get(key)
                return row[ci] if ci is not None and len(row) > ci else None

            eta = _to_date(gcol("eta"))
            ata = _to_date(gcol("ata"))
            ros = _to_date(gcol("ros"))

            records.append({
                "source_file":       path.name,
                "sheet":             ws.title,
                "mr_no":             "",
                "mr_name":           "",
                "po_no":             _str(gcol("po_no")),
                "tag_no":            current_tag,
                "vendor":            "",
                "delivery_location": "",
                "ros":               ros,
                "item_name":         item_name,
                "qty":               _str(gcol("qty")),
                "sub_vendor":        _str(gcol("sub_vendor")),
                "eta":               eta,
                "ata":               ata,
                "effective_delivery": ata or eta,
                "sub_order_received": None,
                "record_type":        "SubOrder",
            })
    return records


# ── 送審管制表 parser ─────────────────────────────────────────────────────────
# Format: 新竹/各廠 材料設備送審管制總表
# Each equipment item spans multiple rows (one per revision version).
# The "main" row is identified by having a non-empty 項次 (sequence no) AND 設備名稱.
# Key columns (detected by header keywords):
#   設備名稱 → item_name
#   文件編碼  → po_no
#   預定進場  → ros   (required on site)
#   預定廠驗  → eta   (expected delivery / factory inspection)
#   廠商 / 統包商 → sub_vendor

_SUBMISSION_SHEET_PAT = re.compile(r"送審|管制|submission|submittal", re.I)

def _is_submission_control(wb) -> bool:
    """Detect 送審管制 format by checking sheet names OR first-row cell content."""
    # Check sheet names first (fast)
    if any(_SUBMISSION_SHEET_PAT.search(s) for s in wb.sheetnames):
        return True
    # Check title cells in first 2 rows of each sheet
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=2, values_only=True):
            for cell in row:
                if cell and _SUBMISSION_SHEET_PAT.search(_str(cell)):
                    return True
    return False


def _parse_submission_sheet(ws, sheet_name: str, file_name: str) -> list[dict]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    # Header rows: scan first 5 non-empty rows
    header_rows = [r for r in rows[:6] if any(c for c in r)][:5]
    col_map = _detect_suborder_columns(header_rows)

    item_col = col_map.get("item")
    if item_col is None:
        return []

    po_col      = col_map.get("po_no")
    ros_col     = col_map.get("ros")
    eta_col     = col_map.get("eta")
    ata_col     = col_map.get("ata")
    vendor_col  = col_map.get("sub_vendor")
    warn_col    = col_map.get("warning_days")

    # Find data start: first row where col 0 or a "no." column is numeric
    # or where item_col has a non-header value
    data_start = _find_header_end(rows, col_map)

    records = []
    seen_items: set = set()

    for row in rows[data_start:]:
        item_raw = _str(row[item_col]) if len(row) > item_col else ""
        item_name = item_raw.split("\n")[0].strip()   # first line only
        if not item_name:
            continue
        if re.search(r"合計|小計|total|subtotal|^\d+$", item_name, re.I):
            continue
        # Skip header-like rows that slipped through
        if re.search(r"^(文件|設備名稱|item|description|材料名稱|編碼)$", item_name, re.I):
            continue

        # De-duplicate: 送審管制表 has multiple revision rows per item.
        # Keep only the first occurrence (latest data is on the main/first row).
        key = (sheet_name, item_name)
        if key in seen_items:
            continue
        seen_items.add(key)

        def g(ci):
            return row[ci] if ci is not None and len(row) > ci else None

        ros = _to_date(g(ros_col))
        eta = _to_date(g(eta_col))
        ata = _to_date(g(ata_col))

        records.append({
            "source_file":        file_name,
            "sheet":              sheet_name,
            "mr_no":              "",
            "mr_name":            "",
            "po_no":              _str(g(po_col)).split("\n")[0],
            "tag_no":             "",
            "vendor":             "",
            "delivery_location":  "",
            "ros":                ros,
            "item_name":          item_name,
            "qty":                "",
            "sub_vendor":         _str(g(vendor_col)).split("\n")[0],
            "eta":                eta,
            "ata":                ata,
            "effective_delivery": ata or eta or ros,
            "sub_order_received": None,
            "record_type":        "送審管制",
        })

    return records


def parse_submission_control(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(str(path), data_only=True)
    records = []
    for sname in wb.sheetnames:
        ws = wb[sname]
        if ws.max_row < 3:
            continue
        records.extend(_parse_submission_sheet(ws, sname, path.name))
    return records


# ── PDF parser ───────────────────────────────────────────────────────────────

def _pdf_sheet_rows(path: Path) -> list[tuple[str, list[tuple]]]:
    """Extract tables from each PDF page, return list of (sheet_name, rows)."""
    import pdfplumber
    sheets = []
    with pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for j, table in enumerate(tables):
                rows = [tuple(cell if cell is not None else "" for cell in row)
                        for row in table if any(cell for cell in row)]
                if rows:
                    sheets.append((f"Page{i+1}_T{j+1}", rows))
    return sheets


def parse_pdf(path: Path) -> list[dict]:
    records = []
    for sheet_name, rows in _pdf_sheet_rows(path):
        if not rows:
            continue
        header_rows = rows[:4]
        col_map = _detect_suborder_columns(header_rows)
        if "item" not in col_map and "tag_no" not in col_map:
            continue
        item_col   = col_map.get("item", col_map.get("tag_no"))
        tag_col    = col_map.get("tag_no") if col_map.get("tag_no") != item_col else None
        data_start = _find_header_end(rows, col_map)
        col_map    = _validate_columns(col_map, rows[data_start:])
        current_tag = ""
        for row in rows[data_start:]:
            raw_tag = _str(row[tag_col]) if tag_col is not None and len(row) > tag_col else ""
            if raw_tag:
                current_tag = " / ".join(raw_tag.splitlines())
            item_name = _str(row[item_col]) if len(row) > item_col else ""
            if not item_name:
                continue
            if re.search(r"sub.?order.*status|delivery status|total|合計", item_name, re.I):
                continue
            def gcol(key, r=row):
                ci = col_map.get(key)
                return r[ci] if ci is not None and len(r) > ci else None
            eta = _to_date(gcol("eta"))
            ata = _to_date(gcol("ata"))
            ros = _to_date(gcol("ros"))
            records.append({
                "source_file": path.name, "sheet": sheet_name,
                "mr_no": "", "mr_name": "",
                "po_no": _str(gcol("po_no")), "tag_no": current_tag,
                "vendor": "", "delivery_location": "", "ros": ros,
                "item_name": item_name, "qty": _str(gcol("qty")),
                "sub_vendor": _str(gcol("sub_vendor")),
                "eta": eta, "ata": ata,
                "effective_delivery": ata or eta,
                "sub_order_received": None, "record_type": "PDF",
            })
    return records


# ── Word (.docx) parser ───────────────────────────────────────────────────────

def _docx_sheet_rows(path: Path) -> list[tuple[str, list[tuple]]]:
    """Extract tables from a Word document."""
    from docx import Document
    doc = Document(str(path))
    sheets = []
    for i, table in enumerate(doc.tables):
        rows = []
        for row in table.rows:
            cells = tuple(cell.text.strip() for cell in row.cells)
            if any(cells):
                rows.append(cells)
        if rows:
            sheets.append((f"Table{i+1}", rows))
    return sheets


def parse_docx(path: Path) -> list[dict]:
    records = []
    for sheet_name, rows in _docx_sheet_rows(path):
        if not rows:
            continue
        header_rows = rows[:4]
        col_map = _detect_suborder_columns(header_rows)
        if "item" not in col_map and "tag_no" not in col_map:
            continue
        item_col   = col_map.get("item", col_map.get("tag_no"))
        tag_col    = col_map.get("tag_no") if col_map.get("tag_no") != item_col else None
        data_start = _find_header_end(rows, col_map)
        col_map    = _validate_columns(col_map, rows[data_start:])
        current_tag = ""
        for row in rows[data_start:]:
            raw_tag = _str(row[tag_col]) if tag_col is not None and len(row) > tag_col else ""
            if raw_tag:
                current_tag = " / ".join(raw_tag.splitlines())
            item_name = _str(row[item_col]) if len(row) > item_col else ""
            if not item_name:
                continue
            if re.search(r"sub.?order.*status|delivery status|total|合計", item_name, re.I):
                continue
            def gcol(key, r=row):
                ci = col_map.get(key)
                return r[ci] if ci is not None and len(r) > ci else None
            eta = _to_date(gcol("eta"))
            ata = _to_date(gcol("ata"))
            ros = _to_date(gcol("ros"))
            records.append({
                "source_file": path.name, "sheet": sheet_name,
                "mr_no": "", "mr_name": "",
                "po_no": _str(gcol("po_no")), "tag_no": current_tag,
                "vendor": "", "delivery_location": "", "ros": ros,
                "item_name": item_name, "qty": _str(gcol("qty")),
                "sub_vendor": _str(gcol("sub_vendor")),
                "eta": eta, "ata": ata,
                "effective_delivery": ata or eta,
                "sub_order_received": None, "record_type": "Word",
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


def parse_file(path: Path) -> list[dict]:
    """Auto-detect format and dispatch to the right parser."""
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xls"):
        if _is_esp_file(path):
            return parse_esp(path)
        try:
            wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
            is_submission = _is_submission_control(wb)
            wb.close()
        except Exception:
            is_submission = False
        if is_submission:
            return parse_submission_control(path)
        return parse_suborder_list(path)
    if ext == ".pdf":
        return parse_pdf(path)
    if ext in (".docx", ".doc"):
        return parse_docx(path)
    return []


def load_all_files(folder: str | Path) -> pd.DataFrame:
    folder = Path(folder)
    all_records: list[dict] = []
    globs = ["*.xlsx", "*.xls", "*.pdf", "*.docx", "*.doc"]
    files = []
    for g in globs:
        files.extend(sorted(folder.glob(g)))
    for f in files:
        try:
            all_records.extend(parse_file(f))
        except Exception as e:
            print(f"[WARN] Could not parse {f.name}: {e}")

    df = pd.DataFrame(all_records) if all_records else pd.DataFrame(columns=[
        "source_file", "sheet", "mr_no", "mr_name", "po_no", "tag_no", "vendor",
        "delivery_location", "ros", "item_name", "qty", "sub_vendor",
        "eta", "ata", "effective_delivery", "sub_order_received", "record_type",
    ])
    for col in ("ros", "eta", "ata", "effective_delivery", "sub_order_received"):
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df
