import os
import re
from datetime import date, datetime
import openpyxl
import requests

TEMPLATE_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Blinkit PO Tracker.xlsx"))
BLOB_PATHNAME = "Blinkit PO Tracker.xlsx"
BLOB_UPLOAD_URL = f"https://blob.vercel-storage.com/{BLOB_PATHNAME}"
BLOB_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _blob_token():
    return os.environ.get("BLOB_READ_WRITE_TOKEN")


def _blob_read_url(token: str) -> str:
    # Token format: vercel_blob_rw_<storeId>_<secret> — the store id maps
    # directly to the public read host, so no separate lookup call is needed.
    m = re.match(r"vercel_blob_rw_([a-zA-Z0-9]+)_", token)
    if not m:
        raise RuntimeError("Could not parse Blob store id from BLOB_READ_WRITE_TOKEN")
    return f"https://{m.group(1)}.public.blob.vercel-storage.com/{BLOB_PATHNAME}"


def _download_blob(token: str):
    resp = requests.get(_blob_read_url(token), headers={"Authorization": f"Bearer {token}"}, timeout=15)
    if resp.status_code == 404:
        return None
    resp.raise_for_status()
    return resp.content


def _upload_blob(token: str, data: bytes) -> None:
    resp = requests.put(
        BLOB_UPLOAD_URL,
        data=data,
        headers={
            "Authorization": f"Bearer {token}",
            "x-api-version": "7",
            "x-content-type": BLOB_MIME,
            "x-add-random-suffix": "0",
            "x-allow-overwrite": "1",
        },
        timeout=30,
    )
    resp.raise_for_status()


def get_live_path():
    """
    Vercel serverless functions are stateless across invocations -- /tmp is
    NOT shared between requests (a different, freshly-initialized container
    can serve any given request). Using /tmp alone as the "live" tracker
    caused uploaded rows/dates to intermittently vanish depending which
    container answered a later request.

    The single source of truth is now a file in Vercel Blob storage, shared
    by every container. Each call re-downloads the current copy into /tmp
    purely so openpyxl has a local file handle to read/write.
    """
    token = _blob_token()
    if not token:
        # Local dev / no Blob store connected yet: fall back to the bundled
        # template on local disk (previous behaviour).
        return TEMPLATE_PATH

    local_path = os.path.join("/tmp", "Blinkit PO Tracker.xlsx")
    data = _download_blob(token)
    if data is None:
        # First run ever: seed the Blob store from the bundled template.
        with open(TEMPLATE_PATH, "rb") as f:
            data = f.read()
        _upload_blob(token, data)
    with open(local_path, "wb") as f:
        f.write(data)
    return local_path


def _save(wb, path) -> None:
    """Save the workbook locally, then push it back to Blob storage (if
    configured) so every other container sees the update immediately."""
    wb.save(path)
    token = _blob_token()
    if token:
        with open(path, "rb") as f:
            _upload_blob(token, f.read())

# Maps our extracted keys → Excel column headers (exact header text in row 1)
COLUMN_MAP = {
    "CHAINS": "CHAINS",
    "SITE CODE": "SITE CODE ",       # note trailing space in original header
    "VENDOR CODE": "VENDOR CODE",
    "VENDOR NAME": "VENDOR NAME",
    "PO NO": "PO NO",
    "PO DATE": "PO DATE",
    "DELIVERY DATE": "DELIVERY DATE",
    "ARTICLE DESCRIPTION": "ARTICLE DESCRIPTION",
    "TOTAL PCS": "TOTAL PCS",
    "LANDING PRICE": "LANDING PRICE",
    "TOTAL BASIC PO VALUE WITH TAX": "TOTAL BASIC PO VALUE WITH TAX",
}


def _load_workbook():
    path = get_live_path()
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")
    return openpyxl.load_workbook(path), path


def _get_header_map(ws) -> dict[str, int]:
    """Return {header_text: col_index (1-based)} from row 1."""
    header_map = {}
    for col in ws.iter_cols(min_row=1, max_row=1):
        for cell in col:
            if cell.value is not None:
                header_map[str(cell.value).strip()] = cell.column
    return header_map


def append_rows(rows: list[dict]) -> dict:
    """Append extracted PO rows to Sheet1 and save."""
    wb, path = _load_workbook()
    ws = wb["Sheet1"]

    # Build header→col map (strip spaces for matching)
    raw_headers = {}
    for col in ws.iter_cols(min_row=1, max_row=1):
        for cell in col:
            if cell.value is not None:
                raw_headers[str(cell.value).strip()] = cell.column

    # Find the next empty row
    next_row = ws.max_row + 1
    # Check if last row is actually empty (openpyxl can report wrong max_row)
    while next_row > 2:
        row_vals = [ws.cell(row=next_row - 1, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in row_vals):
            break
        next_row -= 1

    written = 0
    for row_data in rows:
        for key, header in COLUMN_MAP.items():
            # Match header ignoring trailing spaces
            col_idx = None
            for h, ci in raw_headers.items():
                if h.upper() == header.strip().upper():
                    col_idx = ci
                    break
            if col_idx is None:
                continue

            value = row_data.get(key)
            cell = ws.cell(row=next_row, column=col_idx, value=value)

            # Format dates
            if isinstance(value, (date, datetime)):
                cell.number_format = "MM/DD/YYYY"
        next_row += 1
        written += 1

    _save(wb, path)
    return {"appended": written, "path": path}


def clear_sheet() -> dict:
    """Delete all data rows from Sheet1 (keep header row 1)."""
    wb, path = _load_workbook()
    ws = wb["Sheet1"]

    rows_deleted = 0
    # Delete from bottom up to avoid index shifting
    for row in range(ws.max_row, 1, -1):
        row_vals = [ws.cell(row=row, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in row_vals):
            ws.delete_rows(row)
            rows_deleted += 1

    _save(wb, path)
    return {"deleted": rows_deleted, "path": path}


def get_sheet_preview(max_rows: int = 50) -> dict:
    """Return headers and first N data rows as JSON-friendly dicts."""
    wb, path = _load_workbook()
    ws = wb["Sheet1"]

    headers = []
    for cell in ws[1]:
        headers.append(str(cell.value) if cell.value is not None else "")

    data_rows = []
    for row in ws.iter_rows(min_row=2, max_row=min(ws.max_row, max_rows + 1), values_only=True):
        if any(v is not None for v in row):
            serialized = []
            for v in row:
                if isinstance(v, (date, datetime)):
                    serialized.append(v.isoformat())
                else:
                    serialized.append(v)
            data_rows.append(serialized)

    return {"headers": headers, "rows": data_rows, "total_rows": len(data_rows)}
