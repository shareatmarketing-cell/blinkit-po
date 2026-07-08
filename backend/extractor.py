import re
from datetime import datetime, date

import pdfplumber


# ── Date helpers ───────────────────────────────────────────────────────────────

def _parse_date(raw: str):
    """Parse dates like 'Feb. 17, 2026, 11:45 a.m.' → date object."""
    if not raw:
        return None
    s = raw.strip()
    # Remove trailing time portion (keep only "Month DD, YYYY")
    # Normalise: strip trailing periods from month abbreviations
    s = re.sub(r"([A-Za-z]{3})\.", r"\1", s)   # "Feb." → "Feb"
    # Keep only up to the year (first 3 comma-parts)
    parts = [p.strip() for p in s.split(",")]
    # parts[0] = "Feb 17", parts[1] = "2026", rest = time
    if len(parts) >= 2:
        s = f"{parts[0]}, {parts[1]}"
    for fmt in ["%b %d, %Y", "%B %d, %Y"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _clean(val) -> str:
    """
    Collapse newlines in a cell value.
    If a newline falls mid-word (no space before/after it), join without space;
    otherwise join with a space.
    """
    if val is None:
        return ""
    s = str(val)
    # Join mid-word line-breaks without space (only when continuation is lowercase)
    s = re.sub(r"([A-Za-z(])\n([a-z])", r"\1\2", s)
    # Replace remaining newlines with spaces and collapse whitespace
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _num(val, cast=float):
    """Clean + cast a numeric cell value."""
    if val is None:
        return None
    s = re.sub(r"[\s,]", "", str(val))
    try:
        return cast(s)
    except ValueError:
        return None


def _header_key(cell) -> str:
    """Collapse a (possibly line-wrapped) header cell into a bare lowercase key
    for matching, e.g. 'Land\\ning\\nRate' -> 'landingrate'."""
    if not cell:
        return ""
    return re.sub(r"[^A-Za-z0-9%]", "", str(cell)).lower()


def _find_col(col_map: dict, *substrings) -> "int | None":
    """Return the column index whose header key contains ALL given substrings."""
    for key, idx in col_map.items():
        if all(sub in key for sub in substrings):
            return idx
    return None


def _cell(row, idx, caster=_clean):
    """Safely read row[idx] (idx may be None or out of range)."""
    if idx is None or idx >= len(row) or row[idx] is None:
        return None if caster is not _clean else ""
    return caster(row[idx])


# ── Main extractor ─────────────────────────────────────────────────────────────

def extract_po_data(pdf_path: str) -> list[dict]:
    """
    Parse a Blinkit/Zomato Hyperpure PO PDF.
    Returns a list of dicts (one per line item) ready for the Excel tracker.

    Column positions in the item table shift depending on whether the PO uses
    split intra-state tax columns (CGST% + SGST%) or a single inter-state
    column (IGST%) -- so columns are located dynamically from the header row
    instead of hardcoded indices.
    """
    with pdfplumber.open(pdf_path) as pdf:
        tables = []
        for page in pdf.pages:
            tbls = page.extract_tables()
            tables.extend(tbls)

    if not tables:
        raise ValueError("No tables found in PDF")

    # The PO is typically one large table on page 1
    tbl = tables[0]

    # ── Extract header metadata ────────────────────────────────────────────────
    chain = "Zomato Hyperpure Private Limited"
    site_code = ""
    vendor_name = ""
    vendor_code = None
    po_number = None
    po_date = None
    delivery_date = None

    for row in tbl:
        cell0 = _clean(row[0]) if row[0] else ""

        # The right-hand metadata block can land on any column index depending
        # on how many tax columns the item table has -- find it by content,
        # not by a fixed position.
        meta_cell = ""
        for c in row[1:]:
            if c:
                meta_cell = _clean(c)
                break

        # Row with company info → grab site/warehouse
        if "ZHPL" in cell0:
            m = re.search(r"ZHPL\s*[-–]\s*(.+?)(?:\n|Contact|$)", str(row[0]), re.IGNORECASE)
            if m:
                site_code = m.group(1).strip()

        # Row with vendor name and PO number
        if "Vendor :" in cell0 or "Vendor:" in cell0:
            vm = re.search(r"Vendor\s*:(.+?)(?:\nPAN|\nRegistered|$)", str(row[0]), re.DOTALL)
            if vm:
                vendor_name = _clean(vm.group(1))
            if "P.O. Number" in meta_cell:
                pm = re.search(r"P\.O\. Number\s*:(\S+)", meta_cell)
                if pm:
                    try:
                        po_number = int(pm.group(1))
                    except ValueError:
                        po_number = pm.group(1)
                dm = re.search(r"Date\s*:(.+?)(?:PO Type|Vendor No|$)", meta_cell)
                if dm:
                    po_date = _parse_date(dm.group(1).strip())
                vm2 = re.search(r"Vendor No\.\s*:(\d+)", meta_cell)
                if vm2:
                    try:
                        vendor_code = int(vm2.group(1))
                    except ValueError:
                        vendor_code = vm2.group(1)

        # Row with PO expiry date
        if "PO expiry date" in meta_cell:
            em = re.search(r"PO expiry date\s*:(.+?)(?:PO delivery|GST|$)", meta_cell)
            if em:
                delivery_date = _parse_date(em.group(1).strip())

    # ── Find the item-table header row and locate columns dynamically ──────────
    # The column-header row contains "#" in position 0
    header_idx = None
    for i, row in enumerate(tbl):
        if row[0] and str(row[0]).strip() == "#":
            header_idx = i
            break

    if header_idx is None:
        # No items table found – return header-only row
        return [_make_row(chain, site_code, vendor_code, vendor_name,
                          po_number, po_date, delivery_date, "", None, None, None)]

    col_map = {}
    for i, c in enumerate(tbl[header_idx]):
        key = _header_key(c)
        if key and key not in col_map:
            col_map[key] = i

    desc_idx      = _find_col(col_map, "description")
    qty_idx       = _find_col(col_map, "qty")
    landing_idx   = _find_col(col_map, "landing")
    total_idx     = _find_col(col_map, "totalamt")
    hsn_idx       = _find_col(col_map, "hsncode")
    item_code_idx = _find_col(col_map, "itemcode")
    basic_cost_idx = _find_col(col_map, "basiccostprice")
    mrp_idx       = col_map.get("mrp")
    margin_idx    = _find_col(col_map, "margin")
    tax_amt_idx   = _find_col(col_map, "taxamt")

    # When the item table has many rows, Blinkit/Hyperpure PDFs continue it
    # onto the next page as a separate table with no header row of its own
    # (same column count, numbering picks up where page 1 left off). Splice
    # those continuation rows in so multi-page POs don't silently lose items.
    col_count = len(tbl[header_idx])
    item_rows = list(tbl[header_idx + 1:])
    for extra_tbl in tables[1:]:
        if extra_tbl and len(extra_tbl[0]) == col_count:
            item_rows.extend(extra_tbl)

    items = []
    for row in item_rows:
        raw_num = str(row[0]).strip() if row[0] else ""
        if not raw_num.isdigit():
            break   # reached totals row

        description = _cell(row, desc_idx, _clean)
        qty          = _cell(row, qty_idx, lambda v: _num(v, int))
        landing      = _cell(row, landing_idx, _num)
        total        = _cell(row, total_idx, _num)

        extras = {
            "HSN CODE": _cell(row, hsn_idx, _clean),
            "ITEM CODE": _cell(row, item_code_idx, _clean),
            "BASIC COST PRICE": _cell(row, basic_cost_idx, _num),
            "MRP": _cell(row, mrp_idx, _num),
            "MARGIN %": _cell(row, margin_idx, _num),
            "TAX AMT": _cell(row, tax_amt_idx, _num),
        }

        items.append(_make_row(chain, site_code, vendor_code, vendor_name,
                               po_number, po_date, delivery_date,
                               description, qty, landing, total, extras))

    if not items:
        items.append(_make_row(chain, site_code, vendor_code, vendor_name,
                               po_number, po_date, delivery_date, "", None, None, None))

    return items


def _make_row(chain, site_code, vendor_code, vendor_name,
              po_number, po_date, delivery_date,
              description, qty, landing_price, total_amt, extras: dict | None = None) -> dict:
    row = {
        "CHAINS": chain,
        "SITE CODE": site_code,
        "VENDOR CODE": vendor_code,
        "VENDOR NAME": vendor_name,
        "PO NO": po_number,
        "PO DATE": po_date,
        "DELIVERY DATE": delivery_date,
        "ARTICLE DESCRIPTION": description,
        "TOTAL PCS": qty,
        "LANDING PRICE": landing_price,
        "TOTAL BASIC PO VALUE WITH TAX": total_amt,
    }
    if extras:
        row.update(extras)
    return row
