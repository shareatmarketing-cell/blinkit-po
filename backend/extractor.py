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


# ── Main extractor ─────────────────────────────────────────────────────────────

def extract_po_data(pdf_path: str) -> list[dict]:
    """
    Parse a Blinkit/Zomato Hyperpure PO PDF.
    Returns a list of dicts (one per line item) ready for the Excel tracker.
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
        cell8 = _clean(row[8]) if len(row) > 8 and row[8] else ""

        # Row with company info → grab site/warehouse
        if "ZHPL" in cell0:
            m = re.search(r"ZHPL\s*[-–]\s*(.+?)(?:\n|Contact|$)", str(row[0]), re.IGNORECASE)
            if m:
                site_code = m.group(1).strip()

        # Row with vendor name and PO number
        if "Vendor :" in cell0 or "Vendor:" in cell0:
            vm = re.search(r"Vendor\s*:(.+?)(?:\nPAN|\nRegistered|$)", str(row[0]))
            if vm:
                vendor_name = vm.group(1).strip()
            if "P.O. Number" in cell8:
                pm = re.search(r"P\.O\. Number\s*:(\S+)", cell8)
                if pm:
                    try:
                        po_number = int(pm.group(1))
                    except ValueError:
                        po_number = pm.group(1)
                dm = re.search(r"Date\s*:(.+?)(?:PO Type|Vendor No|$)", cell8)
                if dm:
                    po_date = _parse_date(dm.group(1).strip())
                vm2 = re.search(r"Vendor No\.\s*:(\d+)", cell8)
                if vm2:
                    try:
                        vendor_code = int(vm2.group(1))
                    except ValueError:
                        vendor_code = vm2.group(1)

        # Row with PO expiry date
        if "PO expiry date" in cell8:
            em = re.search(r"PO expiry date\s*:(.+?)(?:PO delivery|GST|$)", cell8)
            if em:
                delivery_date = _parse_date(em.group(1).strip())

    # ── Find item rows ─────────────────────────────────────────────────────────
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

    items = []
    for row in tbl[header_idx + 1:]:
        raw_num = str(row[0]).strip() if row[0] else ""
        if not raw_num.isdigit():
            break   # reached totals row

        description = _clean(row[4]) if len(row) > 4 else ""
        qty         = _num(row[12], int)  if len(row) > 12 else None
        landing     = _num(row[11])        if len(row) > 11 else None
        total       = _num(row[15])        if len(row) > 15 else None

        items.append(_make_row(chain, site_code, vendor_code, vendor_name,
                               po_number, po_date, delivery_date,
                               description, qty, landing, total))

    if not items:
        items.append(_make_row(chain, site_code, vendor_code, vendor_name,
                               po_number, po_date, delivery_date, "", None, None, None))

    return items


def _make_row(chain, site_code, vendor_code, vendor_name,
              po_number, po_date, delivery_date,
              description, qty, landing_price, total_amt) -> dict:
    return {
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
