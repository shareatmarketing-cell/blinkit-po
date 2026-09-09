// Excel tracker read/write, backed by Cloudflare R2 instead of Vercel Blob.
// Ported from excel_handler.py. R2 objects are strongly consistent for a
// single bucket, so (unlike the original Vercel Blob code) no retry-on-404
// dance is needed -- a read always sees the most recent write.

import * as XLSX from "xlsx";

const OBJECT_KEY = "Blinkit-PO-Tracker.xlsx";
const SHEET_NAME = "Sheet1";

// Maps our extracted keys -> Excel column headers (exact header text in row 1).
const COLUMN_MAP = {
  "CHAINS": "CHAINS",
  "SITE CODE": "SITE CODE ", // note trailing space in original header
  "VENDOR CODE": "VENDOR CODE",
  "VENDOR NAME": "VENDOR NAME",
  "PO NO": "PO NO",
  "PO DATE": "PO DATE",
  "DELIVERY DATE": "DELIVERY DATE",
  "ARTICLE DESCRIPTION": "ARTICLE DESCRIPTION",
  "TOTAL PCS": "TOTAL PCS",
  "LANDING PRICE": "LANDING PRICE",
  "TOTAL BASIC PO VALUE WITH TAX": "TOTAL BASIC PO VALUE WITH TAX",
};

async function loadWorkbook(bucket) {
  const obj = await bucket.get(OBJECT_KEY);
  if (!obj) throw new Error("Excel file not found in R2");
  const buf = await obj.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  return wb;
}

async function saveWorkbook(bucket, wb) {
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  await bucket.put(OBJECT_KEY, buf, {
    httpMetadata: {
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  });
}

function sheetDims(ws) {
  const ref = ws["!ref"] || "A1";
  return XLSX.utils.decode_range(ref);
}

/** {header_text: 1-based col index} from row 1. */
function headerMap(ws, range) {
  const map = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell && cell.v !== undefined && cell.v !== null) {
      map[String(cell.v).trim()] = c + 1;
    }
  }
  return map;
}

function rowHasValue(ws, rowIdx0, range) {
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: rowIdx0, c })];
    if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") return true;
  }
  return false;
}

export async function appendRows(bucket, rows) {
  const wb = await loadWorkbook(bucket);
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  let range = sheetDims(ws);
  const rawHeaders = headerMap(ws, range);

  // Find the next empty row (mirrors the "walk back while row is empty" guard
  // in the original, since XLSX's declared range can overshoot actual data).
  let nextRow0 = range.e.r + 1; // 0-based row to write into
  while (nextRow0 > 1) {
    if (rowHasValue(ws, nextRow0 - 1, range)) break;
    nextRow0 -= 1;
  }

  let written = 0;
  for (const rowData of rows) {
    for (const [key, header] of Object.entries(COLUMN_MAP)) {
      let colIdx1 = null;
      const headerNorm = header.trim().toUpperCase();
      for (const [h, ci] of Object.entries(rawHeaders)) {
        if (h.toUpperCase() === headerNorm) { colIdx1 = ci; break; }
      }
      if (colIdx1 === null) continue;

      const value = rowData[key];
      if (value === undefined || value === null) continue;
      const addr = XLSX.utils.encode_cell({ r: nextRow0, c: colIdx1 - 1 });

      if (key === "PO DATE" || key === "DELIVERY DATE") {
        ws[addr] = { t: "d", v: new Date(value), z: "MM/DD/YYYY" };
      } else if (typeof value === "number") {
        ws[addr] = { t: "n", v: value };
      } else {
        ws[addr] = { t: "s", v: String(value) };
      }
    }
    nextRow0 += 1;
    written += 1;
  }

  range.e.r = Math.max(range.e.r, nextRow0 - 1);
  ws["!ref"] = XLSX.utils.encode_range(range);

  await saveWorkbook(bucket, wb);
  return { appended: written };
}

export async function clearSheet(bucket) {
  const wb = await loadWorkbook(bucket);
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const range = sheetDims(ws);
  let deleted = 0;
  for (let r = range.e.r; r >= 1; r--) {
    if (rowHasValue(ws, r, range)) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        delete ws[XLSX.utils.encode_cell({ r, c })];
      }
      deleted += 1;
    }
  }
  range.e.r = 0;
  ws["!ref"] = XLSX.utils.encode_range(range);

  await saveWorkbook(bucket, wb);
  return { deleted };
}

export async function getSheetPreview(bucket, maxRows = 50) {
  const wb = await loadWorkbook(bucket);
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const range = sheetDims(ws);
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    headers.push(cell && cell.v !== undefined ? String(cell.v) : "");
  }

  const dataRows = [];
  const lastRow = Math.min(range.e.r, maxRows);
  for (let r = 1; r <= lastRow; r++) {
    if (!rowHasValue(ws, r, range)) continue;
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v === undefined) { row.push(null); continue; }
      row.push(cell.t === "d" ? new Date(cell.v).toISOString().slice(0, 10) : cell.v);
    }
    dataRows.push(row);
  }

  return { headers, rows: dataRows, total_rows: dataRows.length };
}

export async function getWorkbookBytes(bucket) {
  const obj = await bucket.get(OBJECT_KEY);
  if (!obj) return null;
  return obj.arrayBuffer();
}
