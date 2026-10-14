// Excel tracker read/write, backed by Cloudflare R2.
// Supports session isolation, automatic template fallback, and clean temporary storage.

import * as XLSX from "xlsx";
import { getTemplateWorkbook, getTemplateBytes } from "./template.js";

export const DEFAULT_OBJECT_KEY = "Blinkit-PO-Tracker.xlsx";
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

export function resolveObjectKey(sessionId) {
  if (!sessionId) return DEFAULT_OBJECT_KEY;
  const safeSession = String(sessionId).trim().replace(/[^a-zA-Z0-9_-]/g, "");
  return safeSession ? `sessions/${safeSession}/Blinkit-PO-Tracker.xlsx` : DEFAULT_OBJECT_KEY;
}

const memoryStore = new Map();

async function loadWorkbook(bucket, objectKey = DEFAULT_OBJECT_KEY) {
  if (bucket && typeof bucket.get === "function") {
    try {
      const obj = await bucket.get(objectKey);
      if (obj) {
        const buf = await obj.arrayBuffer();
        return XLSX.read(buf, { type: "array", cellDates: true });
      }
    } catch (e) {
      console.warn(`[excel] R2 get failed for ${objectKey}, using fallback:`, e.message || e);
    }
  }
  if (memoryStore.has(objectKey)) {
    const buf = memoryStore.get(objectKey);
    return XLSX.read(buf, { type: "array", cellDates: true });
  }
  // If workbook doesn't exist yet in R2 or memory, initialize from base template
  return getTemplateWorkbook();
}

async function saveWorkbook(bucket, wb, objectKey = DEFAULT_OBJECT_KEY) {
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  memoryStore.set(objectKey, new Uint8Array(buf));
  if (bucket && typeof bucket.put === "function") {
    try {
      await bucket.put(objectKey, buf, {
        httpMetadata: {
          contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
      });
    } catch (e) {
      console.warn(`[excel] R2 put failed for ${objectKey}, saved to memory store:`, e.message || e);
    }
  }
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

function excelDateSerial(value) {
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) throw new Error(`Invalid extracted date: ${value}`);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const utc = Date.UTC(year, month - 1, day);
  const date = new Date(utc);
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) throw new Error(`Invalid extracted date: ${value}`);

  // Write the canonical Excel serial instead of a JavaScript Date object.
  // This behaves consistently in both Node and the Cloudflare Workers runtime.
  return utc / 86400000 + 25569;
}

function writeRowsToSheet(ws, rawHeaders, rows, startRow0 = 1) {
  let nextRow0 = startRow0;
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
        ws[addr] = { t: "n", v: excelDateSerial(value), z: "MM/DD/YYYY" };
      } else if (typeof value === "number") {
        ws[addr] = { t: "n", v: value };
      } else {
        ws[addr] = { t: "s", v: String(value) };
      }
    }
    nextRow0 += 1;
    written += 1;
  }
  return { nextRow0, written };
}

export async function appendRows(bucket, rows, objectKey = DEFAULT_OBJECT_KEY) {
  const wb = await loadWorkbook(bucket, objectKey);
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  let range = sheetDims(ws);
  const rawHeaders = headerMap(ws, range);

  // Find the next empty row (mirrors the "walk back while row is empty" guard)
  let nextRow0 = range.e.r + 1;
  while (nextRow0 > 1) {
    if (rowHasValue(ws, nextRow0 - 1, range)) break;
    nextRow0 -= 1;
  }

  const { nextRow0: updatedRow0, written } = writeRowsToSheet(ws, rawHeaders, rows, nextRow0);

  range.e.r = Math.max(range.e.r, updatedRow0 - 1);
  ws["!ref"] = XLSX.utils.encode_range(range);

  await saveWorkbook(bucket, wb, objectKey);
  return { appended: written };
}

// Replaces all existing data rows with the new batch, starting from row 2
export async function replaceRows(bucket, rows, objectKey = DEFAULT_OBJECT_KEY) {
  const wb = getTemplateWorkbook(); // Start from clean template
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  let range = sheetDims(ws);
  const rawHeaders = headerMap(ws, range);

  const { nextRow0, written } = writeRowsToSheet(ws, rawHeaders, rows, 1);

  range.e.r = Math.max(range.e.r, nextRow0 - 1);
  ws["!ref"] = XLSX.utils.encode_range(range);

  await saveWorkbook(bucket, wb, objectKey);
  return { written, replaced: true };
}

export async function clearSheet(bucket, objectKey = DEFAULT_OBJECT_KEY) {
  memoryStore.delete(objectKey);
  if (bucket && typeof bucket.delete === "function") {
    try {
      await bucket.delete(objectKey);
    } catch (e) {
      console.warn(`[excel] R2 delete failed for ${objectKey}:`, e.message || e);
    }
  }
  return { deleted: true, status: "cleared" };
}

export async function getSheetPreview(bucket, maxRows = 50, objectKey = DEFAULT_OBJECT_KEY) {
  const wb = await loadWorkbook(bucket, objectKey);
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

export async function getWorkbookBytes(bucket, objectKey = DEFAULT_OBJECT_KEY) {
  if (bucket && typeof bucket.get === "function") {
    try {
      const obj = await bucket.get(objectKey);
      if (obj) {
        return obj.arrayBuffer();
      }
    } catch (e) {
      console.warn(`[excel] R2 get failed in getWorkbookBytes for ${objectKey}:`, e.message || e);
    }
  }
  if (memoryStore.has(objectKey)) {
    return memoryStore.get(objectKey).buffer;
  }
  // If not found in R2 or memory, return the clean template bytes
  return getTemplateBytes();
}
