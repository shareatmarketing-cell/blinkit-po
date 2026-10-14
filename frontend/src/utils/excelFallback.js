import * as XLSX from "xlsx";

const TEMPLATE_BASE64 = "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";
const SHEET_NAME = "Sheet1";

const COLUMN_MAP = {
  "CHAINS": "CHAINS",
  "SITE CODE": "SITE CODE ",
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

function getTemplateWorkbook() {
  const binary = atob(TEMPLATE_BASE64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return XLSX.read(bytes, { type: "array", cellDates: true });
}

function sheetDims(ws) {
  const ref = ws["!ref"] || "A1";
  return XLSX.utils.decode_range(ref);
}

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

function excelDateSerial(value) {
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(value);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const utc = Date.UTC(year, month - 1, day);
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
        const serial = excelDateSerial(value);
        if (typeof serial === "number") {
          ws[addr] = { t: "n", v: serial, z: "MM/DD/YYYY" };
        } else {
          ws[addr] = { t: "s", v: String(value) };
        }
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

export function buildClientExcel(rows) {
  const wb = getTemplateWorkbook();
  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) throw new Error("Template sheet not found");

  const range = sheetDims(ws);
  const rawHeaders = headerMap(ws, range);
  const { nextRow0 } = writeRowsToSheet(ws, rawHeaders, rows, 1);

  range.e.r = Math.max(range.e.r, nextRow0 - 1);
  ws["!ref"] = XLSX.utils.encode_range(range);

  return XLSX.write(wb, { type: "array", bookType: "xlsx" });
}

export function downloadClientExcel(rows, filename = "Blinkit PO Tracker.xlsx") {
  const bytes = buildClientExcel(rows);
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
