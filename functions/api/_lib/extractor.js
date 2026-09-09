// Blinkit / Zomato Hyperpure PO PDF extractor.
//
// Ported from the original Python (pdfplumber-based) extractor. pdfjs-dist
// gives us positioned text runs (x, y, str) instead of pdfplumber's
// line-merged table cells, so metadata fields are recovered by grouping runs
// into visual lines/columns by coordinate instead of regexing table cells,
// and the item table's columns are located dynamically from the header run
// positions (mirrors the original's header-keyword column lookup) since
// column x-offsets shift depending on whether the PO has split CGST+SGST tax
// columns or a single IGST column.

import { getDocument } from "pdfjs-dist/build/pdf.mjs";
import { WorkerMessageHandler } from "pdfjs-dist/build/pdf.worker.mjs";

// Workers have no `Worker` global to run pdf.js's parser thread in, so wire
// its worker message handler directly into this thread -- pdf.js checks
// `globalThis.pdfjsWorker` before trying to spawn a real Worker.
globalThis.pdfjsWorker = { WorkerMessageHandler };

const LEFT_RIGHT_SPLIT_X = 250;
const ROW_GROUP_GAP = 12; // px gap that separates the header block from row 1
const COL_MATCH_TOLERANCE = 12; // px tolerance when assigning a text run to a column band

function parseDate(raw) {
  if (!raw) return null;
  let s = raw.trim();
  s = s.replace(/([A-Za-z]{3})\./g, "$1"); // "Feb." -> "Feb"
  const parts = s.split(",").map((p) => p.trim());
  if (parts.length >= 2) s = `${parts[0]}, ${parts[1]}`;

  const months = {
    jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2, apr: 3, april: 3,
    may: 4, jun: 5, june: 5, jul: 6, july: 6, aug: 7, august: 7, sep: 8, september: 8,
    oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11,
  };
  const m = s.match(/^([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})$/);
  if (!m) return null;
  const month = months[m[1].toLowerCase()];
  if (month === undefined) return null;
  const day = parseInt(m[2], 10);
  const year = parseInt(m[3], 10);
  const d = new Date(Date.UTC(year, month, day));
  if (Number.isNaN(d.getTime())) return null;
  return d.toISOString().slice(0, 10); // YYYY-MM-DD
}

function clean(val) {
  if (val === null || val === undefined) return "";
  let s = String(val);
  s = s.replace(/([A-Za-z(])\n([a-z])/g, "$1$2");
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

function num(val, cast = "float") {
  if (val === null || val === undefined) return null;
  const s = String(val).replace(/[\s,]/g, "");
  if (s === "") return null;
  const n = cast === "int" ? parseInt(s, 10) : parseFloat(s);
  return Number.isNaN(n) ? null : n;
}

function headerKey(str) {
  if (!str) return "";
  return str.replace(/[^A-Za-z0-9%]/g, "").toLowerCase();
}

function findCol(columns, ...substrings) {
  for (const col of columns) {
    if (substrings.every((sub) => col.key.includes(sub))) return col;
  }
  return null;
}

/** Group text runs into visual lines by y-coordinate. */
function groupLines(items) {
  const byY = new Map();
  for (const it of items) {
    if (!it.str || !it.str.trim()) continue;
    const y = Math.round(it.y);
    let bucket = null;
    for (const key of byY.keys()) {
      if (Math.abs(key - y) <= 2) { bucket = key; break; }
    }
    if (bucket === null) bucket = y;
    if (!byY.has(bucket)) byY.set(bucket, []);
    byY.get(bucket).push(it);
  }
  const lines = [...byY.entries()].map(([y, runs]) => ({
    y,
    runs: runs.sort((a, b) => a.x - b.x),
  }));
  lines.sort((a, b) => b.y - a.y); // top of page first
  return lines;
}

function lineText(runs) {
  return runs.map((r) => r.str).join(" ").replace(/\s+/g, " ").trim();
}

/** Parse the "Label : value" metadata fields out of one page's text runs. */
function extractMetadata(items) {
  const lines = groupLines(items);
  const meta = {
    siteCode: "",
    vendorName: "",
    vendorCode: null,
    poNumber: null,
    poDate: null,
    deliveryDate: null,
  };

  for (const { runs } of lines) {
    const left = runs.filter((r) => r.x < LEFT_RIGHT_SPLIT_X);
    const right = runs.filter((r) => r.x >= LEFT_RIGHT_SPLIT_X);

    for (const band of [left, right]) {
      if (!band.length) continue;
      const text = lineText(band);

      if (/ZHPL/i.test(text)) {
        const m = text.match(/ZHPL\s*[-–]\s*(.+)/i);
        if (m) meta.siteCode = clean(m[1]);
        continue;
      }

      const colonIdx = text.indexOf(":");
      if (colonIdx === -1) continue;
      const label = text.slice(0, colonIdx).trim();
      const value = text.slice(colonIdx + 1).trim();
      if (!label || !value) continue;
      const labelLc = label.toLowerCase();

      if (labelLc === "vendor") {
        meta.vendorName = clean(value);
      } else if (labelLc.includes("vendor no")) {
        const n = num(value, "int");
        meta.vendorCode = n === null ? value : n;
      } else if (labelLc.includes("p.o. number") || labelLc.includes("po number")) {
        const n = num(value, "int");
        meta.poNumber = n === null ? value : n;
      } else if (labelLc === "date") {
        meta.poDate = parseDate(value);
      } else if (labelLc.includes("po delivery")) {
        meta.deliveryDate = parseDate(value);
      }
    }
  }

  return meta;
}

/** Locate the item table's header row and derive column bands from it. */
function findColumns(items) {
  const hashRun = items.find((it) => it.str.trim() === "#");
  if (!hashRun) return null;
  const headerTopY = Math.round(hashRun.y);
  const col0X = hashRun.x;

  // Row 1's "#" column value is a lone digit run at the same x as the header,
  // sitting well below the (multi-line) header block.
  const candidates = items
    .filter((it) => Math.abs(it.x - col0X) <= COL_MATCH_TOLERANCE && /^\d+$/.test(it.str.trim()) && it.y < headerTopY - 2)
    .sort((a, b) => b.y - a.y);
  const row1Y = candidates.length ? Math.round(candidates[0].y) : headerTopY - 100;

  const headerRuns = items.filter(
    (it) => it.str.trim() && it.y <= headerTopY + 2 && it.y > row1Y + ROW_GROUP_GAP
  );

  // Cluster header runs into column bands by x-proximity.
  const bands = [];
  for (const run of headerRuns) {
    let band = bands.find((b) => Math.abs(b.x - run.x) <= COL_MATCH_TOLERANCE);
    if (!band) {
      band = { x: run.x, runs: [] };
      bands.push(band);
    }
    band.runs.push(run);
  }
  bands.sort((a, b) => a.x - b.x);

  const columns = bands.map((b) => {
    const text = b.runs
      .sort((a, c) => c.y - a.y)
      .map((r) => r.str)
      .join(" ");
    return { x: b.x, key: headerKey(text) };
  });

  return { columns, headerTopY, row1Y };
}

function assignColumn(columns, x) {
  let best = null;
  let bestDist = Infinity;
  for (const col of columns) {
    const dist = Math.abs(col.x - x);
    if (dist < bestDist) { bestDist = dist; best = col; }
  }
  return bestDist <= COL_MATCH_TOLERANCE * 2 ? best : null;
}

// Markers that close the item table -- the totals/summary block, terms &
// conditions, or the signature block. Everything from the topmost of these
// markers downward is footer, not table content.
const FOOTER_MARKERS = [
  /^Total Quantity/i,
  /^Total Items/i,
  /^Total weight/i,
  /^Total Amount/i,
  /^Terms\s*&\s*Conditions/i,
  /^Other Conditions/i,
  /^Prepared By/i,
  /^Subject to the Jurisdiction/i,
];

function findFooterCutoffY(items) {
  let cutoff = -Infinity;
  for (const it of items) {
    const text = it.str.trim();
    if (FOOTER_MARKERS.some((re) => re.test(text))) {
      cutoff = Math.max(cutoff, it.y);
    }
  }
  return cutoff;
}

/** Extract item rows from one page given the column bands (derived from page 1's header). */
function extractItemRows(items, columns, col0X, belowY) {
  const footerY = findFooterCutoffY(items);
  const tableItems = footerY === -Infinity ? items : items.filter((it) => it.y > footerY);

  const rowAnchors = tableItems
    .filter((it) => Math.abs(it.x - col0X) <= COL_MATCH_TOLERANCE && /^\d+$/.test(it.str.trim()) && it.y < belowY)
    .sort((a, b) => b.y - a.y);

  if (!rowAnchors.length) return [];

  const rows = [];
  for (let i = 0; i < rowAnchors.length; i++) {
    const topY = rowAnchors[i].y;
    const bottomY = i + 1 < rowAnchors.length ? rowAnchors[i + 1].y : footerY;
    const rowItems = tableItems.filter((it) => it.y <= topY + 1 && it.y > bottomY + 2);

    const byCol = new Map();
    for (const it of rowItems) {
      if (!it.str.trim()) continue;
      const col = assignColumn(columns, it.x);
      if (!col) continue;
      if (!byCol.has(col.key)) byCol.set(col.key, []);
      byCol.get(col.key).push(it);
    }
    const cellText = (key) => {
      const runs = byCol.get(key);
      if (!runs) return null;
      return runs.sort((a, b) => b.y - a.y).map((r) => r.str).join("\n");
    };
    rows.push({ cellText });
  }
  return rows;
}

function makeRow(chain, siteCode, vendorCode, vendorName, poNumber, poDate, deliveryDate, description, qty, landingPrice, totalAmt, extras) {
  return {
    "CHAINS": chain,
    "SITE CODE": siteCode,
    "VENDOR CODE": vendorCode,
    "VENDOR NAME": vendorName,
    "PO NO": poNumber,
    "PO DATE": poDate,
    "DELIVERY DATE": deliveryDate,
    "ARTICLE DESCRIPTION": description,
    "TOTAL PCS": qty,
    "LANDING PRICE": landingPrice,
    "TOTAL BASIC PO VALUE WITH TAX": totalAmt,
    ...(extras || {}),
  };
}

export async function extractPoData(pdfBytes) {
  const pdf = await getDocument({
    data: pdfBytes,
    isEvalSupported: false,
    useSystemFonts: true,
  }).promise;

  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const items = content.items
      .filter((it) => "str" in it)
      .map((it) => ({
        str: it.str,
        x: it.transform[4],
        y: it.transform[5],
        width: it.width,
      }));
    pages.push(items);
  }

  if (!pages.length || !pages[0].length) {
    throw new Error("No text found in PDF");
  }

  const chain = "Zomato Hyperpure Private Limited";
  const meta = extractMetadata(pages[0]);

  const colInfo = findColumns(pages[0]);
  if (!colInfo) {
    return [makeRow(chain, meta.siteCode, meta.vendorCode, meta.vendorName, meta.poNumber, meta.poDate, meta.deliveryDate, "", null, null, null)];
  }

  const { columns, headerTopY, row1Y } = colInfo;
  const hashRun = pages[0].find((it) => it.str.trim() === "#");
  const col0X = hashRun.x;

  const descCol = findCol(columns, "description");
  const qtyCol = findCol(columns, "qty");
  const landingCol = findCol(columns, "landing");
  const totalCol = findCol(columns, "totalamt");
  const hsnCol = findCol(columns, "hsncode");
  const itemCodeCol = findCol(columns, "itemcode");
  const basicCostCol = findCol(columns, "basiccostprice");
  const mrpCol = columns.find((c) => c.key === "mrp");
  const marginCol = findCol(columns, "margin");
  const taxAmtCol = findCol(columns, "taxamt");

  const allRows = [];
  // Page 1: item rows live strictly below the header block.
  allRows.push(...extractItemRows(pages[0], columns, col0X, row1Y + ROW_GROUP_GAP));
  // Continuation pages: no header of their own, reuse page-1 column bands.
  for (let p = 1; p < pages.length; p++) {
    allRows.push(...extractItemRows(pages[p], columns, col0X, Infinity));
  }

  const items = [];
  for (const row of allRows) {
    const description = clean(row.cellText(descCol?.key));
    const qty = num(row.cellText(qtyCol?.key), "int");
    const landing = num(row.cellText(landingCol?.key));
    const total = num(row.cellText(totalCol?.key));

    const extras = {
      "HSN CODE": clean(row.cellText(hsnCol?.key)),
      "ITEM CODE": clean(row.cellText(itemCodeCol?.key)),
      "BASIC COST PRICE": num(row.cellText(basicCostCol?.key)),
      "MRP": num(row.cellText(mrpCol?.key)),
      "MARGIN %": num(row.cellText(marginCol?.key)),
      "TAX AMT": num(row.cellText(taxAmtCol?.key)),
    };

    items.push(makeRow(chain, meta.siteCode, meta.vendorCode, meta.vendorName, meta.poNumber, meta.poDate, meta.deliveryDate, description, qty, landing, total, extras));
  }

  if (!items.length) {
    items.push(makeRow(chain, meta.siteCode, meta.vendorCode, meta.vendorName, meta.poNumber, meta.poDate, meta.deliveryDate, "", null, null, null));
  }

  return items;
}
