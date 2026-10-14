import assert from "node:assert/strict";
import test from "node:test";
import path from "node:path";
import { readdir, readFile } from "node:fs/promises";
import * as XLSX from "xlsx";

import { extractPoData } from "./extractor.js";
import { appendRows, getSheetPreview } from "./excel.js";

function memoryBucket(initialBytes) {
  let stored = new Uint8Array(initialBytes);
  return {
    bytes() {
      return stored;
    },
    async get() {
      return {
        async arrayBuffer() {
          return stored.buffer.slice(stored.byteOffset, stored.byteOffset + stored.byteLength);
        },
      };
    },
    async put(_key, bytes) {
      stored = new Uint8Array(bytes);
    },
  };
}

test("persists extracted PO and expiry dates in the downloadable workbook", async () => {
  const template = await readFile(new URL("../../../Blinkit PO Tracker.xlsx", import.meta.url));
  const pdf = await readFile(new URL("../../../6478810007248.pdf", import.meta.url));
  const rows = await extractPoData(new Uint8Array(pdf));
  const bucket = memoryBucket(template);

  await appendRows(bucket, rows);
  const preview = await getSheetPreview(bucket, 10);

  assert.equal(preview.rows.length, 1);
  assert.equal(preview.rows[0][6], 6478810007248);
  assert.equal(preview.rows[0][7], "2026-02-17");
  assert.equal(preview.rows[0][8], "2026-02-27");

  const saved = XLSX.read(bucket.bytes(), { type: "array", cellNF: true });
  assert.equal(saved.Sheets.Sheet1.H2.z, "MM/DD/YYYY");
  assert.equal(saved.Sheets.Sheet1.I2.z, "MM/DD/YYYY");
});

const realSamplesDir = process.env.BLINKIT_PO_FIXTURES;

test("round-trips dates from every real PDF through the workbook", {
  skip: realSamplesDir ? false : "Set BLINKIT_PO_FIXTURES to run real-PDF regression tests",
}, async () => {
  const template = await readFile(new URL("../../../Blinkit PO Tracker.xlsx", import.meta.url));
  const bucket = memoryBucket(template);
  const expectedByPo = new Map();
  const pdfNames = (await readdir(realSamplesDir)).filter((name) => name.toLowerCase().endsWith(".pdf")).sort();

  for (const name of pdfNames) {
    const pdf = await readFile(path.join(realSamplesDir, name));
    const rows = await extractPoData(new Uint8Array(pdf));
    expectedByPo.set(String(rows[0]["PO NO"]), {
      poDate: rows[0]["PO DATE"],
      expiryDate: rows[0]["DELIVERY DATE"],
      rowCount: rows.length,
    });
    await appendRows(bucket, rows);
  }

  const preview = await getSheetPreview(bucket, 1000);
  assert.equal(preview.rows.length, [...expectedByPo.values()].reduce((sum, po) => sum + po.rowCount, 0));
  for (const row of preview.rows) {
    const expected = expectedByPo.get(String(row[6]));
    assert.ok(expected, `unexpected PO ${row[6]}`);
    assert.equal(row[7], expected.poDate, `PO ${row[6]}: persisted PO date`);
    assert.equal(row[8], expected.expiryDate, `PO ${row[6]}: persisted PO expiry date`);
  }
});

test("replaceRows resets old rows and leaves only new rows", async () => {
  const mapBucket = new Map();
  const bucket = {
    async get(key) {
      const data = mapBucket.get(key);
      if (!data) return null;
      return { arrayBuffer: async () => data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) };
    },
    async put(key, bytes) {
      mapBucket.set(key, new Uint8Array(bytes));
    },
    async delete(key) {
      mapBucket.delete(key);
    },
  };

  const pdf = await readFile(new URL("../../../6478810007248.pdf", import.meta.url));
  const rows = await extractPoData(new Uint8Array(pdf));

  // Initial upload of rows
  await appendRows(bucket, rows, "test-session.xlsx");
  let preview = await getSheetPreview(bucket, 10, "test-session.xlsx");
  assert.equal(preview.rows.length, 1);

  // New batch replaces old batch
  const newRow = [{ ...rows[0], "PO NO": 9999999999999, "ARTICLE DESCRIPTION": "New Item Only" }];
  const { replaceRows, clearSheet } = await import("./excel.js");
  await replaceRows(bucket, newRow, "test-session.xlsx");

  preview = await getSheetPreview(bucket, 10, "test-session.xlsx");
  assert.equal(preview.rows.length, 1);
  assert.equal(preview.rows[0][6], 9999999999999);
  assert.equal(preview.rows[0][10], "New Item Only");

  // Clear sheet deletes object from R2 bucket
  await clearSheet(bucket, "test-session.xlsx");
  assert.equal(mapBucket.has("test-session.xlsx"), false);

  // Previewing empty bucket returns 0 rows without error
  const emptyPreview = await getSheetPreview(bucket, 10, "test-session.xlsx");
  assert.equal(emptyPreview.rows.length, 0);
  assert.ok(emptyPreview.headers.length > 0);
});

