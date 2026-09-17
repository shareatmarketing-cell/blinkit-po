import assert from "node:assert/strict";
import test from "node:test";
import path from "node:path";
import { readdir, readFile } from "node:fs/promises";

import { extractPoData, parseDate } from "./extractor.js";

test("parses Hyperpure textual dates including Sept. abbreviation and time", () => {
  assert.equal(parseDate("Sept. 8, 2026, 11:45 a.m."), "2026-09-08");
  assert.equal(parseDate("Sep. 09, 2026, 11:59 p.m."), "2026-09-09");
  assert.equal(parseDate("September 10, 2026"), "2026-09-10");
});

test("parses supported numeric date layouts and rejects invalid dates", () => {
  assert.equal(parseDate("2026-09-08"), "2026-09-08");
  assert.equal(parseDate("08/09/2026"), "2026-09-08");
  assert.equal(parseDate("Feb. 31, 2026"), null);
});

test("still extracts dates from the repository sample PO", async () => {
  const pdf = await readFile(new URL("../../../6478810007248.pdf", import.meta.url));
  const rows = await extractPoData(new Uint8Array(pdf));

  assert.equal(rows[0]["PO DATE"], "2026-02-17");
  assert.equal(rows[0]["DELIVERY DATE"], "2026-02-27");
});

const realSamplesDir = process.env.BLINKIT_PO_FIXTURES;
const expectedRealSampleDates = {
  "1829010019295.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-03", items: 3, quantity: 156, amount: 5531.4 },
  "1829110016580.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-01", items: 1, quantity: 60, amount: 2079 },
  "1829610025192.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-03", items: 4, quantity: 240, amount: 9912 },
  "2108410019230.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-03", items: 1, quantity: 108, amount: 6010.2 },
  "2275310017899.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-01", items: 4, quantity: 240, amount: 9912 },
  "2576510013929.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-01", items: 2, quantity: 120, amount: 4998 },
  "2738210010114.pdf": { poDate: "2026-08-24", expiryDate: "2026-08-31", items: 5, quantity: 600, amount: 26040 },
  "27580810002521.pdf": { poDate: "2026-08-24", expiryDate: "2026-08-31", items: 2, quantity: 144, amount: 5409.6 },
  "2873510005101.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-03", items: 4, quantity: 444, amount: 23952.6 },
  "43891710002940.pdf": { poDate: "2026-08-24", expiryDate: "2026-08-31", items: 1, quantity: 60, amount: 2499 },
  "49990110002178.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-03", items: 4, quantity: 360, amount: 12810 },
  "50993910000428.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-03", items: 1, quantity: 120, amount: 2898 },
  "51076210000145.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-03", items: 5, quantity: 540, amount: 23583 },
  "52029610000026.pdf": { poDate: "2026-08-24", expiryDate: "2026-08-30", items: 1, quantity: 120, amount: 2898 },
  "5384110011332.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-03", items: 1, quantity: 180, amount: 4347 },
  "5630310006711.pdf": { poDate: "2026-08-24", expiryDate: "2026-08-29", items: 1, quantity: 420, amount: 17493 },
  "6478810008757.pdf": { poDate: "2026-08-24", expiryDate: "2026-09-03", items: 3, quantity: 144, amount: 6039.6 },
  "6769910007761.pdf": { poDate: "2026-08-24", expiryDate: "2026-08-30", items: 2, quantity: 264, amount: 8307.6 },
};

test("extracts exact PO and expiry dates from the real PDF sample set", {
  skip: realSamplesDir ? false : "Set BLINKIT_PO_FIXTURES to run real-PDF regression tests",
}, async () => {
  const pdfNames = (await readdir(realSamplesDir)).filter((name) => name.toLowerCase().endsWith(".pdf")).sort();
  assert.deepEqual(pdfNames, Object.keys(expectedRealSampleDates).sort());

  for (const name of pdfNames) {
    const rows = await extractPoData(new Uint8Array(await readFile(path.join(realSamplesDir, name))));
    const expected = expectedRealSampleDates[name];
    assert.ok(rows.length > 0, `${name}: expected at least one item row`);
    assert.equal(String(rows[0]["PO NO"]), path.basename(name, ".pdf"), `${name}: PO number`);
    assert.equal(rows[0]["PO DATE"], expected.poDate, `${name}: PO date`);
    assert.equal(rows[0]["DELIVERY DATE"], expected.expiryDate, `${name}: PO expiry date`);
    assert.equal(rows.length, expected.items, `${name}: item count`);
    assert.equal(rows.reduce((sum, row) => sum + row["TOTAL PCS"], 0), expected.quantity, `${name}: total quantity`);
    assert.ok(
      Math.abs(rows.reduce((sum, row) => sum + row["TOTAL BASIC PO VALUE WITH TAX"], 0) - expected.amount) <= 0.011,
      `${name}: total amount`,
    );
    for (const row of rows) {
      assert.equal(row["PO DATE"], expected.poDate, `${name}: repeated PO date`);
      assert.equal(row["DELIVERY DATE"], expected.expiryDate, `${name}: repeated PO expiry date`);
      assert.ok(row["ARTICLE DESCRIPTION"], `${name}: item description`);
      assert.ok(Math.abs(row["TOTAL PCS"] * row["LANDING PRICE"] - row["TOTAL BASIC PO VALUE WITH TAX"]) <= 0.011, `${name}: line math`);
    }
  }
});
