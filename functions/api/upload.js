import JSZip from "jszip";
import { extractPoData } from "./_lib/extractor.js";
import { appendRows, replaceRows, resolveObjectKey } from "./_lib/excel.js";

async function processPdfBytes(fileName, bytes, results, allRows) {
  try {
    const rows = await extractPoData(bytes);
    allRows.push(...rows);
    results.push({
      file: fileName,
      status: "ok",
      rows_extracted: rows.length,
      po_number: rows[0]?.["PO NO"] ?? null,
      po_date: rows[0]?.["PO DATE"] ?? null,
      po_expiry_date: rows[0]?.["DELIVERY DATE"] ?? null,
      vendor: rows[0]?.["VENDOR NAME"] ?? null,
      raw_rows: rows,
      items: rows.map((r) => ({
        description: r["ARTICLE DESCRIPTION"],
        qty: r["TOTAL PCS"],
        landing_price: r["LANDING PRICE"],
        total: r["TOTAL BASIC PO VALUE WITH TAX"],
      })),
    });
  } catch (e) {
    results.push({ file: fileName, status: "error", message: String(e.message || e) });
  }
}

export async function onRequestPost({ request, env }) {
  const url = new URL(request.url);
  const form = await request.formData();
  const files = form.getAll("files");

  const sessionId =
    url.searchParams.get("sessionId") ||
    request.headers.get("x-session-id") ||
    form.get("sessionId");
  const objectKey = resolveObjectKey(sessionId);

  // Default to replace mode so new upload batches overwrite previous temporary POs
  const mode = url.searchParams.get("mode") || form.get("mode") || "replace";

  const results = [];
  const allRows = [];

  for (const upload of files) {
    if (!(upload instanceof File)) {
      results.push({ file: upload?.name ?? "unknown", status: "error", message: "Invalid file object" });
      continue;
    }

    const lowerName = upload.name.toLowerCase();
    if (lowerName.endsWith(".zip")) {
      try {
        const zipBytes = await upload.arrayBuffer();
        const zip = await JSZip.loadAsync(zipBytes);
        const pdfEntries = Object.entries(zip.files).filter(([name, entry]) =>
          !entry.dir && name.toLowerCase().endsWith(".pdf") && !name.startsWith("__MACOSX/")
        );
        if (pdfEntries.length === 0) {
          results.push({ file: upload.name, status: "error", message: "No PDF files found inside ZIP archive" });
          continue;
        }
        for (const [entryPath, entry] of pdfEntries) {
          const entryName = entryPath.split("/").pop();
          const pdfBytes = await entry.async("uint8array");
          await processPdfBytes(entryName, pdfBytes, results, allRows);
        }
      } catch (e) {
        results.push({ file: upload.name, status: "error", message: `Failed to unpack ZIP: ${e.message || e}` });
      }
    } else if (lowerName.endsWith(".pdf")) {
      const bytes = new Uint8Array(await upload.arrayBuffer());
      await processPdfBytes(upload.name, bytes, results, allRows);
    } else {
      results.push({ file: upload.name, status: "error", message: "Only .pdf or .zip files supported" });
    }
  }

  let excelInfo;
  if (allRows.length) {
    try {
      if (mode === "append") {
        const writeResult = await appendRows(env.EXCEL_BUCKET, allRows, objectKey);
        excelInfo = { rows_appended: writeResult.appended, mode: "append" };
      } else {
        const writeResult = await replaceRows(env.EXCEL_BUCKET, allRows, objectKey);
        excelInfo = { rows_written: writeResult.written, mode: "replace" };
      }
    } catch (e) {
      excelInfo = { excel_error: String(e.message || e) };
    }
  } else {
    excelInfo = { rows_written: 0 };
  }

  return Response.json({ files: results, excel: excelInfo });
}
