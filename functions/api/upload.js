import { extractPoData } from "./_lib/extractor.js";
import { appendRows } from "./_lib/excel.js";

export async function onRequestPost({ request, env }) {
  const form = await request.formData();
  const files = form.getAll("files");

  const results = [];
  const allRows = [];

  for (const upload of files) {
    if (!(upload instanceof File) || !upload.name.toLowerCase().endsWith(".pdf")) {
      results.push({ file: upload?.name ?? "unknown", status: "error", message: "Not a PDF file" });
      continue;
    }

    try {
      const bytes = new Uint8Array(await upload.arrayBuffer());
      const rows = await extractPoData(bytes);
      allRows.push(...rows);
      results.push({
        file: upload.name,
        status: "ok",
        rows_extracted: rows.length,
        po_number: rows[0]?.["PO NO"] ?? null,
        vendor: rows[0]?.["VENDOR NAME"] ?? null,
        items: rows.map((r) => ({
          description: r["ARTICLE DESCRIPTION"],
          qty: r["TOTAL PCS"],
          landing_price: r["LANDING PRICE"],
          total: r["TOTAL BASIC PO VALUE WITH TAX"],
        })),
      });
    } catch (e) {
      results.push({ file: upload.name, status: "error", message: String(e.message || e) });
    }
  }

  let excelInfo;
  if (allRows.length) {
    try {
      const writeResult = await appendRows(env.EXCEL_BUCKET, allRows);
      excelInfo = { rows_appended: writeResult.appended };
    } catch (e) {
      excelInfo = { excel_error: String(e.message || e) };
    }
  } else {
    excelInfo = { rows_appended: 0 };
  }

  return Response.json({ files: results, excel: excelInfo });
}
