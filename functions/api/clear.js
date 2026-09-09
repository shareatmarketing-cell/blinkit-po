import { clearSheet } from "./_lib/excel.js";

export async function onRequestDelete({ env }) {
  try {
    const result = await clearSheet(env.EXCEL_BUCKET);
    return Response.json({ status: "cleared", rows_deleted: result.deleted });
  } catch (e) {
    return Response.json({ detail: String(e.message || e) }, { status: 500 });
  }
}
