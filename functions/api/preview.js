import { getSheetPreview } from "./_lib/excel.js";

export async function onRequestGet({ request, env }) {
  const url = new URL(request.url);
  const maxRows = parseInt(url.searchParams.get("max_rows") || "50", 10);
  try {
    const preview = await getSheetPreview(env.EXCEL_BUCKET, maxRows);
    return Response.json(preview);
  } catch (e) {
    return Response.json({ detail: String(e.message || e) }, { status: 500 });
  }
}
