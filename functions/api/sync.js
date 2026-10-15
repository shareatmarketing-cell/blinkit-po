import { appendRows, replaceRows, resolveObjectKey } from "./_lib/excel.js";

export async function onRequestPost({ request, env }) {
  try {
    const url = new URL(request.url);
    const sessionId =
      url.searchParams.get("sessionId") ||
      request.headers.get("x-session-id");
    const objectKey = resolveObjectKey(sessionId);
    const body = await request.json();
    const rows = body.rows || [];
    const mode = body.mode || "replace";

    let result;
    if (rows.length) {
      if (mode === "append") {
        result = await appendRows(env.EXCEL_BUCKET, rows, objectKey);
      } else {
        result = await replaceRows(env.EXCEL_BUCKET, rows, objectKey);
      }
    } else {
      result = { written: 0 };
    }

    return Response.json({ ok: true, result });
  } catch (e) {
    return Response.json({ ok: false, error: String(e.message || e) }, { status: 500 });
  }
}
