import { getWorkbookBytes, resolveObjectKey } from "./_lib/excel.js";

export async function onRequestGet({ request, env }) {
  const url = new URL(request.url);
  const sessionId = url.searchParams.get("sessionId") || request.headers.get("x-session-id");
  const objectKey = resolveObjectKey(sessionId);

  const bytes = await getWorkbookBytes(env.EXCEL_BUCKET, objectKey);
  if (!bytes) {
    return Response.json({ detail: "Excel file not found" }, { status: 404 });
  }
  return new Response(bytes, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": 'attachment; filename="Blinkit PO Tracker.xlsx"',
    },
  });
}
