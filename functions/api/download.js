import { getWorkbookBytes } from "./_lib/excel.js";

export async function onRequestGet({ env }) {
  const bytes = await getWorkbookBytes(env.EXCEL_BUCKET);
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
