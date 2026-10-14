import { clearSheet, resolveObjectKey, DEFAULT_OBJECT_KEY } from "./_lib/excel.js";

async function handleClear({ request, env }) {
  try {
    const url = new URL(request.url);
    const sessionId = url.searchParams.get("sessionId") || request.headers.get("x-session-id");
    const objectKey = resolveObjectKey(sessionId);

    const result = await clearSheet(env.EXCEL_BUCKET, objectKey);

    // If requested to clear all or no session ID given, ensure default object is cleared as well
    if (!sessionId || url.searchParams.get("all") === "true") {
      await clearSheet(env.EXCEL_BUCKET, DEFAULT_OBJECT_KEY);
    }

    return Response.json({ status: "cleared", key: objectKey });
  } catch (e) {
    return Response.json({ detail: String(e.message || e) }, { status: 500 });
  }
}

export async function onRequestDelete(context) {
  return handleClear(context);
}

export async function onRequestPost(context) {
  return handleClear(context);
}
