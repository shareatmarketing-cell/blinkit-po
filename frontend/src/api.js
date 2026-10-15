const BASE = "/api";
const SESSION_STORAGE_KEY = "blinkit_po_session_id";

export function getSessionId() {
  if (typeof window === "undefined") return "";
  let id = sessionStorage.getItem(SESSION_STORAGE_KEY);
  if (!id) {
    id = "tab_" + Math.random().toString(36).substring(2, 9) + "_" + Date.now().toString(36);
    sessionStorage.setItem(SESSION_STORAGE_KEY, id);
  }
  return id;
}

export function cleanupSession() {
  const sessionId = getSessionId();
  if (!sessionId) return;
  const url = `${BASE}/clear?sessionId=${encodeURIComponent(sessionId)}`;
  if (typeof navigator !== "undefined" && typeof navigator.sendBeacon === "function") {
    navigator.sendBeacon(url);
  } else {
    fetch(url, { method: "POST", keepalive: true }).catch(() => {});
  }
}

import { downloadClientExcel } from "./utils/excelFallback";

export async function syncRowsToServer(rows, { mode = "replace" } = {}) {
  const sessionId = getSessionId();
  try {
    const res = await fetch(`${BASE}/sync?sessionId=${encodeURIComponent(sessionId)}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Session-ID": sessionId,
      },
      body: JSON.stringify({ rows, mode }),
    });
    return await res.json();
  } catch (e) {
    console.warn("Failed to sync rows to server:", e);
    return { ok: false, error: e.message };
  }
}

export async function uploadPDFs(files, { mode = "replace" } = {}) {
  const sessionId = getSessionId();
  const form = new FormData();
  for (const f of files) form.append("files", f);
  form.append("sessionId", sessionId);
  form.append("mode", mode);

  const res = await fetch(`${BASE}/upload?sessionId=${encodeURIComponent(sessionId)}&mode=${mode}`, {
    method: "POST",
    headers: {
      "X-Session-ID": sessionId,
    },
    body: form,
  });
  if (!res.ok) {
    const errorText = await res.text().catch(() => "");
    throw new Error(`Upload failed: ${res.status} ${res.statusText} ${errorText.slice(0, 100)}`);
  }
  return res.json();
}

/**
 * Uploads files in chunks (default 4 files per batch) to prevent hitting
 * Cloudflare Worker CPU / execution resource limits on large batches (e.g. 79 PDFs).
 */
export async function uploadInBatches(files, { onProgress, batchSize = 4 } = {}) {
  const allResults = [];
  let lastExcelInfo = null;
  const total = files.length;
  let processed = 0;

  for (let i = 0; i < files.length; i += batchSize) {
    const batch = files.slice(i, i + batchSize);
    const mode = i === 0 ? "replace" : "append";
    const res = await uploadPDFs(batch, { mode });

    if (res.files) {
      allResults.push(...res.files);
    }
    lastExcelInfo = res.excel || lastExcelInfo;
    processed += batch.length;

    if (onProgress) {
      onProgress({
        processed,
        total,
        currentBatch: Math.floor(i / batchSize) + 1,
        totalBatches: Math.ceil(files.length / batchSize),
        latestFiles: res.files,
      });
    }
  }

  return { files: allResults, excel: lastExcelInfo };
}

export async function clearSheet() {
  const sessionId = getSessionId();
  const res = await fetch(`${BASE}/clear?sessionId=${encodeURIComponent(sessionId)}`, {
    method: "DELETE",
    headers: {
      "X-Session-ID": sessionId,
    },
  });
  if (!res.ok) throw new Error(`Clear failed: ${res.status}`);
  return res.json();
}

export async function fetchPreview() {
  const sessionId = getSessionId();
  const res = await fetch(`${BASE}/preview?sessionId=${encodeURIComponent(sessionId)}`, {
    headers: {
      "X-Session-ID": sessionId,
    },
  });
  if (!res.ok) throw new Error(`Preview failed: ${res.status}`);
  return res.json();
}

export async function downloadExcel(fallbackRows = []) {
  const sessionId = getSessionId();
  try {
    const res = await fetch(`${BASE}/download?sessionId=${encodeURIComponent(sessionId)}`, {
      headers: {
        "X-Session-ID": sessionId,
      },
    });
    if (res.ok) {
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Blinkit PO Tracker.xlsx";
      a.click();
      URL.revokeObjectURL(url);
      return;
    }
  } catch (e) {
    console.warn("Server download failed, attempting client fallback:", e);
  }

  // Fallback: If server returned error or R2 is disabled, generate workbook directly
  if (fallbackRows && fallbackRows.length > 0) {
    downloadClientExcel(fallbackRows);
    return;
  }
  throw new Error("Download failed and no local rows available to export.");
}
