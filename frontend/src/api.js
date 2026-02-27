const BASE = "/api";

export async function uploadPDFs(files) {
  const form = new FormData();
  for (const f of files) form.append("files", f);
  const res = await fetch(`${BASE}/upload`, { method: "POST", body: form });
  if (!res.ok) throw new Error(`Upload failed: ${res.status}`);
  return res.json();
}

export async function clearSheet() {
  const res = await fetch(`${BASE}/clear`, { method: "DELETE" });
  if (!res.ok) throw new Error(`Clear failed: ${res.status}`);
  return res.json();
}

export async function fetchPreview() {
  const res = await fetch(`${BASE}/preview`);
  if (!res.ok) throw new Error(`Preview failed: ${res.status}`);
  return res.json();
}

export async function downloadExcel() {
  const res = await fetch(`${BASE}/download`);
  if (!res.ok) throw new Error(`Download failed: ${res.status}`);
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Blinkit PO Tracker.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}
