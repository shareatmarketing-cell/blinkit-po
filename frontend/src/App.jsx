import { useState, useEffect, useCallback } from "react";
import UploadZone from "./components/UploadZone";
import ConfirmModal from "./components/ConfirmModal";
import PreviewTable from "./components/PreviewTable";
import { syncRowsToServer, clearSheet, fetchPreview, downloadExcel, cleanupSession } from "./api";
import { unpackFiles } from "./utils/zip";
import { extractPoData } from "./utils/extractor";

function Toast({ message, type, onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 4500);
    return () => clearTimeout(t);
  }, [onClose]);
  return (
    <div className={`toast toast-${type}`}>
      <span>{message}</span>
      <button onClick={onClose}>×</button>
    </div>
  );
}

export default function App() {
  const [files, setFiles] = useState([]);
  const [results, setResults] = useState({});
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(null);
  const [showModal, setShowModal] = useState(false);
  const [clearing, setClearing] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [toasts, setToasts] = useState([]);
  const [preview, setPreview] = useState(null);

  const addToast = (message, type = "success") => {
    const id = Date.now() + Math.random();
    setToasts((t) => [...t, { id, message, type }]);
  };
  const removeToast = (id) => setToasts((t) => t.filter((x) => x.id !== id));

  const loadPreview = useCallback(async () => {
    try {
      setPreview(await fetchPreview());
    } catch { /* ignore */ }
  }, []);

  useEffect(() => { loadPreview(); }, [loadPreview]);

  // Clean up temporary R2 storage when the tab or window is closed
  useEffect(() => {
    const handleUnload = () => {
      cleanupSession();
    };
    window.addEventListener("pagehide", handleUnload);
    window.addEventListener("beforeunload", handleUnload);
    return () => {
      window.removeEventListener("pagehide", handleUnload);
      window.removeEventListener("beforeunload", handleUnload);
    };
  }, []);

  const hasFiles = files.length > 0;
  const allDone = hasFiles && files.every((f) => results[f.name]);

  async function handleFilesSelected(newFiles) {
    try {
      const { files: unpacked, unpackedZipCount, totalPdfsFromZips } = await unpackFiles(newFiles);
      if (unpackedZipCount > 0) {
        addToast(`Extracted ${totalPdfsFromZips} PDF${totalPdfsFromZips !== 1 ? "s" : ""} from ZIP archive.`, "success");
      }

      // If the previous batch was already completed, replace with the newly selected files
      if (allDone) {
        setFiles(unpacked);
        setResults({});
        return;
      }

      setFiles((prev) => {
        const existing = new Set(prev.map((f) => f.name));
        return [...prev, ...unpacked.filter((f) => !existing.has(f.name))];
      });
    } catch (err) {
      addToast(`Error reading files: ${err.message}`, "error");
    }
  }

  function resetBatch() {
    setFiles([]);
    setResults({});
    setUploadProgress(null);
  }

  function removeFile(name) {
    setFiles((f) => f.filter((x) => x.name !== name));
    setResults((r) => { const c = { ...r }; delete c[name]; return c; });
  }

  async function handleUpload() {
    if (!files.length || uploading) return;
    setUploading(true);
    setUploadProgress({
      processed: 0,
      total: files.length,
      currentBatch: 1,
      totalBatches: files.length,
    });

    const newResults = { ...results };
    const allExtractedRows = [];
    let ok = 0;
    let fail = 0;

    try {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        setUploadProgress({
          processed: i + 1,
          total: files.length,
          currentBatch: i + 1,
          totalBatches: files.length,
        });

        try {
          const arrayBuffer = await file.arrayBuffer();
          const bytes = new Uint8Array(arrayBuffer);
          const rows = await extractPoData(bytes);
          allExtractedRows.push(...rows);
          newResults[file.name] = {
            file: file.name,
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
          };
          ok++;
        } catch (e) {
          console.error(`Error parsing ${file.name}:`, e);
          newResults[file.name] = {
            file: file.name,
            status: "error",
            message: e.message || String(e),
          };
          fail++;
        }
        setResults({ ...newResults });
      }

      if (ok > 0) {
        addToast(`${ok} PO${ok !== 1 ? "s" : ""} processed successfully.`, "success");
        // Sync extracted rows to server session (clean, lightweight JSON with zero PDF overhead)
        syncRowsToServer(allExtractedRows, { mode: "replace" }).catch(() => {});

        // Build live preview locally immediately
        const previewHeaders = [
          "CHAINS", "SITE CODE ", "VENDOR CODE", "VENDOR NAME", "PO NO",
          "PO DATE", "DELIVERY DATE", "ARTICLE DESCRIPTION", "TOTAL PCS",
          "LANDING PRICE", "TOTAL BASIC PO VALUE WITH TAX"
        ];
        const previewRows = allExtractedRows.slice(0, 50).map((r) => [
          r["CHAINS"], r["SITE CODE"], r["VENDOR CODE"], r["VENDOR NAME"],
          r["PO NO"], r["PO DATE"], r["DELIVERY DATE"], r["ARTICLE DESCRIPTION"],
          r["TOTAL PCS"], r["LANDING PRICE"], r["TOTAL BASIC PO VALUE WITH TAX"]
        ]);
        setPreview({ headers: previewHeaders, rows: previewRows, total_rows: allExtractedRows.length });
      }

      if (fail > 0) {
        addToast(`${fail} file${fail !== 1 ? "s" : ""} failed to parse.`, "error");
      }
    } catch (e) {
      addToast(`Processing error: ${e.message}`, "error");
    } finally {
      setUploading(false);
      setUploadProgress(null);
    }
  }

  async function handleDownload() {
    if (downloading) return;
    setDownloading(true);
    try {
      // Gather any raw_rows extracted in this session for fallback
      const fallbackRows = [];
      for (const res of Object.values(results)) {
        if (res.raw_rows && Array.isArray(res.raw_rows)) {
          fallbackRows.push(...res.raw_rows);
        }
      }
      await downloadExcel(fallbackRows);
    } catch (e) {
      addToast(`Download failed: ${e.message}`, "error");
    } finally {
      setDownloading(false);
    }
  }

  async function handleClearConfirm() {
    setShowModal(false);
    setClearing(true);
    try {
      await clearSheet();
      addToast("Temporary data cleared from tracker storage.", "success");
      setFiles([]);
      setResults({});
      setUploadProgress(null);
      await loadPreview();
    } catch (e) {
      addToast(`Clear failed: ${e.message}`, "error");
    } finally {
      setClearing(false);
    }
  }

  return (
    <div className="app">
      <header className="header">
        <div className="header-inner">
          <div className="brand">
            <div className="brand-icon">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            </div>
            <div>
              <h1>Blinkit PO Extractor</h1>
              <p>Zomato Hyperpure · Purchase Order Tracker</p>
            </div>
          </div>
          <button
            className={`btn-download ${downloading ? "loading" : ""}`}
            onClick={handleDownload}
            disabled={downloading}
          >
            {downloading ? <span className="spinner" /> : (
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <path d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            )}
            Download Excel
          </button>
          <button
            className={`btn-clear ${clearing ? "loading" : ""}`}
            onClick={() => setShowModal(true)}
            disabled={clearing}
          >
            {clearing ? <span className="spinner" /> : (
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <path d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            )}
            Clear Sheet
          </button>
        </div>
      </header>

      <main className="main">
        <section className="card upload-card">
          <UploadZone onFilesSelected={handleFilesSelected} disabled={uploading} />

          {hasFiles && (
            <>
              <div className="file-list-header">
                <span>{files.length} file{files.length !== 1 ? "s" : ""} selected</span>
                {!uploading && (
                  <button className="btn-ghost" onClick={resetBatch}>
                    {allDone ? "Start new batch" : "Clear all"}
                  </button>
                )}
              </div>

              <div className="file-scroll">
                {files.map((f) => {
                  const r = results[f.name];
                  return (
                    <div key={f.name} className={`file-item ${r ? (r.status === "ok" ? "ok" : "err") : "pending"}`}>
                      <div className="file-info">
                        <svg className="file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"/>
                          <path d="M13 3v6h6" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                        <div>
                          <span className="file-name">{f.name}</span>
                          <span className="file-size">{(f.size / 1024).toFixed(1)} KB</span>
                        </div>
                      </div>
                      <div className="file-status">
                        {!r && (uploading
                          ? <span className="badge badge-processing"><span className="dot-pulse" />Processing</span>
                          : <span className="badge badge-pending">Queued</span>
                        )}
                        {r?.status === "ok" && (
                          <span className="badge badge-ok">
                            {r.rows_extracted} item{r.rows_extracted !== 1 ? "s" : ""} · PO {r.po_number}
                          </span>
                        )}
                        {r?.status === "error" && (
                          <span className="badge badge-err" title={r.message}>Failed</span>
                        )}
                      </div>
                      {!uploading && !r && (
                        <button className="remove-btn" onClick={() => removeFile(f.name)}>×</button>
                      )}
                    </div>
                  );
                })}
              </div>

              {uploadProgress && (
                <div className="upload-progress-container" style={{ margin: "1rem 0" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "0.4rem", fontSize: "0.85rem", color: "#64748b" }}>
                    <span>Processing batch {uploadProgress.currentBatch} of {uploadProgress.totalBatches}…</span>
                    <span>{Math.min(uploadProgress.processed, uploadProgress.total)} / {uploadProgress.total} ({Math.round((Math.min(uploadProgress.processed, uploadProgress.total) / uploadProgress.total) * 100)}%)</span>
                  </div>
                  <div style={{ width: "100%", height: "6px", backgroundColor: "#e2e8f0", borderRadius: "999px", overflow: "hidden" }}>
                    <div
                      style={{
                        width: `${Math.round((Math.min(uploadProgress.processed, uploadProgress.total) / uploadProgress.total) * 100)}%`,
                        height: "100%",
                        backgroundColor: "#10b981",
                        transition: "width 0.3s ease",
                      }}
                    />
                  </div>
                </div>
              )}

              <div className="upload-footer" style={{ display: "flex", gap: "10px" }}>
                {allDone && (
                  <button className="btn-ghost" onClick={resetBatch} style={{ padding: "0.6rem 1.2rem" }}>
                    + Upload New POs
                  </button>
                )}
                <button
                  className={`btn-primary ${uploading ? "loading" : ""}`}
                  onClick={handleUpload}
                  disabled={uploading || allDone}
                >
                  {uploading ? (
                    <><span className="spinner" /> Processing…</>
                  ) : allDone ? (
                    <>Batch Processed ✓</>
                  ) : (
                    <>
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                      Extract &amp; Process POs
                    </>
                  )}
                </button>
              </div>
            </>
          )}
        </section>

        <PreviewTable data={preview} />
      </main>

      {showModal && (
        <ConfirmModal onConfirm={handleClearConfirm} onCancel={() => setShowModal(false)} />
      )}

      <div className="toast-container">
        {toasts.map((t) => (
          <Toast key={t.id} message={t.message} type={t.type} onClose={() => removeToast(t.id)} />
        ))}
      </div>
    </div>
  );
}
