import { useState, useEffect, useCallback } from "react";
import UploadZone from "./components/UploadZone";
import ConfirmModal from "./components/ConfirmModal";
import PreviewTable from "./components/PreviewTable";
import { uploadPDFs, clearSheet, fetchPreview, downloadExcel } from "./api";

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
  const [showModal, setShowModal] = useState(false);
  const [clearing, setClearing] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [toasts, setToasts] = useState([]);
  const [preview, setPreview] = useState(null);

  const addToast = (message, type = "success") => {
    const id = Date.now();
    setToasts((t) => [...t, { id, message, type }]);
  };
  const removeToast = (id) => setToasts((t) => t.filter((x) => x.id !== id));

  const loadPreview = useCallback(async () => {
    try {
      setPreview(await fetchPreview());
    } catch { /* ignore */ }
  }, []);

  useEffect(() => { loadPreview(); }, [loadPreview]);

  function handleFilesSelected(newFiles) {
    setFiles((prev) => {
      const existing = new Set(prev.map((f) => f.name));
      return [...prev, ...newFiles.filter((f) => !existing.has(f.name))];
    });
  }

  function removeFile(name) {
    setFiles((f) => f.filter((x) => x.name !== name));
    setResults((r) => { const c = { ...r }; delete c[name]; return c; });
  }

  async function handleUpload() {
    if (!files.length || uploading) return;
    setUploading(true);
    try {
      const data = await uploadPDFs(files);
      const map = {};
      for (const r of data.files) map[r.file] = r;
      setResults(map);
      const ok = data.files.filter((r) => r.status === "ok").length;
      const fail = data.files.filter((r) => r.status === "error").length;
      if (ok) addToast(`${ok} PO${ok !== 1 ? "s" : ""} appended to tracker.`, "success");
      if (fail) addToast(`${fail} file${fail !== 1 ? "s" : ""} failed to parse.`, "error");
      if (data.excel?.excel_error) addToast(`Excel: ${data.excel.excel_error}`, "error");
      await loadPreview();
    } catch (e) {
      addToast(`Upload error: ${e.message}`, "error");
    } finally {
      setUploading(false);
    }
  }

  async function handleDownload() {
    if (downloading) return;
    setDownloading(true);
    try {
      await downloadExcel();
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
      const data = await clearSheet();
      addToast(`Sheet cleared — ${data.rows_deleted} row${data.rows_deleted !== 1 ? "s" : ""} removed.`, "success");
      setFiles([]);
      setResults({});
      await loadPreview();
    } catch (e) {
      addToast(`Clear failed: ${e.message}`, "error");
    } finally {
      setClearing(false);
    }
  }

  const hasFiles = files.length > 0;
  const allDone = hasFiles && files.every((f) => results[f.name]);

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
                  <button className="btn-ghost" onClick={() => { setFiles([]); setResults({}); }}>
                    Clear all
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

              <div className="upload-footer">
                <button
                  className={`btn-primary ${uploading ? "loading" : ""}`}
                  onClick={handleUpload}
                  disabled={uploading || allDone}
                >
                  {uploading ? (
                    <><span className="spinner" /> Processing…</>
                  ) : allDone ? (
                    <>Done ✓</>
                  ) : (
                    <>
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                      Extract &amp; Append to Tracker
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
