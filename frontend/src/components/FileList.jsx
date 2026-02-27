export default function FileList({ files, results }) {
  return (
    <div className="file-list">
      {files.map((f, i) => {
        const r = results[f.name];
        return (
          <div key={i} className={`file-item ${r ? (r.status === "ok" ? "ok" : "err") : "pending"}`}>
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
              {!r && <span className="badge badge-pending">Queued</span>}
              {r?.status === "ok" && (
                <span className="badge badge-ok">
                  {r.rows_extracted} item{r.rows_extracted !== 1 ? "s" : ""} · PO {r.po_number}
                </span>
              )}
              {r?.status === "error" && (
                <span className="badge badge-err" title={r.message}>Failed</span>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}
