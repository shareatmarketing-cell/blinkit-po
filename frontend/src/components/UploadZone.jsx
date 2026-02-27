import { useState, useRef } from "react";

export default function UploadZone({ onFilesSelected, disabled }) {
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef(null);

  function handleDrop(e) {
    e.preventDefault();
    setDragging(false);
    if (disabled) return;
    const files = Array.from(e.dataTransfer.files).filter((f) =>
      f.name.toLowerCase().endsWith(".pdf")
    );
    if (files.length) onFilesSelected(files);
  }

  function handleChange(e) {
    const files = Array.from(e.target.files);
    if (files.length) onFilesSelected(files);
    e.target.value = "";
  }

  return (
    <div
      className={`upload-zone ${dragging ? "dragging" : ""} ${disabled ? "disabled" : ""}`}
      onDragOver={(e) => { e.preventDefault(); if (!disabled) setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      onClick={() => !disabled && inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".pdf"
        multiple
        style={{ display: "none" }}
        onChange={handleChange}
      />
      <div className="upload-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
          <path d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" strokeLinecap="round" strokeLinejoin="round"/>
        </svg>
      </div>
      <p className="upload-title">Drop PO PDFs here</p>
      <p className="upload-sub">or click to browse · multiple files supported</p>
    </div>
  );
}
