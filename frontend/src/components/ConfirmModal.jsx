export default function ConfirmModal({ onConfirm, onCancel }) {
  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
            <path d="M12 9v4m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" strokeLinecap="round" strokeLinejoin="round"/>
          </svg>
        </div>
        <h3>Clear all data?</h3>
        <p>
          This will permanently delete all rows from <strong>Blinkit PO Tracker.xlsx</strong>.
          The header row will be preserved. This action cannot be undone.
        </p>
        <div className="modal-actions">
          <button className="btn-secondary" onClick={onCancel}>Cancel</button>
          <button className="btn-danger" onClick={onConfirm}>Yes, clear sheet</button>
        </div>
      </div>
    </div>
  );
}
