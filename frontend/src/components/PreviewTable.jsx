const VISIBLE_COLS = [
  "PO NO", "PO DATE", "DELIVERY DATE", "VENDOR NAME",
  "SITE CODE ", "ARTICLE DESCRIPTION", "TOTAL PCS", "LANDING PRICE",
  "TOTAL BASIC PO VALUE WITH TAX"
];

export default function PreviewTable({ data }) {
  if (!data || !data.rows.length) return null;

  const headers = data.headers;
  const visibleIdx = VISIBLE_COLS.map((c) =>
    headers.findIndex((h) => h.trim() === c.trim())
  ).filter((i) => i !== -1);

  const displayHeaders = visibleIdx.map((i) => headers[i].trim());

  return (
    <div className="preview-section">
      <div className="preview-header">
        <h3>Tracker Preview</h3>
        <span className="row-count">{data.total_rows} row{data.total_rows !== 1 ? "s" : ""}</span>
      </div>
      <div className="table-wrapper">
        <table>
          <thead>
            <tr>
              {displayHeaders.map((h, i) => <th key={i}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {data.rows.map((row, ri) => (
              <tr key={ri}>
                {visibleIdx.map((ci, i) => (
                  <td key={i}>{row[ci] ?? "—"}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
