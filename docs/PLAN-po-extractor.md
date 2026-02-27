# PLAN - Blinkit PO Extractor

## Overview
A web-based tool designed to automate the extraction of Purchase Order (PO) details from multiple system-generated PDFs and consolidate them into an existing Excel tracker (`Blinkit PO Tracker.xlsx`). The tool will feature a batch upload interface, progress tracking, and sheet management (appending/clearing).

## Project Type: WEB
- **Primary Agent**: `frontend-specialist` (UI/UX), `backend-specialist` (PDF Extraction Logic)

## Success Criteria
- [ ] Successfully parse system-generated PDFs to extract PO Number, Vendor, Items, Quantities, and Prices.
- [ ] Batch upload functionality for multiple PDFs in one go.
- [ ] Append extracted data to the correct columns in `Blinkit PO Tracker.xlsx`.
- [ ] "Clear Sheet" button with a confirmation dialog to wipe existing data from the tracker.
- [ ] Responsive and premium UI for file management.

## Tech Stack
- **Frontend**: Vite + React + Vanilla CSS (for premium aesthetics).
- **Backend**: Python (FastAPI) for high-performance PDF parsing.
- **Library (PDF)**: `pdfplumber` or `pypdf` (optimized for system-generated text PDFs).
- **Library (Excel)**: `pandas` + `openpyxl` for Excel manipulation.

## File Structure
```text
/
├── backend/
│   ├── main.py            # FastAPI entry point
│   ├── extractor.py       # PDF parsing logic
│   └── excel_handler.py   # Excel read/write/clear logic
├── frontend/
│   ├── src/
│   │   ├── components/    # Upload Zone, Confirmation Modal
│   │   └── App.jsx
│   └── index.css
├── Blinkit PO Tracker.xlsx # Target Excel sheet
└── docs/
    └── PLAN-po-extractor.md
```

## Task Breakdown

### Phase 1: Analysis & Environment Setup
| Task ID | Name | Agent | Skills | Priority | Dependencies | INPUT→OUTPUT→VERIFY |
|---------|------|-------|--------|----------|--------------|----------------------|
| T1.1 | Setup Backend Env | `backend-specialist` | `python-patterns` | P0 | None | Python reqs → FastAPI app → Test endpoint works. |
| T1.2 | Setup Frontend Env | `frontend-specialist` | `clean-code` | P0 | None | Vite React → Home page → Localhost renders. |

### Phase 2: PDF Extraction Logic (Backend)
| Task ID | Name | Agent | Skills | Priority | Dependencies | INPUT→OUTPUT→VERIFY |
|---------|------|-------|--------|----------|--------------|----------------------|
| T2.1 | PDF Parsing Service | `backend-specialist` | `api-patterns` | P1 | T1.1 | PDF sample → JSON data → Verify extracted fields match PDF values. |
| T2.2 | Excel Integration | `backend-specialist` | `database-design` | P1 | T1.1 | JSON data → Append to Excel → Open Excel to check rows. |
| T2.3 | Clear Sheet Logic | `backend-specialist` | `clean-code` | P2 | T2.2 | API call → Empty Sheet1 content → Verify Excel is empty. |

### Phase 3: Dashboard & UX (Frontend)
| Task ID | Name | Agent | Skills | Priority | Dependencies | INPUT→OUTPUT→VERIFY |
|---------|------|-------|--------|----------|--------------|----------------------|
| T3.1 | Core Dashboard UI | `frontend-specialist` | `frontend-design` | P1 | T1.2 | Design Mock → React UI → Drag-and-drop zone visible. |
| T3.2 | Batch Upload Imp. | `frontend-specialist` | `clean-code` | P1 | T3.1, T2.1 | Select files → Upload to API → Show processing status. |
| T3.3 | Sheet Management | `frontend-specialist` | `frontend-design` | P2 | T3.1, T2.3 | Clear button → Modal → Confirm → API Call → Success toast. |

### Phase X: Verification
- [ ] Run `python .agent/scripts/verify_all.py .`
- [ ] Manual check: Upload `6478810007248.pdf` and verify columns in `Blinkit PO Tracker.xlsx`.
- [ ] Manual check: Click "Clear Sheet", confirm, and verify Excel is empty (except headers).
- [ ] Verify no purple/violet hex codes in UI.

## ✅ PHASE X COMPLETE
- Lint: [ ]
- Security: [ ]
- Build: [ ]
- Date: [NOT YET]
