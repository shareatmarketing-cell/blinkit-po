import os
import shutil
import tempfile
from typing import List

from fastapi import APIRouter, FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

from .extractor import extract_po_data
from .excel_handler import append_rows, clear_sheet, get_sheet_preview, get_live_path

app = FastAPI(title="Blinkit PO Extractor API")

# Use a router to handle the /api prefix correctly for Vercel and local dev
router = APIRouter(prefix="/api")

@router.get("/health")
def health():
    return {"status": "ok"}


@router.post("/upload")
async def upload_pdfs(files: List[UploadFile] = File(...)):
    """
    Accept one or more PDF files, extract PO data, append to Excel, and
    return a summary of results.
    """
    results = []
    all_rows = []

    for upload in files:
        if not upload.filename.lower().endswith(".pdf"):
            results.append({"file": upload.filename, "status": "error", "message": "Not a PDF file"})
            continue

        # Save upload to a temp file
        suffix = os.path.splitext(upload.filename)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            shutil.copyfileobj(upload.file, tmp)
            tmp_path = tmp.name

        try:
            rows = extract_po_data(tmp_path)
            all_rows.extend(rows)
            results.append({
                "file": upload.filename,
                "status": "ok",
                "rows_extracted": len(rows),
                "po_number": rows[0].get("PO NO") if rows else None,
                "vendor": rows[0].get("VENDOR NAME") if rows else None,
                "items": [
                    {
                        "description": r.get("ARTICLE DESCRIPTION"),
                        "qty": r.get("TOTAL PCS"),
                        "landing_price": r.get("LANDING PRICE"),
                        "total": r.get("TOTAL BASIC PO VALUE WITH TAX"),
                    }
                    for r in rows
                ],
            })
        except Exception as e:
            results.append({"file": upload.filename, "status": "error", "message": str(e)})
        finally:
            os.unlink(tmp_path)

    if all_rows:
        try:
            write_result = append_rows(all_rows)
            excel_info = {"rows_appended": write_result["appended"]}
        except Exception as e:
            excel_info = {"excel_error": str(e)}
    else:
        excel_info = {"rows_appended": 0}

    return JSONResponse({"files": results, "excel": excel_info})


@router.delete("/clear")
def clear():
    """Wipe all data rows from Sheet1 (keeps header)."""
    try:
        result = clear_sheet()
        return {"status": "cleared", "rows_deleted": result["deleted"]}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/download")
def download():
    """Download the Excel tracker file."""
    path = get_live_path()
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Excel file not found")
    return FileResponse(
        path=path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Blinkit PO Tracker.xlsx",
    )


@router.get("/preview")
def preview(max_rows: int = 50):
    """Return a preview of the current tracker data."""
    try:
        return get_sheet_preview(max_rows)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# Include routes and middleware
app.include_router(router)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
