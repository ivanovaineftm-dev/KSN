from __future__ import annotations

from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from app.processor import process_excel

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
PROCESSED_DIR = BASE_DIR / "processed"
TEMPLATES_DIR = BASE_DIR / "app" / "templates"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Excel Processor")
app.mount("/static", StaticFiles(directory=BASE_DIR / "app" / "static"), name="static")


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((TEMPLATES_DIR / "index.html").read_text(encoding="utf-8"))


@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)) -> dict[str, str | list[dict[str, int | str]]]:
    filename = file.filename or "uploaded.xlsx"
    if not filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Можно загружать только .xlsx файлы")

    token = uuid4().hex
    input_path = UPLOAD_DIR / f"{token}_{filename}"
    output_name = f"processed_{filename}"
    output_path = PROCESSED_DIR / f"{token}_{output_name}"

    data = await file.read()
    input_path.write_bytes(data)

    analytics = process_excel(input_path=input_path, output_path=output_path)

    return {
        "download_url": f"/download/{token}/{output_name}",
        "filename": output_name,
        "analytics": analytics,
    }


@app.get("/download/{token}/{filename}")
def download_processed(token: str, filename: str) -> FileResponse:
    output_path = PROCESSED_DIR / f"{token}_{filename}"
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден")

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )
