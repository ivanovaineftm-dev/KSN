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


ALLOWED_EXTENSIONS = (".xlsx", ".xls")


@app.post("/upload")
async def upload_excel(
    main_file: UploadFile = File(...),
    locations_file: UploadFile = File(...),
) -> dict[str, str | list[dict[str, int | str]]]:
    main_filename = main_file.filename or "uploaded.xlsx"
    locations_filename = locations_file.filename or "locations.xlsx"
    if not main_filename.lower().endswith(ALLOWED_EXTENSIONS):
        raise HTTPException(status_code=400, detail="Основной файл должен быть в формате .xlsx или .xls")
    if not locations_filename.lower().endswith(ALLOWED_EXTENSIONS):
        raise HTTPException(status_code=400, detail='Файл "Локации" должен быть в формате .xlsx или .xls')

    token = uuid4().hex
    input_path = UPLOAD_DIR / f"{token}_main_{main_filename}"
    locations_path = UPLOAD_DIR / f"{token}_locations_{locations_filename}"
    output_name = f"processed_{Path(main_filename).stem}.xlsx"
    output_path = PROCESSED_DIR / f"{token}_{output_name}"

    input_data = await main_file.read()
    locations_data = await locations_file.read()
    input_path.write_bytes(input_data)
    locations_path.write_bytes(locations_data)

    try:
        analytics = process_excel(input_path=input_path, locations_path=locations_path, output_path=output_path)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

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
