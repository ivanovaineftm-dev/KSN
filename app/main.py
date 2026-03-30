from __future__ import annotations

from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from app.processor import ProcessingError, process_excel

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
PROCESSED_DIR = BASE_DIR / "processed"
TEMPLATES_DIR = BASE_DIR / "app" / "templates"
SUPPORTED_EXTENSIONS = (".xlsx", ".xls")

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Excel Processor")
app.mount("/static", StaticFiles(directory=BASE_DIR / "app" / "static"), name="static")


def _ensure_supported_excel(upload: UploadFile, field_name: str) -> str:
    filename = upload.filename or ""
    if not filename:
        raise HTTPException(status_code=400, detail=f"Не задано имя файла для поля '{field_name}'")
    if not filename.lower().endswith(SUPPORTED_EXTENSIONS):
        raise HTTPException(status_code=400, detail="Поддерживаются только файлы .xlsx и .xls")
    return filename


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((TEMPLATES_DIR / "index.html").read_text(encoding="utf-8"))


@app.post("/upload")
async def upload_excel(
    main_file: UploadFile = File(...),
    locations_file: UploadFile | None = File(None),
) -> dict[str, str | list[dict[str, int | str]]]:
    main_filename = _ensure_supported_excel(main_file, "main_file")
    locations_filename: str | None = None

    if locations_file is not None and locations_file.filename:
        locations_filename = _ensure_supported_excel(locations_file, "locations_file")

    token = uuid4().hex
    input_path = UPLOAD_DIR / f"{token}_main_{main_filename}"

    main_data = await main_file.read()
    input_path.write_bytes(main_data)

    locations_path: Path | None = None
    if locations_file is not None and locations_filename:
        locations_path = UPLOAD_DIR / f"{token}_locations_{locations_filename}"
        locations_data = await locations_file.read()
        locations_path.write_bytes(locations_data)

    output_name = f"processed_{Path(main_filename).stem}.xlsx"
    output_path = PROCESSED_DIR / f"{token}_{output_name}"

    try:
        analytics = process_excel(input_path=input_path, output_path=output_path, locations_path=locations_path)
    except ProcessingError as exc:
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
