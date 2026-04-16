from __future__ import annotations

from pathlib import Path
from uuid import uuid4

from typing import Annotated, Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from app.processor import process_excel

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
PROCESSED_DIR = BASE_DIR / "processed"
TEMPLATES_DIR = BASE_DIR / "app" / "templates"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Excel Processor")
app.mount("/static", StaticFiles(directory=BASE_DIR / "app" / "static"), name="static")

ALLOWED_EXTENSIONS = (".xlsx", ".xls")
UPLOAD_SESSIONS: dict[str, dict[str, str | None]] = {}


class ProcessRequest(BaseModel):
    file_id: str


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((TEMPLATES_DIR / "index.html").read_text(encoding="utf-8"))


def _validate_excel_filename(filename: str, error_text: str) -> None:
    if not filename.lower().endswith(ALLOWED_EXTENSIONS):
        raise HTTPException(status_code=400, detail=error_text)


async def _store_upload_files(
    main_file: UploadFile,
    locations_file: UploadFile | None,
    barista_file: UploadFile | None,
) -> tuple[str, str]:
    main_filename = main_file.filename or "uploaded.xlsx"
    locations_filename = (locations_file.filename or "locations.xlsx") if locations_file else None
    barista_filename = (barista_file.filename or "barista.xlsx") if barista_file else None

    _validate_excel_filename(main_filename, "Основной файл должен быть в формате .xlsx или .xls")
    if locations_filename:
        _validate_excel_filename(locations_filename, 'Файл "Локации" должен быть в формате .xlsx или .xls')
    if barista_filename:
        _validate_excel_filename(barista_filename, 'Файл "Бариста" должен быть в формате .xlsx или .xls')

    token = uuid4().hex
    input_path = UPLOAD_DIR / f"{token}_main_{main_filename}"
    locations_path = UPLOAD_DIR / f"{token}_locations_{locations_filename}" if locations_filename else None
    barista_path = UPLOAD_DIR / f"{token}_barista_{barista_filename}" if barista_filename else None
    output_name = f"processed_{Path(main_filename).stem}.xlsx"

    input_path.write_bytes(await main_file.read())
    if locations_file and locations_path:
        locations_path.write_bytes(await locations_file.read())
    if barista_file and barista_path:
        barista_path.write_bytes(await barista_file.read())

    UPLOAD_SESSIONS[token] = {
        "main": str(input_path),
        "locations": str(locations_path) if locations_path else None,
        "barista": str(barista_path) if barista_path else None,
        "output_name": output_name,
    }
    return token, output_name


def _run_processing(token: str) -> dict[str, Any]:
    session = UPLOAD_SESSIONS.get(token)
    if not session:
        raise HTTPException(status_code=404, detail="Загрузка не найдена. Повторите загрузку файлов.")

    output_name = str(session["output_name"])
    output_path = PROCESSED_DIR / f"{token}_{output_name}"
    try:
        metrics = process_excel(
            input_path=Path(str(session["main"])),
            locations_path=Path(str(session["locations"])) if session.get("locations") else None,
            barista_path=Path(str(session["barista"])) if session.get("barista") else None,
            output_path=output_path,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return {
        "file_id": token,
        **metrics,
    }


@app.post("/upload/")
@app.post("/upload")
async def upload_excel(
    main_file: Annotated[UploadFile, File(...)],
    locations_file: Annotated[UploadFile | None, File()] = None,
    barista_file: Annotated[UploadFile | None, File()] = None,
) -> dict[str, str]:
    file_id, _ = await _store_upload_files(main_file, locations_file, barista_file)
    return {"file_id": file_id}


@app.post("/process/")
def process_uploaded(payload: ProcessRequest) -> dict[str, Any]:
    return _run_processing(payload.file_id)


@app.post("/upload-process")
async def upload_and_process(
    main_file: Annotated[UploadFile, File(...)],
    locations_file: Annotated[UploadFile | None, File()] = None,
    barista_file: Annotated[UploadFile | None, File()] = None,
) -> dict[str, Any]:
    token, _ = await _store_upload_files(main_file, locations_file, barista_file)
    return _run_processing(token)


@app.get("/download/{file_id}")
def download_by_file_id(file_id: str) -> FileResponse:
    session = UPLOAD_SESSIONS.get(file_id)
    if not session:
        raise HTTPException(status_code=404, detail="Файл не найден")

    output_name = str(session["output_name"])
    output_path = PROCESSED_DIR / f"{file_id}_{output_name}"
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден")

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_name,
    )


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
