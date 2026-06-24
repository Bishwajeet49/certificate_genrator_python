#!/usr/bin/env python3
"""Web interface for the certificate generator."""

from __future__ import annotations

import json
import queue
import shutil
import tempfile
import threading
import uuid
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from generate import (
    CertificateGeneratorError,
    InvalidDataError,
    MissingColumnsError,
    MissingFileError,
    REQUIRED_COLUMNS,
    generate_all_certificates,
    load_excel,
)

APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"


@dataclass
class GenerationJob:
    temp_dir: Path
    zip_path: Path | None = None
    ready: threading.Event = field(default_factory=threading.Event)
    error: str | None = None


jobs: dict[str, GenerationJob] = {}

app = FastAPI(title="Certificate Generator")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/")
async def index() -> FileResponse:
    """Serve the upload page."""
    return FileResponse(STATIC_DIR / "index.html")


def _validate_upload(upload: UploadFile, allowed_suffix: str, field_name: str) -> None:
    if not upload.filename:
        raise HTTPException(status_code=400, detail=f"{field_name} is required.")

    suffix = Path(upload.filename).suffix.lower()
    if suffix != allowed_suffix:
        raise HTTPException(
            status_code=400,
            detail=f"{field_name} must be a {allowed_suffix} file.",
        )


def _validate_certificate_seq_start(value: int) -> int:
    if value < 1:
        raise HTTPException(
            status_code=400,
            detail="Certificate sequence start number must be at least 1.",
        )
    return value


def _build_zip(output_paths: list[Path], zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in output_paths:
            archive.write(path, arcname=path.name)


def _stream_generation_events(
    job_id: str,
    excel_path: Path,
    template_path: Path,
    output_dir: Path,
    certificate_seq_start: int,
    event_queue: queue.Queue,
) -> None:
    job = jobs[job_id]

    try:
        generated_count, output_paths, skipped_count = generate_all_certificates(
            excel_path=excel_path,
            template_path=template_path,
            output_dir=output_dir,
            certificate_number_start=certificate_seq_start,
            on_start=lambda total_rows: event_queue.put(
                {
                    "event": "started",
                    "total_rows": total_rows,
                    "generated": 0,
                    "skipped": 0,
                    "percent": 0,
                }
            ),
            on_generating=lambda info: event_queue.put(
                {"event": "generating", **info}
            ),
            on_generated=lambda info: event_queue.put(
                {"event": "generated", **info}
            ),
            on_skipped=lambda info: event_queue.put(
                {"event": "skipped", **info}
            ),
        )

        if generated_count == 0:
            job.error = "No valid certificates could be generated from the Excel file."
            event_queue.put(
                {
                    "event": "error",
                    "message": job.error,
                }
            )
            return

        event_queue.put(
            {
                "event": "zipping",
                "generated": generated_count,
                "skipped": skipped_count,
                "total_rows": generated_count + skipped_count,
                "percent": 99,
            }
        )

        zip_path = job.temp_dir / "certificates.zip"
        _build_zip(output_paths, zip_path)
        job.zip_path = zip_path

        event_queue.put(
            {
                "event": "complete",
                "job_id": job_id,
                "generated": generated_count,
                "skipped": skipped_count,
                "percent": 100,
            }
        )
    except CertificateGeneratorError as exc:
        job.error = str(exc)
        event_queue.put({"event": "error", "message": str(exc)})
    except Exception as exc:
        job.error = f"Unexpected error: {exc}"
        event_queue.put({"event": "error", "message": job.error})
    finally:
        job.ready.set()
        event_queue.put(None)


def _ndjson_stream(event_queue: queue.Queue, worker: threading.Thread) -> Iterator[str]:
    while True:
        item = event_queue.get()
        if item is None:
            break
        yield json.dumps(item) + "\n"

    worker.join()


@app.post("/generate")
async def generate_certificates(
    excel_file: UploadFile = File(...),
    template_image: UploadFile = File(...),
    certificate_seq_start: int = Form(...),
) -> StreamingResponse:
    """Generate certificates and stream live progress as NDJSON."""
    _validate_upload(excel_file, ".xlsx", "Excel file")
    _validate_upload(template_image, ".png", "Template image")
    certificate_seq_start = _validate_certificate_seq_start(certificate_seq_start)

    temp_dir = Path(tempfile.mkdtemp(prefix="certificates_"))
    excel_path = temp_dir / "participants.xlsx"
    template_path = temp_dir / "template.png"
    output_dir = temp_dir / "output"

    try:
        excel_path.write_bytes(await excel_file.read())
        template_path.write_bytes(await template_image.read())
        load_excel(excel_path)
    except MissingColumnsError as exc:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(
            status_code=400,
            detail={
                "message": str(exc),
                "required_columns": REQUIRED_COLUMNS,
            },
        ) from exc
    except (MissingFileError, InvalidDataError, CertificateGeneratorError) as exc:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    job_id = uuid.uuid4().hex
    jobs[job_id] = GenerationJob(temp_dir=temp_dir)
    event_queue: queue.Queue = queue.Queue()

    worker = threading.Thread(
        target=_stream_generation_events,
        args=(
            job_id,
            excel_path,
            template_path,
            output_dir,
            certificate_seq_start,
            event_queue,
        ),
        daemon=True,
    )
    worker.start()

    return StreamingResponse(
        _ndjson_stream(event_queue, worker),
        media_type="application/x-ndjson",
    )


@app.get("/download/{job_id}")
async def download_certificates(job_id: str) -> FileResponse:
    """Download the generated ZIP for a completed job."""
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Download not found or expired.")

    job.ready.wait(timeout=300)

    if job.error:
        shutil.rmtree(job.temp_dir, ignore_errors=True)
        jobs.pop(job_id, None)
        raise HTTPException(status_code=422, detail=job.error)

    if not job.zip_path or not job.zip_path.is_file():
        raise HTTPException(status_code=404, detail="ZIP file is not ready yet.")

    zip_path = job.zip_path
    temp_dir = job.temp_dir
    jobs.pop(job_id, None)

    return FileResponse(
        path=zip_path,
        media_type="application/zip",
        filename="certificates.zip",
        background=BackgroundTask(lambda: shutil.rmtree(temp_dir, ignore_errors=True)),
    )
