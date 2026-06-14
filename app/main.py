from __future__ import annotations

import re
import uuid
import json
from pathlib import Path
from typing import Any

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.ai import analyze_document, ai_polish_document
from app.config import settings
from app.docx_pipeline import GenerationMetadata, extract_docx, render_document, validate_output
from app.schemas import JobRecord, JobStatus


settings.ensure_dirs()

app = FastAPI(title="CQU Study Guide Automator", version="0.1.0")
app.mount("/static", StaticFiles(directory=Path(__file__).parent / "static"), name="static")

jobs: dict[str, JobRecord] = {}


@app.get("/", include_in_schema=False)
def index() -> FileResponse:
    return FileResponse(Path(__file__).parent / "static" / "index.html")


@app.post("/generate")
async def generate(
    background_tasks: BackgroundTasks,
    source_docx: UploadFile = File(...),
    template_docx: UploadFile | None = File(default=None),
    unit_code: str = Form(default=""),
    unit_name: str = Form(default=""),
    week_number: str = Form(default=""),
    week_title: str = Form(default=""),
    version: str = Form(default=""),
    year: str = Form(default=""),
    term: str = Form(default=""),
    accepted_corrections: str = Form(default="[]"),
) -> JobRecord:
    if not source_docx.filename or not source_docx.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Upload a .docx source document.")

    job_id = uuid.uuid4().hex
    job_dir = settings.upload_dir / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    source_path = job_dir / safe_filename(source_docx.filename)
    await save_upload(source_docx, source_path)

    template_path = settings.default_template_path
    if template_docx and template_docx.filename:
        if not template_docx.filename.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail="Upload a .docx template document.")
        template_path = job_dir / safe_filename(template_docx.filename)
        await save_upload(template_docx, template_path)

    metadata = GenerationMetadata(
        unit_code=unit_code.strip(),
        unit_name=unit_name.strip(),
        week_number=week_number.strip(),
        week_title=week_title.strip(),
        version=version.strip(),
        year=year.strip(),
        term=term.strip(),
    )
    output_name = build_output_filename(source_docx.filename, metadata)
    output_path = settings.output_dir / f"{job_id}_{output_name}"
    corrections = parse_accepted_corrections(accepted_corrections)
    record = JobRecord(
        job_id=job_id,
        status=JobStatus.uploaded,
        message="Upload received.",
        filename=source_docx.filename,
        output_filename=output_name,
        output_path=output_path,
        progress=["uploaded: Upload received."],
    )
    jobs[job_id] = record
    background_tasks.add_task(
        run_generation_job,
        job_id,
        source_path,
        template_path,
        output_path,
        metadata,
        corrections,
    )
    return record


@app.post("/analyze")
async def analyze(
    source_docx: UploadFile = File(...),
    unit_code: str = Form(default=""),
    unit_name: str = Form(default=""),
    week_number: str = Form(default=""),
    week_title: str = Form(default=""),
    version: str = Form(default=""),
    year: str = Form(default=""),
    term: str = Form(default=""),
) -> dict[str, object]:
    if not source_docx.filename or not source_docx.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Upload a .docx source document.")

    analysis_id = uuid.uuid4().hex
    analysis_dir = settings.upload_dir / "analysis" / analysis_id
    analysis_dir.mkdir(parents=True, exist_ok=True)
    source_path = analysis_dir / safe_filename(source_docx.filename)
    await save_upload(source_docx, source_path)

    metadata = GenerationMetadata(
        unit_code=unit_code.strip(),
        unit_name=unit_name.strip(),
        week_number=week_number.strip(),
        week_title=week_title.strip(),
        version=version.strip(),
        year=year.strip(),
        term=term.strip(),
    )
    structured = extract_docx(source_path, metadata)
    hyperlinks = extract_hyperlinks(source_path)
    result = analyze_document(structured, hyperlinks)
    result["analysis_id"] = analysis_id
    return result


@app.get("/jobs/{job_id}")
def get_job(job_id: str) -> JobRecord:
    record = jobs.get(job_id)
    if record is None:
        raise HTTPException(status_code=404, detail="Job not found.")
    return record


@app.get("/download/{job_id}")
def download(job_id: str) -> FileResponse:
    record = jobs.get(job_id)
    if record is None:
        raise HTTPException(status_code=404, detail="Job not found.")
    if record.status != JobStatus.completed or not record.output_path or not record.output_path.exists():
        raise HTTPException(status_code=409, detail="Job is not complete.")
    return FileResponse(
        record.output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=record.output_filename,
    )


async def save_upload(upload: UploadFile, destination: Path) -> None:
    with destination.open("wb") as handle:
        while chunk := await upload.read(1024 * 1024):
            handle.write(chunk)


def run_generation_job(
    job_id: str,
    source_path: Path,
    template_path: Path,
    output_path: Path,
    metadata: GenerationMetadata,
    accepted_corrections: list[dict[str, str]],
) -> None:
    record = jobs[job_id]
    try:
        record.add_progress(JobStatus.extracting, "Extracting source DOCX structure, tables, and images.")
        structured = extract_docx(source_path, metadata)
        if accepted_corrections:
            applied = apply_accepted_corrections(structured, accepted_corrections)
            record.progress.append(f"validating: Applied {applied} accepted proofreading corrections.")
        record.add_progress(JobStatus.structuring, "Building structured JSON blocks.")
        if accepted_corrections:
            record.progress.append("structuring: Skipped automatic AI polish because manual corrections were selected.")
        else:
            structured = ai_polish_document(structured)
        record.add_progress(JobStatus.validating, "Checking structured blocks against extracted assets.")
        missing = [
            block.image_id
            for block in structured.body_blocks
            if block.type == "image" and block.image_id not in structured.images
        ]
        if missing:
            raise ValueError(f"Structured output referenced missing images: {', '.join(missing)}")
        record.add_progress(JobStatus.rendering, "Rendering into the CQU Word template.")
        render_document(structured, template_path, output_path)
        record.add_progress(JobStatus.checking_output, "Validating generated DOCX.")
        validate_output(output_path, structured)
        record.add_progress(JobStatus.completed, "Study guide is ready to download.")
    except Exception as exc:
        record.status = JobStatus.failed
        record.message = "Generation failed."
        record.error = str(exc)
        record.progress.append(f"failed: {exc}")


def build_output_filename(original: str, metadata: GenerationMetadata) -> str:
    parts = [metadata.unit_code, metadata.year, metadata.term]
    if metadata.week_number:
        parts.append(f"week-{metadata.week_number}")
    if metadata.version:
        parts.append(f"v{metadata.version}")
    clean_parts = [slug(part) for part in parts if part]
    if clean_parts:
        return "_".join(clean_parts) + ".docx"
    return safe_filename(original)


def safe_filename(filename: str) -> str:
    name = Path(filename).name
    return re.sub(r"[^A-Za-z0-9._() -]+", "_", name).strip() or "document.docx"


def slug(value: str) -> str:
    value = value.strip().replace(".", "-")
    return re.sub(r"[^A-Za-z0-9-]+", "-", value).strip("-").lower()


def extract_hyperlinks(path: Path) -> list[dict[str, str]]:
    from docx import Document

    document = Document(path)
    links: list[dict[str, str]] = []
    for rel in document.part.rels.values():
        if "hyperlink" not in rel.reltype:
            continue
        links.append({"text": rel.target_ref, "url": rel.target_ref, "status": "present"})
    return links


def parse_accepted_corrections(raw: str) -> list[dict[str, str]]:
    try:
        values = json.loads(raw or "[]")
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="accepted_corrections must be valid JSON.")
    if not isinstance(values, list):
        raise HTTPException(status_code=400, detail="accepted_corrections must be a JSON array.")

    corrections: list[dict[str, str]] = []
    for item in values:
        if not isinstance(item, dict):
            continue
        original = str(item.get("original", "")).strip()
        replacement = str(item.get("replacement", "")).strip()
        if not original or not replacement or original == replacement:
            continue
        corrections.append(
            {
                "block_index": item.get("block_index", ""),
                "original": original,
                "replacement": replacement,
            }
        )
    return corrections


def apply_accepted_corrections(structured: Any, corrections: list[dict[str, str]]) -> int:
    applied = 0
    for correction in corrections:
        if apply_correction_by_index(structured, correction) or apply_correction_globally(structured, correction):
            applied += 1
    return applied


def apply_correction_by_index(structured: Any, correction: dict[str, str]) -> bool:
    try:
        index = int(correction["block_index"])
    except (TypeError, ValueError):
        return False
    if index < 0 or index >= len(structured.body_blocks):
        return False
    return apply_correction_to_block(structured.body_blocks[index], correction)


def apply_correction_globally(structured: Any, correction: dict[str, str]) -> bool:
    for block in structured.body_blocks:
        if apply_correction_to_block(block, correction):
            return True
    return False


def apply_correction_to_block(block: Any, correction: dict[str, str]) -> bool:
    original = correction["original"]
    replacement = correction["replacement"]
    if block.text and original in block.text:
        block.text = block.text.replace(original, replacement, 1)
        return True
    for index, item in enumerate(block.items):
        if original in item:
            block.items[index] = item.replace(original, replacement, 1)
            return True
    return False
