from __future__ import annotations

from enum import StrEnum
from pathlib import Path

from pydantic import BaseModel, Field


class JobStatus(StrEnum):
    uploaded = "uploaded"
    extracting = "extracting"
    structuring = "structuring"
    validating = "validating"
    rendering = "rendering"
    checking_output = "checking_output"
    completed = "completed"
    failed = "failed"


class JobRecord(BaseModel):
    job_id: str
    status: JobStatus
    message: str = ""
    filename: str = ""
    output_filename: str = ""
    output_path: Path | None = None
    error: str | None = None
    progress: list[str] = Field(default_factory=list)

    def add_progress(self, status: JobStatus, message: str) -> None:
        self.status = status
        self.message = message
        self.progress.append(f"{status.value}: {message}")
