"""
pmg.worker.result_schema
========================
Structured schemas for the PMG Windows PowerPoint Validation Worker (Upgrade A).
"""

from dataclasses import dataclass, field, asdict
from typing import Dict, Any, List, Optional
import json


@dataclass
class WorkerHealthStatus:
    status: str  # "healthy", "degraded", "blocked"
    platform: str
    powerpoint_available: bool
    ffmpeg_available: bool
    opencv_available: bool
    worker_version: str
    active_jobs: int = 0
    max_file_size_mb: int = 100
    details: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class WorkerJobRequest:
    job_id: str
    filename: str
    file_size_bytes: int
    hold_duration_ms: int = 1200
    target_fps: int = 30
    export_fps: int = 60
    timeout_sec: int = 240
    expected_timeline: Optional[Dict[str, Any]] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class WorkerJobResult:
    job_id: str
    state: str  # "queued", "validating", "exporting_video", "recording_slideshow", "analyzing", "completed", "failed", "cancelled", "timed_out"
    validation_status: str  # "PASS", "FAIL", "BLOCKED", "NOT_VERIFIED"
    worker_identity: str = "pmg-windows-worker"
    static_validation: Dict[str, Any] = field(default_factory=dict)
    expected_timeline: Dict[str, Any] = field(default_factory=dict)
    export_video: Dict[str, Any] = field(default_factory=dict)
    recording_video: Dict[str, Any] = field(default_factory=dict)
    recording_telemetry: Dict[str, Any] = field(default_factory=dict)
    review_report: Dict[str, Any] = field(default_factory=dict)
    video_analysis: Dict[str, Any] = field(default_factory=dict)
    error: Optional[str] = None
    elapsed_sec: float = 0.0
    clean_open: Dict[str, Any] = field(default_factory=dict)
    validated_pptx_sha256: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), indent=2, ensure_ascii=False)

