"""
pmg.worker.powerpoint_runner
============================
Manages Microsoft PowerPoint Desktop COM readiness checks, native video export,
edit-mode reference image export, and safe PowerPoint process cleanup for the
Windows PowerPoint Validation Worker.
"""

import os
import sys
import shutil
import subprocess
from typing import Dict, Any, List

from pmg.version import VERSION
from pmg.models import VideoMetadata
from pmg.video_pipeline import (
    is_powerpoint_available,
    export_video_with_powerpoint,
    export_edit_mode_slide_images,
)
from pmg.worker.result_schema import WorkerHealthStatus


def check_worker_health(active_jobs: int = 0, max_file_size_mb: int = 100) -> WorkerHealthStatus:
    """
    Runs a full environment diagnostic (`pmg worker doctor` / `/health`) checking:
    - Windows OS
    - Microsoft PowerPoint Desktop COM registration
    - FFmpeg / FFprobe availability
    - OpenCV availability
    """
    ppt_ok = is_powerpoint_available()
    ffmpeg_ok = shutil.which("ffmpeg") is not None and shutil.which("ffprobe") is not None
    try:
        import cv2  # noqa: F401
        cv_ok = True
    except Exception:
        cv_ok = False

    if ppt_ok and ffmpeg_ok and cv_ok:
        status = "healthy"
    elif cv_ok and ffmpeg_ok:
        status = "blocked"
    else:
        status = "degraded"

    return WorkerHealthStatus(
        status=status,
        platform=sys.platform,
        powerpoint_available=ppt_ok,
        ffmpeg_available=ffmpeg_ok,
        opencv_available=cv_ok,
        worker_version=VERSION,
        active_jobs=active_jobs,
        max_file_size_mb=max_file_size_mb,
        details={
            "ffmpeg_path": shutil.which("ffmpeg") or "",
            "ffprobe_path": shutil.which("ffprobe") or "",
        },
    )


def run_worker_native_export(
    pptx_path: str,
    output_video_path: str,
    slide_duration_sec: float = 1.5,
    requested_fps: int = 60,
    resolution: int = 1080,
    timeout_sec: int = 180,
) -> VideoMetadata:
    """Executes native PowerPoint MP4 export and edit-mode slide reference PNG export."""
    return export_video_with_powerpoint(
        pptx_path=pptx_path,
        output_video_path=output_video_path,
        slide_duration=slide_duration_sec,
        requested_fps=requested_fps,
        resolution=resolution,
        timeout_sec=timeout_sec,
    )

