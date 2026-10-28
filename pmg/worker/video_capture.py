"""
pmg.worker.video_capture
========================
Slideshow screen capture and frame sampling wrapper for the Windows PowerPoint
Validation Worker.
"""

from typing import Tuple, Dict, Any
from pmg.models import VideoMetadata, ExpectedTimeline
from pmg.video_pipeline import record_actual_slideshow


def run_worker_slideshow_capture(
    pptx_path: str,
    output_video_path: str,
    expected_timeline: ExpectedTimeline,
    target_fps: int = 30,
    timeout_sec: int = 240,
) -> Tuple[VideoMetadata, Dict[str, Any]]:
    """Launches full-screen Microsoft PowerPoint Slide Show and captures real-time playback."""
    return record_actual_slideshow(
        pptx_path=pptx_path,
        output_video_path=output_video_path,
        expected_timeline=expected_timeline,
        target_fps=target_fps,
        timeout_sec=timeout_sec,
    )

