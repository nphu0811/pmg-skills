"""
pmg.worker.jobs
===============
Manages validation job lifecycle, progress states, timeouts, cancellation,
and isolated workspace cleanup for the PMG Windows PowerPoint Validation Worker.
"""

import os
import time
import uuid
import shutil
import datetime
import threading
from typing import Dict, Any, Optional

from pmg.models import ValidationStatus, VideoMetadata, ReviewReport, ReviewDefect
from pmg.validator import StaticValidator, extract_expected_timeline, inspect_powerpoint_compatibility
from pmg.worker.powerpoint_probe import verify_clean_powerpoint_open, file_sha256
from pmg.visual_reviewer import VisualAnimationReviewer
from pmg.worker.result_schema import WorkerJobRequest, WorkerJobResult
from pmg.worker.security import (
    validate_safe_pptx_upload,
    IsolatedJobWorkspace,
    DEFAULT_MAX_PPTX_BYTES,
)
from pmg.worker.powerpoint_runner import is_powerpoint_available, run_worker_native_export
from pmg.worker.video_capture import run_worker_slideshow_capture


class WorkerJobManager:
    """
    Executes and tracks PowerPoint validation jobs on a Windows worker host.
    Serializes COM access via an internal lock so concurrent HTTP requests do not
    collide on the desktop PowerPoint Slide Show window.
    """

    def __init__(
        self,
        worker_identity: str = "pmg-windows-worker",
        max_file_size_bytes: int = DEFAULT_MAX_PPTX_BYTES,
        cleanup_artifacts: bool = True,
    ):
        self.worker_identity = worker_identity
        self.max_file_size_bytes = max_file_size_bytes
        self.cleanup_artifacts = cleanup_artifacts
        self._jobs: Dict[str, WorkerJobResult] = {}
        self._cancelled: set = set()
        self._lock = threading.Lock()
        self._com_lock = threading.Lock()
        self.validator = StaticValidator()
        self.reviewer = VisualAnimationReviewer()
        self._hash_cache: Dict[str, WorkerJobResult] = {}
        self._job_timestamps: Dict[str, float] = {}

    def get_cached_by_hash(self, sha256_hash: str) -> Optional[WorkerJobResult]:
        with self._lock:
            return self._hash_cache.get(sha256_hash)

    def cleanup_stale_jobs(self, max_age_sec: float = 3600) -> int:
        now = time.time()
        to_del = []
        with self._lock:
            for jid, t_created in self._job_timestamps.items():
                if (now - t_created) > max_age_sec:
                    to_del.append(jid)
            for jid in to_del:
                self._jobs.pop(jid, None)
                self._job_timestamps.pop(jid, None)
        return len(to_del)

    def active_job_count(self) -> int:
        with self._lock:
            return sum(
                1
                for j in self._jobs.values()
                if j.state in ("queued", "validating", "exporting_video", "recording_slideshow", "analyzing")
            )

    def get_job(self, job_id: str) -> Optional[WorkerJobResult]:
        with self._lock:
            return self._jobs.get(job_id)

    def cancel_job(self, job_id: str) -> bool:
        with self._lock:
            self._cancelled.add(job_id)
            if job_id in self._jobs and self._jobs[job_id].state not in ("completed", "failed", "timed_out"):
                self._jobs[job_id].state = "cancelled"
                self._jobs[job_id].validation_status = ValidationStatus.BLOCKED.value
                self._jobs[job_id].error = "Job cancelled by caller."
                return True
        return False

    def execute_job_sync(
        self,
        pptx_source_path: str,
        original_filename: str = "submission.pptx",
        hold_duration_ms: int = 1200,
        target_fps: int = 30,
        export_fps: int = 60,
        timeout_sec: int = 240,
        job_id: Optional[str] = None,
        output_qa_dir: Optional[str] = None,
        allow_cache: bool = True,
    ) -> WorkerJobResult:
        """
        Runs a complete validation job (Security Check -> Static Validation ->
        Native PowerPoint Export -> Live Slideshow Recording -> AI Visual Review)
        and returns a structured WorkerJobResult.
        """
        jid = job_id or f"job_{uuid.uuid4().hex[:10]}"
        t0 = time.perf_counter()
        input_hash = file_sha256(pptx_source_path) if os.path.exists(pptx_source_path) else ""
        if allow_cache and input_hash:
            cached = self.get_cached_by_hash(input_hash)
            if cached is not None and cached.state == "completed":
                return cached
        result = WorkerJobResult(
            job_id=jid,
            state="queued",
            validation_status=ValidationStatus.NOT_VERIFIED.value,
            worker_identity=self.worker_identity,
            validated_pptx_sha256=input_hash,
            created_timestamp=datetime.datetime.utcnow().isoformat() + "Z",
        )
        with self._lock:
            self._jobs[jid] = result
            self._job_timestamps[jid] = time.time()
        result.checks_performed.append("security_scan")

        # 1. Security validation before opening in PowerPoint
        is_safe, sec_msg = validate_safe_pptx_upload(
            file_path=pptx_source_path,
            original_filename=original_filename,
            max_bytes=self.max_file_size_bytes,
        )
        if not is_safe:
            result.state = "failed"
            result.validation_status = ValidationStatus.FAIL.value
            result.error = sec_msg
            result.elapsed_sec = round(time.perf_counter() - t0, 3)
            result.completed_timestamp = datetime.datetime.utcnow().isoformat() + "Z"
            if result.clean_open and result.clean_open.get("version"):
                result.powerpoint_version = str(result.clean_open.get("version"))
            if input_hash and result.validation_status == ValidationStatus.PASS.value:
                with self._lock:
                    self._hash_cache[input_hash] = result
            return result

        with IsolatedJobWorkspace(jid, cleanup_on_exit=self.cleanup_artifacts) as ws_dir:
            qa_dir = output_qa_dir or ws_dir
            os.makedirs(qa_dir, exist_ok=True)
            os.makedirs(os.path.join(qa_dir, "frame_samples"), exist_ok=True)

            staged_pptx = os.path.join(ws_dir, "candidate.pptx")
            shutil.copyfile(pptx_source_path, staged_pptx)

            if jid in self._cancelled:
                result.state = "cancelled"
                result.validation_status = ValidationStatus.BLOCKED.value
                result.error = "Job cancelled before execution."
                return result

            # 2. Static validation & ExpectedTimeline extraction
            result.state = "validating"
            static_val = self.validator.validate(staged_pptx)
            result.static_validation = static_val
            if not static_val["is_valid"]:
                result.state = "failed"
                result.validation_status = ValidationStatus.FAIL.value
                result.review_report = ReviewReport(
                    status=ValidationStatus.FAIL,
                    defects=[ReviewDefect(**d) for d in static_val.get("defects", [])],
                    summary="Read-only package audit failed; worker never repairs uploads.").to_dict()
                return result
            timeline = extract_expected_timeline(staged_pptx, hold_duration_ms=hold_duration_ms)
            timeline.save(os.path.join(qa_dir, "expected_timeline.json"))
            result.static_validation = static_val
            result.expected_timeline = timeline.to_dict()

            if not is_powerpoint_available():
                result.state = "failed"
                result.validation_status = ValidationStatus.BLOCKED.value
                result.error = "BLOCKED: Microsoft PowerPoint Desktop COM is not available on this worker."
                result.review_report = ReviewReport(
                    status=ValidationStatus.BLOCKED,
                    summary=result.error,
                ).to_dict()
                result.elapsed_sec = round(time.perf_counter() - t0, 3)
            result.completed_timestamp = datetime.datetime.utcnow().isoformat() + "Z"
            if result.clean_open and result.clean_open.get("version"):
                result.powerpoint_version = str(result.clean_open.get("version"))
            if input_hash and result.validation_status == ValidationStatus.PASS.value:
                with self._lock:
                    self._hash_cache[input_hash] = result
                return result

            export_mp4 = os.path.join(qa_dir, "powerpoint_export.mp4")
            record_mp4 = os.path.join(qa_dir, "actual_slideshow_recording.mp4")

            with self._com_lock:
                if (time.perf_counter() - t0) > timeout_sec:
                    result.state = "timed_out"
                    result.validation_status = ValidationStatus.BLOCKED.value
                    result.error = f"Job timed out waiting for COM lock ({timeout_sec}s)."
                    return result

                before_hash = file_sha256(staged_pptx)
                result.clean_open = verify_clean_powerpoint_open(staged_pptx)
                if not result.clean_open.get("is_clean"):
                    result.state = "failed"
                    result.validation_status = ValidationStatus.FAIL.value
                    result.error = result.clean_open.get("error", "Worker PowerPoint opening was not clean.")
                    return result

                # 3. Native PowerPoint Video Export (+ Edit-Mode Reference Slide PNGs)
                result.state = "exporting_video"
                slide_dur_sec = max(1.0, hold_duration_ms / 1000.0)
                exp_meta = run_worker_native_export(
                    pptx_path=staged_pptx,
                    output_video_path=export_mp4,
                    slide_duration_sec=slide_dur_sec,
                    requested_fps=export_fps,
                    resolution=1080,
                    timeout_sec=max(30, int(timeout_sec * 0.6)),
                )
                result.export_video = exp_meta.to_dict()

                if jid in self._cancelled:
                    result.state = "cancelled"
                    result.validation_status = ValidationStatus.BLOCKED.value
                    result.error = "Job cancelled after video export."
                    return result

                # 4. Record Actual Slideshow Playback
                result.state = "recording_slideshow"
                rec_meta, telemetry = run_worker_slideshow_capture(
                    pptx_path=staged_pptx,
                    output_video_path=record_mp4,
                    expected_timeline=timeline,
                    target_fps=target_fps,
                    timeout_sec=max(30, int(timeout_sec * 0.8)),
                )
                result.recording_video = rec_meta.to_dict()
                result.recording_telemetry = telemetry
                result.clean_open = verify_clean_powerpoint_open(staged_pptx)
                result.validated_pptx_sha256 = file_sha256(staged_pptx)

            # 5. AI & Computer Vision Visual Review
            result.state = "analyzing"
            report, video_analysis = self.reviewer.review(
                export_meta=exp_meta,
                recording_meta=rec_meta,
                recording_telemetry=telemetry,
                expected_timeline=timeline,
                static_validation=static_val,
                qa_dir=qa_dir,
            )
            result.review_report = report.to_dict()
            compatibility_defects = inspect_powerpoint_compatibility(static_val, result.clean_open)
            if result.validated_pptx_sha256 != before_hash:
                compatibility_defects.append(ReviewDefect("powerpoint_compatibility", "critical", 0, None,
                                                          "Uploaded PPTX bytes changed during worker validation."))
            if compatibility_defects:
                report.defects.extend(compatibility_defects)
                report.status = ValidationStatus.FAIL
                result.review_report = report.to_dict()
            result.video_analysis = video_analysis
            result.validation_status = report.status.value
            result.state = "completed"
            result.elapsed_sec = round(time.perf_counter() - t0, 3)
            result.completed_timestamp = datetime.datetime.utcnow().isoformat() + "Z"
            if result.clean_open and result.clean_open.get("version"):
                result.powerpoint_version = str(result.clean_open.get("version"))
            if input_hash and result.validation_status == ValidationStatus.PASS.value:
                with self._lock:
                    self._hash_cache[input_hash] = result

        return result

