"""
pmg.orchestrator
================
Canonical PMG Orchestration Pipeline (Steps 1 through 9):
1. Build -> 2. Static Validation & Timeline -> 3. Native PowerPoint Video Export ->
4. Actual Slideshow Playback Recording -> 5. AI/CV Visual Review ->
6. 3-Loop Self-Healing -> 7. Repeat Rendering & Recording ->
8. Final Acceptance -> 9. Validated PPTX Delivery.
"""

import os
import json
import uuid
import shutil
from datetime import datetime
from typing import Optional, Callable, Dict, Any, Tuple, List

from pmg.version import get_engine_provenance, get_repo_root
from pmg.models import (
    ValidationStatus,
    ExpectedTimeline,
    VideoMetadata,
    ReviewReport,
    ExecutionManifest,
)
from pmg.validator import StaticValidator, extract_expected_timeline
from pmg.video_pipeline import (
    is_powerpoint_available,
    export_video_with_powerpoint,
    record_actual_slideshow,
)
from pmg.visual_reviewer import VisualAnimationReviewer
from pmg.self_healing import execute_repair_loop


class PMGOrchestrator:
    """
    Canonical entrypoint for building, validating, rendering, recording,
    reviewing, self-healing, and delivering PMG presentations.
    """

    def __init__(self, qa_root_dir: Optional[str] = None):
        self.qa_root_dir = qa_root_dir or os.path.join(get_repo_root(), "internal_qa")
        os.makedirs(self.qa_root_dir, exist_ok=True)
        self.validator = StaticValidator()
        self.reviewer = VisualAnimationReviewer()

    def create_run_dir(self, run_id: Optional[str] = None) -> Tuple[str, str]:
        if not run_id:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            short_uuid = uuid.uuid4().hex[:6]
            run_id = f"run_{ts}_{short_uuid}"
        qa_dir = os.path.join(self.qa_root_dir, run_id)
        os.makedirs(qa_dir, exist_ok=True)
        os.makedirs(os.path.join(qa_dir, "frame_samples"), exist_ok=True)
        return run_id, qa_dir

    def _run_validation_and_playback_pass(
        self,
        candidate_pptx: str,
        qa_dir: str,
        hold_duration_ms: int = 1200,
        require_video_playback: bool = True,
    ) -> Tuple[Dict[str, Any], ExpectedTimeline, VideoMetadata, VideoMetadata, Dict[str, Any], ReviewReport]:
        """Executes Steps 2, 3, 4, and 5 on candidate_pptx, generating fresh videos and reports."""
        # STEP 2: Static validation & Expected Timeline
        static_val = self.validator.validate(candidate_pptx)
        timeline = extract_expected_timeline(candidate_pptx, hold_duration_ms=hold_duration_ms)
        timeline.save(os.path.join(qa_dir, "expected_timeline.json"))

        export_mp4 = os.path.join(qa_dir, "powerpoint_export.mp4")
        record_mp4 = os.path.join(qa_dir, "actual_slideshow_recording.mp4")

        if not require_video_playback:
            export_meta = VideoMetadata(path=export_mp4, error="Skipped playback validation")
            rec_meta = VideoMetadata(path=record_mp4, error="Skipped playback validation")
            telemetry: Dict[str, Any] = {}
            status = ValidationStatus.NOT_VERIFIED if static_val["is_valid"] else ValidationStatus.FAIL
            report = ReviewReport(
                status=status,
                summary="Static validation only (NOT_VERIFIED for real PowerPoint playback).",
            )
            return static_val, timeline, export_meta, rec_meta, telemetry, report

        if not is_powerpoint_available():
            export_meta = VideoMetadata(
                path=export_mp4,
                error="BLOCKED: Microsoft PowerPoint Desktop COM is not available in this environment.",
            )
            rec_meta = VideoMetadata(
                path=record_mp4,
                error="BLOCKED: Microsoft PowerPoint Desktop COM is not available in this environment.",
            )
            telemetry = {"status": "BLOCKED"}
            report = ReviewReport(
                status=ValidationStatus.BLOCKED,
                summary="BLOCKED: Microsoft PowerPoint Desktop required for real playback verification.",
            )
            return static_val, timeline, export_meta, rec_meta, telemetry, report

        # STEP 3: Native PowerPoint Video Export
        slide_dur_sec = max(1.0, hold_duration_ms / 1000.0)
        export_meta = export_video_with_powerpoint(
            pptx_path=candidate_pptx,
            output_video_path=export_mp4,
            slide_duration=slide_dur_sec,
            requested_fps=60,
            resolution=1080,
        )

        # STEP 4: Record Actual Slideshow Playback
        rec_meta, telemetry = record_actual_slideshow(
            pptx_path=candidate_pptx,
            output_video_path=record_mp4,
            expected_timeline=timeline,
            target_fps=30,
        )

        # STEP 5: AI & Computer Vision Visual Review
        report, video_analysis = self.reviewer.review(
            export_meta=export_meta,
            recording_meta=rec_meta,
            recording_telemetry=telemetry,
            expected_timeline=timeline,
            static_validation=static_val,
            qa_dir=qa_dir,
        )

        with open(os.path.join(qa_dir, "video_analysis.json"), "w", encoding="utf-8") as f:
            json.dump(video_analysis, f, indent=2, ensure_ascii=False)
        report.save(os.path.join(qa_dir, "review_report.json"))

        return static_val, timeline, export_meta, rec_meta, telemetry, report

    def run_pipeline(
        self,
        output_pptx_path: str,
        builder_fn: Optional[Callable[[str], str]] = None,
        existing_pptx_path: Optional[str] = None,
        hold_duration_ms: int = 1200,
        max_repair_loops: int = 3,
        require_video_playback: bool = True,
        run_id: Optional[str] = None,
    ) -> ExecutionManifest:
        """
        Executes the full 9-Step Mandatory Generation, Video Review, 3-Loop Self-Healing,
        and Delivery Gating Pipeline.
        """
        run_id, qa_dir = self.create_run_dir(run_id=run_id)
        candidate_pptx = os.path.join(qa_dir, "candidate.pptx")

        # STEP 1: Build or stage candidate.pptx
        if builder_fn is not None:
            builder_fn(candidate_pptx)
        elif existing_pptx_path is not None and os.path.exists(existing_pptx_path):
            shutil.copyfile(existing_pptx_path, candidate_pptx)
        else:
            raise ValueError("Either builder_fn or an existing_pptx_path must be provided.")

        # Preserve original candidate before any repair loops
        backup_pptx = os.path.join(qa_dir, "candidate.pptx.pre_heal.bak")
        shutil.copyfile(candidate_pptx, backup_pptx)

        repair_history: List[Dict[str, Any]] = []

        # Initial Pass (Steps 2, 3, 4, 5)
        static_val, timeline, export_meta, rec_meta, telemetry, report = (
            self._run_validation_and_playback_pass(
                candidate_pptx=candidate_pptx,
                qa_dir=qa_dir,
                hold_duration_ms=hold_duration_ms,
                require_video_playback=require_video_playback,
            )
        )

        # STEPS 6 & 7: 3-Loop Self-Healing + Repeat Rendering & Recording
        loops_executed = 0
        while (
            report.status == ValidationStatus.FAIL
            and loops_executed < max_repair_loops
        ):
            loops_executed += 1
            repair_record = execute_repair_loop(
                pptx_path=candidate_pptx,
                loop_number=loops_executed,
                defects=report.defects,
                static_validation=static_val,
            )
            repair_history.append(repair_record)

            # Re-run Steps 2, 3, 4, 5 with fresh video export and fresh slideshow recording!
            static_val, timeline, export_meta, rec_meta, telemetry, report = (
                self._run_validation_and_playback_pass(
                    candidate_pptx=candidate_pptx,
                    qa_dir=qa_dir,
                    hold_duration_ms=hold_duration_ms,
                    require_video_playback=require_video_playback,
                )
            )
            repair_record["post_loop_status"] = report.status.value

        with open(os.path.join(qa_dir, "repair_history.json"), "w", encoding="utf-8") as f:
            json.dump(repair_history, f, indent=2, ensure_ascii=False)

        # STEPS 8 & 9: Final Acceptance & Delivery Gate
        delivered_path: Optional[str] = None
        if report.status == ValidationStatus.PASS:
            abs_out = os.path.abspath(output_pptx_path)
            os.makedirs(os.path.dirname(abs_out), exist_ok=True)
            shutil.copyfile(candidate_pptx, abs_out)
            delivered_path = abs_out

        manifest = ExecutionManifest(
            run_id=run_id,
            provenance=get_engine_provenance(entrypoint="pmg.orchestrator.PMGOrchestrator.run_pipeline"),
            source_pptx=candidate_pptx,
            delivered_pptx=delivered_path,
            qa_dir=qa_dir,
            status=report.status,
            static_validation=static_val,
            export_video=export_meta.to_dict(),
            recording_video=rec_meta.to_dict(),
            review_summary=report.to_dict(),
            repair_loops_executed=loops_executed,
            repair_history=repair_history,
        )
        manifest.save(os.path.join(qa_dir, "execution_manifest.json"))
        return manifest

