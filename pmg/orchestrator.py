"""
pmg.orchestrator
================
Canonical PMG Orchestration Pipeline (Steps 1 through 9):
1. Build -> 2. Static Validation & Timeline -> 3. Native PowerPoint Video Export ->
4. Actual Slideshow Playback Recording -> 5. AI/CV Visual & Visibility Review ->
6. 3-Loop Self-Healing -> 7. Repeat Rendering & Recording ->
8. Final Acceptance -> 9. Validated PPTX Delivery (with embedded provenance manifest).

Supports v1.5.1 Worker Validation Modes (`local`, `remote`, `auto`, `static`) via
`PMG_QA_MODE`, `PMG_QA_WORKER_URL`, and `PMG_QA_WORKER_TOKEN`.
"""

import os
import json
import uuid
import shutil
import zipfile
from datetime import datetime
from typing import Optional, Callable, Dict, Any, Tuple, List

from pmg.version import get_engine_provenance, get_repo_root
from pmg.models import (
    ValidationStatus,
    ExpectedTimeline,
    VideoMetadata,
    ReviewDefect,
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
from pmg.worker.client import PMGWorkerClient


def _embed_manifest_in_pptx(pptx_path: str, manifest_dict: Dict[str, Any]) -> None:
    """
    Embeds a lightweight, non-intrusive provenance manifest (`ppt/pmg_execution_manifest.json`)
    inside the delivered `.pptx` ZIP container (Upgrade C) without altering PowerPoint slide XML
    or [Content_Types].xml validity.
    """
    if not os.path.exists(pptx_path):
        return
    compact_manifest = {
        "engine_name": manifest_dict.get("provenance", {}).get("engine_name", "pmg-skills"),
        "engine_version": manifest_dict.get("provenance", {}).get("engine_version", "1.5.1"),
        "git_commit": manifest_dict.get("provenance", {}).get("git_commit", "unknown"),
        "canonical_entrypoint": manifest_dict.get("provenance", {}).get(
            "canonical_entrypoint", "pmg.orchestrator.PMGOrchestrator.run_pipeline"
        ),
        "run_id": manifest_dict.get("run_id", ""),
        "validation_method": manifest_dict.get("validation_method", "local_powerpoint_com"),
        "worker_identity": manifest_dict.get("worker_identity", "local"),
        "repair_loops_executed": manifest_dict.get("repair_loops_executed", 0),
        "final_status": manifest_dict.get("status", "PASS"),
        "timestamp_utc": manifest_dict.get("provenance", {}).get("timestamp_utc", ""),
    }
    tmp_zip = pptx_path + ".manifest.tmp"
    try:
        with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "ppt/pmg_execution_manifest.json":
                    continue
                zout.writestr(item, zin.read(item.filename))
            zout.writestr(
                "ppt/pmg_execution_manifest.json",
                json.dumps(compact_manifest, indent=2, ensure_ascii=False).encode("utf-8"),
            )
        shutil.move(tmp_zip, pptx_path)
    except Exception:
        if os.path.exists(tmp_zip):
            try:
                os.remove(tmp_zip)
            except Exception:
                pass


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
        worker_mode: str = "auto",
        worker_url: Optional[str] = None,
        worker_token: Optional[str] = None,
    ) -> Tuple[Dict[str, Any], ExpectedTimeline, VideoMetadata, VideoMetadata, Dict[str, Any], ReviewReport, str, str]:
        """
        Executes Steps 2, 3, 4, and 5 on candidate_pptx (either locally or via remote Windows worker),
        returning (static_val, timeline, export_meta, rec_meta, telemetry, report, validation_method, worker_identity).
        """
        # STEP 2: Static validation & Expected Timeline
        static_val = self.validator.validate(candidate_pptx)
        timeline = extract_expected_timeline(candidate_pptx, hold_duration_ms=hold_duration_ms)
        timeline.save(os.path.join(qa_dir, "expected_timeline.json"))

        export_mp4 = os.path.join(qa_dir, "powerpoint_export.mp4")
        record_mp4 = os.path.join(qa_dir, "actual_slideshow_recording.mp4")

        eff_mode = (worker_mode or os.environ.get("PMG_QA_MODE") or "auto").lower()
        if not require_video_playback or eff_mode == "static":
            export_meta = VideoMetadata(path=export_mp4, error="Skipped playback validation")
            rec_meta = VideoMetadata(path=record_mp4, error="Skipped playback validation")
            telemetry: Dict[str, Any] = {}
            status = ValidationStatus.NOT_VERIFIED if static_val["is_valid"] else ValidationStatus.FAIL
            defects = [
                ReviewDefect(
                    category=sd.get("category", "openxml_structure"),
                    severity=sd.get("severity", "major"),
                    slide_number=sd.get("slide_number", 0),
                    object_id=sd.get("object_id"),
                    description=sd.get("description", ""),
                    remediation_hint=sd.get("remediation_hint", ""),
                )
                for sd in static_val.get("defects", [])
            ]
            report = ReviewReport(
                status=status,
                defects=defects,
                summary="Static validation only (NOT_VERIFIED for real PowerPoint playback).",
            )
            return static_val, timeline, export_meta, rec_meta, telemetry, report, "static_only", "local"

        eff_worker_url = worker_url or os.environ.get("PMG_QA_WORKER_URL")
        local_ppt_ok = is_powerpoint_available()

        use_remote = False
        if eff_mode == "remote":
            use_remote = True
        elif eff_mode == "auto" and not local_ppt_ok and eff_worker_url:
            use_remote = True

        if use_remote:
            client = PMGWorkerClient(worker_url=eff_worker_url, worker_token=worker_token)
            w_res = client.validate_remote(candidate_pptx, hold_duration_ms=hold_duration_ms)
            exp_dict = w_res.export_video or {}
            rec_dict = w_res.recording_video or {}
            export_meta = VideoMetadata(
                **{k: v for k, v in exp_dict.items() if k in VideoMetadata.__dataclass_fields__}
            ) if exp_dict else VideoMetadata(path=export_mp4, error=w_res.error or "Remote worker error")
            rec_meta = VideoMetadata(
                **{k: v for k, v in rec_dict.items() if k in VideoMetadata.__dataclass_fields__}
            ) if rec_dict else VideoMetadata(path=record_mp4, error=w_res.error or "Remote worker error")
            telemetry = w_res.recording_telemetry or {"worker_job_id": w_res.job_id}

            try:
                v_status = ValidationStatus(w_res.validation_status)
            except Exception:
                v_status = ValidationStatus.BLOCKED

            rev_dict = w_res.review_report or {}
            raw_defects = rev_dict.get("defects", static_val.get("defects", []))
            defects = [
                ReviewDefect(
                    category=d.get("category", "openxml_structure"),
                    severity=d.get("severity", "major"),
                    slide_number=d.get("slide_number", 0),
                    object_id=d.get("object_id"),
                    description=d.get("description", ""),
                    remediation_hint=d.get("remediation_hint", ""),
                )
                for d in raw_defects
            ]
            report = ReviewReport(
                status=v_status,
                defects=defects,
                visual_quality_passed=rev_dict.get("visual_quality_passed", v_status == ValidationStatus.PASS),
                animation_behavior_passed=rev_dict.get("animation_behavior_passed", v_status == ValidationStatus.PASS),
                temporal_quality_passed=rev_dict.get("temporal_quality_passed", v_status == ValidationStatus.PASS),
                capture_quality_passed=rev_dict.get("capture_quality_passed", v_status == ValidationStatus.PASS),
                object_visibility_passed=rev_dict.get("object_visibility_passed", v_status == ValidationStatus.PASS),
                slide_steady_states_verified=rev_dict.get("slide_steady_states_verified", 0),
                transitions_verified=rev_dict.get("transitions_verified", 0),
                tracked_entities=rev_dict.get("tracked_entities", {}),
                object_visibility_checks=rev_dict.get("object_visibility_checks", []),
                temporal_deviation_ms=rev_dict.get("temporal_deviation_ms", 0.0),
                geometric_deviation_pct=rev_dict.get("geometric_deviation_pct", 0.0),
                export_vs_recording_similarity=rev_dict.get("export_vs_recording_similarity", 0.0),
                summary=rev_dict.get("summary") or w_res.error or f"Remote worker validation: {v_status.value}",
            )
            with open(os.path.join(qa_dir, "video_analysis.json"), "w", encoding="utf-8") as f:
                json.dump(w_res.video_analysis or {}, f, indent=2, ensure_ascii=False)
            report.save(os.path.join(qa_dir, "review_report.json"))
            return (
                static_val,
                timeline,
                export_meta,
                rec_meta,
                telemetry,
                report,
                "remote_windows_worker",
                w_res.worker_identity or (eff_worker_url or "remote"),
            )

        if not local_ppt_ok:
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
                summary="BLOCKED: Microsoft PowerPoint Desktop or PMG_QA_WORKER_URL required for real playback verification.",
            )
            return static_val, timeline, export_meta, rec_meta, telemetry, report, "blocked_no_powerpoint", "none"

        # STEP 3: Native PowerPoint Video Export (+ Edit-Mode Reference Slide PNGs)
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

        return static_val, timeline, export_meta, rec_meta, telemetry, report, "local_powerpoint_com", "local"

    def run_pipeline(
        self,
        output_pptx_path: str,
        builder_fn: Optional[Callable[[str], str]] = None,
        existing_pptx_path: Optional[str] = None,
        hold_duration_ms: int = 1200,
        max_repair_loops: int = 3,
        require_video_playback: bool = True,
        worker_mode: Optional[str] = None,
        worker_url: Optional[str] = None,
        worker_token: Optional[str] = None,
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
        eff_mode = worker_mode or os.environ.get("PMG_QA_MODE", "auto")

        # Initial Pass (Steps 2, 3, 4, 5)
        static_val, timeline, export_meta, rec_meta, telemetry, report, val_method, worker_id = (
            self._run_validation_and_playback_pass(
                candidate_pptx=candidate_pptx,
                qa_dir=qa_dir,
                hold_duration_ms=hold_duration_ms,
                require_video_playback=require_video_playback,
                worker_mode=eff_mode,
                worker_url=worker_url,
                worker_token=worker_token,
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
            static_val, timeline, export_meta, rec_meta, telemetry, report, val_method, worker_id = (
                self._run_validation_and_playback_pass(
                    candidate_pptx=candidate_pptx,
                    qa_dir=qa_dir,
                    hold_duration_ms=hold_duration_ms,
                    require_video_playback=require_video_playback,
                    worker_mode=eff_mode,
                    worker_url=worker_url,
                    worker_token=worker_token,
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
            provenance=get_engine_provenance(
                entrypoint="pmg.orchestrator.PMGOrchestrator.run_pipeline",
                validation_method=val_method,
                worker_identity=worker_id,
            ),
            source_pptx=candidate_pptx,
            delivered_pptx=delivered_path,
            qa_dir=qa_dir,
            status=report.status,
            validation_method=val_method,
            worker_identity=worker_id,
            static_validation=static_val,
            export_video=export_meta.to_dict(),
            recording_video=rec_meta.to_dict(),
            review_summary=report.to_dict(),
            repair_loops_executed=loops_executed,
            repair_history=repair_history,
        )
        manifest.save(os.path.join(qa_dir, "execution_manifest.json"))

        if delivered_path and os.path.exists(delivered_path):
            _embed_manifest_in_pptx(delivered_path, manifest.to_dict())

        return manifest
