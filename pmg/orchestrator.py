"""
pmg.orchestrator
================
Canonical PMG production and delivery pipeline:
Build/package -> read-only audit -> clean-open preflight -> native export ->
actual recording -> visual review -> at most three repair attempts with fresh
packaging/playback -> final destination reopen and immutable-byte certification.

Supports v1.5.1 Worker Validation Modes (`local`, `remote`, `auto`, `static`) via
`PMG_QA_MODE`, `PMG_QA_WORKER_URL`, and `PMG_QA_WORKER_TOKEN`.
"""

import os
import json
import uuid
import shutil
import zipfile
import tempfile
from lxml import etree
from datetime import datetime
from typing import Optional, Callable, Dict, Any, Tuple, List

from pmg.version import VERSION, get_engine_provenance, get_repo_root
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
from pmg.worker.powerpoint_probe import file_sha256, verify_clean_powerpoint_open


def _embed_manifest_in_pptx(pptx_path: str, manifest_dict: Dict[str, Any], write_companion: bool = True) -> None:
    """Atomically package provenance with a valid MIME type before validation.

    Embedded provenance is not the certificate: the final byte hash and delivery
    status live in the external manifest, avoiding a self-referential hash.
    Packaging failures must propagate to the orchestrator; never silently pass.
    """
    provenance = manifest_dict.get("provenance", {})
    status = manifest_dict.get("status", ValidationStatus.NOT_VERIFIED)
    compact_manifest = {
        "engine_name": provenance.get("engine_name", "pmg-skills"),
        "engine_version": provenance.get("engine_version", VERSION),
        "git_commit": provenance.get("git_commit", "unknown"),
        "canonical_entrypoint": provenance.get("canonical_entrypoint", "pmg.orchestrator.PMGOrchestrator.run_pipeline"),
        "run_id": manifest_dict.get("run_id", ""),
        "validation_method": manifest_dict.get("validation_method", "pending"),
        "worker_identity": manifest_dict.get("worker_identity", "pending"),
        "repair_loops_executed": manifest_dict.get("repair_loops_executed", 0),
        "final_status": status.value if isinstance(status, ValidationStatus) else status,
        "certification_source": "external_companion_manifest",
        "timestamp_utc": provenance.get("timestamp_utc", ""),
    }
    descriptor, temporary = tempfile.mkstemp(suffix=".pptx", dir=os.path.dirname(os.path.abspath(pptx_path)))
    os.close(descriptor)
    try:
        namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
        with zipfile.ZipFile(pptx_path) as source, zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as output:
            for item in source.infolist():
                if item.filename == "ppt/pmg_execution_manifest.json":
                    continue
                data = source.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    root = etree.fromstring(data)
                    for node in list(root):
                        if node.get("Extension", "").lower() == "json" or node.get("PartName") == "/ppt/pmg_execution_manifest.json":
                            root.remove(node)
                    etree.SubElement(root, f"{{{namespace}}}Default", Extension="json", ContentType="application/json")
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                output.writestr(item, data)
            output.writestr("ppt/pmg_execution_manifest.json", json.dumps(compact_manifest, indent=2).encode("utf-8"))
        os.replace(temporary, pptx_path)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)
    if write_companion:
        with open(f"{pptx_path}.manifest.json", "w", encoding="utf-8") as stream:
            json.dump(manifest_dict, stream, indent=2, ensure_ascii=False)


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
        try:
            timeline = extract_expected_timeline(candidate_pptx, hold_duration_ms=hold_duration_ms)
        except Exception:
            timeline = ExpectedTimeline("Unparseable presentation", VERSION, 0)
        timeline.save(os.path.join(qa_dir, "expected_timeline.json"))

        export_mp4 = os.path.join(qa_dir, "powerpoint_export.mp4")
        record_mp4 = os.path.join(qa_dir, "actual_slideshow_recording.mp4")

        if not static_val["is_valid"]:
            report = ReviewReport(status=ValidationStatus.FAIL,
                                  defects=[ReviewDefect(**item) for item in static_val.get("defects", [])],
                                  summary="Static audit failed; repair before opening PowerPoint.")
            return (static_val, timeline, VideoMetadata(export_mp4), VideoMetadata(record_mp4), {}, report,
                    "static_only", "local")

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
            telemetry = dict(w_res.recording_telemetry or {}, worker_job_id=w_res.job_id,
                             clean_open=w_res.clean_open,
                             validated_pptx_sha256=w_res.validated_pptx_sha256)

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
                visual_quality_passed=rev_dict.get("visual_quality_passed", False),
                animation_behavior_passed=rev_dict.get("animation_behavior_passed", False),
                temporal_quality_passed=rev_dict.get("temporal_quality_passed", False),
                capture_quality_passed=rev_dict.get("capture_quality_passed", False),
                object_visibility_passed=rev_dict.get("object_visibility_passed", False),
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

        clean_open = verify_clean_powerpoint_open(candidate_pptx)
        if not clean_open.get("is_clean"):
            report = ReviewReport(
                status=ValidationStatus.BLOCKED if clean_open.get("skipped") else ValidationStatus.FAIL,
                defects=[ReviewDefect("powerpoint_compatibility", "critical", 0, None,
                                      clean_open.get("error", "PowerPoint opening was not clean."))],
                summary="PowerPoint clean-open preflight failed.")
            return (static_val, timeline, VideoMetadata(export_mp4), VideoMetadata(record_mp4),
                    {"clean_open": clean_open}, report, "local_powerpoint_com", "local")

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

        clean_open = verify_clean_powerpoint_open(candidate_pptx)
        telemetry["clean_open"] = clean_open
        telemetry["validated_pptx_sha256"] = file_sha256(candidate_pptx)
        from pmg.validator import inspect_powerpoint_compatibility
        compatibility_defects = inspect_powerpoint_compatibility(static_val, clean_open)
        if compatibility_defects:
            report.defects.extend(compatibility_defects)
            report.status = ValidationStatus.FAIL
            report.summary = "PowerPoint compatibility audit rejected playback evidence."

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
        allow_substitution: bool = True,
    ) -> ExecutionManifest:
        """One repair budget and one final authority for every production path."""
        if isinstance(max_repair_loops, bool) or not isinstance(max_repair_loops, int) or not 0 <= max_repair_loops <= 3:
            raise ValueError("max_repair_loops must be an integer between 0 and 3")
        eff_mode = (worker_mode or os.environ.get("PMG_QA_MODE", "auto")).lower()
        if eff_mode not in ("auto", "local", "remote", "static"):
            raise ValueError(f"Unknown validation worker mode: {eff_mode}")
        run_id, qa_dir = self.create_run_dir(run_id=run_id)
        candidate_pptx = os.path.join(qa_dir, "candidate.pptx")
        if builder_fn is not None:
            builder_fn(candidate_pptx)
        elif existing_pptx_path is not None and os.path.exists(existing_pptx_path):
            shutil.copyfile(existing_pptx_path, candidate_pptx)
        else:
            raise ValueError("Either builder_fn or an existing_pptx_path must be provided.")
        shutil.copyfile(candidate_pptx, candidate_pptx + ".pre_heal.bak")

        repair_history: List[Dict[str, Any]] = []
        loops_executed = 0
        provenance = get_engine_provenance(entrypoint="pmg.orchestrator.PMGOrchestrator.run_pipeline")
        final_verification: Dict[str, Any] = {}
        # This is the only repair orchestration loop, including compatible fallback.
        while True:
            packaging_error = None
            try:
                _embed_manifest_in_pptx(candidate_pptx, {
                    "run_id": run_id, "provenance": provenance,
                    "status": ValidationStatus.NOT_VERIFIED.value,
                    "repair_loops_executed": loops_executed,
                    "validation_method": "remote_windows_worker" if eff_mode == "remote" or (
                        eff_mode == "auto" and not is_powerpoint_available() and (worker_url or os.environ.get("PMG_QA_WORKER_URL"))
                    ) else ("static_only" if eff_mode == "static" or not require_video_playback else "local_powerpoint_com"),
                }, write_companion=False)
            except Exception as exc:
                packaging_error = str(exc)
            checked_hash = file_sha256(candidate_pptx)
            static_val, timeline, export_meta, rec_meta, telemetry, report, val_method, worker_id = (
                self._run_validation_and_playback_pass(
                    candidate_pptx=candidate_pptx, qa_dir=qa_dir,
                    hold_duration_ms=hold_duration_ms, require_video_playback=require_video_playback,
                    worker_mode=eff_mode, worker_url=worker_url, worker_token=worker_token))
            if packaging_error:
                report.defects.append(ReviewDefect("openxml_structure", "critical", 0, None,
                                                  f"Manifest packaging failed: {packaging_error}"))
                report.status = ValidationStatus.FAIL
            if checked_hash != file_sha256(candidate_pptx):
                report.defects.append(ReviewDefect("openxml_structure", "critical", 0, None,
                                                  "PPTX bytes changed during validation."))
                report.status = ValidationStatus.FAIL
            if report.status == ValidationStatus.PASS:
                clean_open = telemetry.get("clean_open", {})
                complete = (static_val.get("is_valid", False) and not report.has_blocking_defects()
                            and all((report.visual_quality_passed, report.animation_behavior_passed,
                                     report.temporal_quality_passed, report.capture_quality_passed,
                                     report.object_visibility_passed))
                            and report.slide_steady_states_verified == timeline.total_slides
                            and timeline.total_slides > 0
                            and export_meta.exists and export_meta.decodable and export_meta.actual_fps > 0
                            and rec_meta.exists and rec_meta.decodable and rec_meta.actual_fps > 0
                            and rec_meta.width >= 1280 and rec_meta.height >= 720
                            and clean_open.get("is_clean") and not clean_open.get("skipped")
                            and clean_open.get("source_unchanged")
                            and clean_open.get("sha256") == checked_hash
                            and telemetry.get("validated_pptx_sha256") == checked_hash)
                if not complete:
                    report.status = ValidationStatus.BLOCKED
                    report.summary = "BLOCKED: complete playback and clean-open evidence bound to the final PPTX is required."
                final_verification = dict(clean_open=clean_open, sha256=checked_hash,
                                          playback_verified=bool(complete))
            if repair_history:
                repair_history[-1]["post_loop_status"] = report.status.value
            if report.status != ValidationStatus.FAIL or loops_executed >= max_repair_loops:
                break
            loops_executed += 1
            try:
                repair_record = execute_repair_loop(
                    pptx_path=candidate_pptx, loop_number=loops_executed,
                    defects=report.defects, static_validation=static_val,
                    allow_substitution=allow_substitution)
            except Exception as exc:
                repair_history.append({"loop_number": loops_executed, "actions_taken": [], "error": str(exc)})
                report.summary = f"FAIL: supported repair could not be applied: {exc}"
                break
            repair_history.append(repair_record)
            # Every repaired byte state is repackaged and fully revalidated above.

        delivered_path = None
        if report.status == ValidationStatus.PASS:
            abs_out = os.path.abspath(output_pptx_path)
            os.makedirs(os.path.dirname(abs_out), exist_ok=True)
            # Stage certified bytes and preserve an existing user's output until
            # the exact destination has also passed its final clean reopen.
            staging_dir = tempfile.mkdtemp(prefix="delivery_", dir=qa_dir)
            staged_output = os.path.join(staging_dir, os.path.basename(abs_out))
            shutil.copyfile(candidate_pptx, staged_output)
            if file_sha256(staged_output) != checked_hash:
                report.status = ValidationStatus.FAIL
                report.summary = "Delivery rejected: final PPTX bytes changed after playback verification."
            if report.status == ValidationStatus.PASS:
                # Use a same-filesystem temporary file for atomic publication.
                descriptor, publish_path = tempfile.mkstemp(suffix=".pptx", dir=os.path.dirname(abs_out))
                os.close(descriptor)
                previous_output = None
                if os.path.isfile(abs_out):
                    previous_output = os.path.join(staging_dir, 'previous_output.pptx')
                    shutil.copyfile(abs_out, previous_output)
                try:
                    shutil.copyfile(staged_output, publish_path)
                    if file_sha256(publish_path) != checked_hash:
                        raise RuntimeError("Certified PPTX changed while staging delivery")
                    os.replace(publish_path, abs_out)
                    if val_method == "local_powerpoint_com":
                        final_open = verify_clean_powerpoint_open(abs_out)
                        final_verification["delivery_clean_open"] = final_open
                        if (not final_open.get("is_clean") or final_open.get("skipped")
                                or file_sha256(abs_out) != checked_hash):
                            report.status = ValidationStatus.FAIL
                            report.summary = "Delivery rejected: final destination failed clean reopen/hash verification."
                            if previous_output:
                                shutil.copyfile(previous_output, publish_path)
                                os.replace(publish_path, abs_out)
                            else:
                                os.remove(abs_out)
                finally:
                    if os.path.exists(publish_path):
                        os.remove(publish_path)
                if report.status == ValidationStatus.PASS:
                    delivered_path = abs_out
                    provenance["final_pptx_sha256"] = checked_hash

        provenance.update(validation_method=val_method, worker_identity=worker_id)
        report.save(os.path.join(qa_dir, "review_report.json"))
        with open(os.path.join(qa_dir, "repair_history.json"), "w", encoding="utf-8") as stream:
            json.dump(repair_history, stream, indent=2, ensure_ascii=False)
        manifest = ExecutionManifest(
            run_id=run_id, provenance=provenance, source_pptx=candidate_pptx,
            delivered_pptx=delivered_path, qa_dir=qa_dir, status=report.status,
            validation_method=val_method, worker_identity=worker_id,
            static_validation=static_val, export_video=export_meta.to_dict(),
            recording_video=rec_meta.to_dict(), review_summary=report.to_dict(),
            repair_loops_executed=loops_executed, repair_history=repair_history,
            final_verification=final_verification)
        manifest.save(os.path.join(qa_dir, "execution_manifest.json"))
        if delivered_path:
            manifest.save(f"{delivered_path}.manifest.json")
        return manifest
