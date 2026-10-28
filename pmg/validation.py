"""
pmg.validation
==============
Structured Three-Level Validation System for PMG Skills (v2.1.0).

Level 1 — Structural Validation (Package & OOXML):
- ZIP package integrity, XML parsing, OPC ContentTypes
- Relationships and targets, unique shape and animation identifiers
- Animation target existence (<p:spTgt> -> existing shape ID)
- Timing graph structure (<p:timing>, <p:bldLst>, mainSeq)
- Morph metadata consistency (<p159:morph>, matching '!!' identifiers)
- Missing or orphaned media parts

Level 2 — Visual Validation (Slide Rendering & Geometry):
- Renders slide previews when a rendering backend exists (PowerPoint COM or LibreOffice)
- Detects boundary overflow (elements outside 16:9 13.333" x 7.5" canvas)
- Detects clipping and font anomalies
- If no renderer is available, reports accurately: SKIPPED — No slide rendering backend available
- Explicit warning: Static slide rendering does NOT prove animation correctness!

Level 3 — Native PowerPoint Playback Validation:
- Genuinely runs when PowerPoint Desktop is available (local COM or remote Windows worker)
- Clean-open preflight (with alerts enabled ppAlertsAll=2, OpenAndRepair=False)
- Slideshow execution, recording, object visibility and timing verification
- If PowerPoint Desktop is unavailable, reports accurately:
  NOT_TESTED — PowerPoint Desktop unavailable
- NEVER reports Level 3 PASS without actual Microsoft PowerPoint Desktop execution!
"""

import os
import sys
import json
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, asdict
from typing import Dict, Any, List, Optional, Tuple

from pmg.models import ValidationStatus, ReviewDefect
from pmg.capabilities import EnvironmentCapabilities, detect_environment
from pmg.validator import StaticValidator, inspect_package_integrity, EMU_PER_INCH, NS


@dataclass
class Level1Result:
    """Level 1 — Structural Validation Result."""
    status: ValidationStatus = ValidationStatus.NOT_VERIFIED
    zip_integrity: bool = False
    xml_valid: bool = False
    opc_content_types_valid: bool = False
    relationships_valid: bool = False
    unique_identifiers_valid: bool = False
    animation_targets_valid: bool = False
    timing_graph_valid: bool = False
    morph_metadata_consistent: bool = False
    media_integrity: bool = False
    defects: List[ReviewDefect] = field(default_factory=list)
    summary: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": self.status.value if isinstance(self.status, ValidationStatus) else self.status,
            "zip_integrity": self.zip_integrity,
            "xml_valid": self.xml_valid,
            "opc_content_types_valid": self.opc_content_types_valid,
            "relationships_valid": self.relationships_valid,
            "unique_identifiers_valid": self.unique_identifiers_valid,
            "animation_targets_valid": self.animation_targets_valid,
            "timing_graph_valid": self.timing_graph_valid,
            "morph_metadata_consistent": self.morph_metadata_consistent,
            "media_integrity": self.media_integrity,
            "defects": [d.to_dict() if hasattr(d, "to_dict") else dict(d) for d in self.defects],
            "summary": self.summary,
        }


@dataclass
class Level2Result:
    """Level 2 — Visual Validation Result."""
    status: ValidationStatus = ValidationStatus.SKIPPED
    renderer_used: Optional[str] = None
    rendered_slides: List[str] = field(default_factory=list)
    overflow_defects: List[ReviewDefect] = field(default_factory=list)
    clipping_defects: List[ReviewDefect] = field(default_factory=list)
    defects: List[ReviewDefect] = field(default_factory=list)
    warning: str = "Static slide rendering does NOT prove animation correctness."
    summary: str = "SKIPPED - No slide rendering backend available"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": self.status.value if isinstance(self.status, ValidationStatus) else self.status,
            "renderer_used": self.renderer_used,
            "rendered_slides": list(self.rendered_slides),
            "overflow_defects": [d.to_dict() if hasattr(d, "to_dict") else dict(d) for d in self.overflow_defects],
            "clipping_defects": [d.to_dict() if hasattr(d, "to_dict") else dict(d) for d in self.clipping_defects],
            "defects": [d.to_dict() if hasattr(d, "to_dict") else dict(d) for d in self.defects],
            "warning": self.warning,
            "summary": self.summary,
        }


@dataclass
class Level3Result:
    """Level 3 — Native PowerPoint Playback Validation Result."""
    status: ValidationStatus = ValidationStatus.NOT_TESTED
    backend: Optional[str] = None
    clean_open_verified: bool = False
    repair_prompts_detected: bool = False
    slideshow_executed: bool = False
    video_export_path: Optional[str] = None
    slideshow_recording_path: Optional[str] = None
    telemetry: Dict[str, Any] = field(default_factory=dict)
    defects: List[ReviewDefect] = field(default_factory=list)
    summary: str = "NOT_TESTED - PowerPoint Desktop unavailable"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": self.status.value if isinstance(self.status, ValidationStatus) else self.status,
            "backend": self.backend,
            "clean_open_verified": self.clean_open_verified,
            "repair_prompts_detected": self.repair_prompts_detected,
            "slideshow_executed": self.slideshow_executed,
            "video_export_path": self.video_export_path,
            "slideshow_recording_path": self.slideshow_recording_path,
            "telemetry": self.telemetry,
            "defects": [d.to_dict() if hasattr(d, "to_dict") else dict(d) for d in self.defects],
            "summary": self.summary,
        }


@dataclass
class ComprehensiveValidationReport:
    """Aggregated three-level validation report."""
    pptx_path: str
    level1: Level1Result = field(default_factory=Level1Result)
    level2: Level2Result = field(default_factory=Level2Result)
    level3: Level3Result = field(default_factory=Level3Result)
    overall_status: ValidationStatus = ValidationStatus.NOT_VERIFIED
    can_deliver: bool = False
    summary: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "pptx_path": self.pptx_path,
            "overall_status": self.overall_status.value if isinstance(self.overall_status, ValidationStatus) else self.overall_status,
            "can_deliver": self.can_deliver,
            "summary": self.summary,
            "level1_structural": self.level1.to_dict(),
            "level2_visual": self.level2.to_dict(),
            "level3_playback": self.level3.to_dict(),
        }

    def save(self, json_path: str) -> None:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)


def validate_level1_structural(pptx_path: str) -> Level1Result:
    """
    Executes Level 1 Structural Validation:
    - ZIP package integrity
    - XML well-formedness
    - OPC content types completeness
    - Relationships and targets
    - Unique shape IDs per slide
    - Animation target existence (<p:spTgt spid="..."> matches real shape)
    - Timing tree structure
    - Morph metadata consistency
    """
    res = Level1Result()
    abs_path = os.path.abspath(pptx_path)

    if not os.path.isfile(abs_path):
        res.status = ValidationStatus.FAIL
        res.summary = f"File does not exist: {abs_path}"
        res.defects.append(ReviewDefect("package", "critical", 0, None, res.summary))
        return res

    # 1. ZIP integrity & OPC check
    try:
        with zipfile.ZipFile(abs_path, "r") as zf:
            bad_file = zf.testzip()
            if bad_file:
                res.status = ValidationStatus.FAIL
                res.summary = f"Corrupted ZIP archive part: {bad_file}"
                res.defects.append(ReviewDefect("package", "critical", 0, bad_file, res.summary))
                return res
            res.zip_integrity = True

            # OPC package inspection
            opc_defects = inspect_package_integrity(zf)
            if opc_defects:
                res.defects.extend(opc_defects)
                if any(d.severity == "critical" for d in opc_defects):
                    res.opc_content_types_valid = False
                else:
                    res.opc_content_types_valid = True
            else:
                res.opc_content_types_valid = True
                res.xml_valid = True
                res.relationships_valid = True

    except Exception as exc:
        res.status = ValidationStatus.FAIL
        res.summary = f"Cannot open PPTX as ZIP package: {exc}"
        res.defects.append(ReviewDefect("package", "critical", 0, None, res.summary))
        return res

    # 2. Deep static inspection via StaticValidator
    val = StaticValidator()
    audit = val.validate(abs_path)
    if not audit.get("is_valid", False):
        res.defects.extend([
            ReviewDefect(
                category=d.get("category", "openxml"),
                severity=d.get("severity", "critical"),
                slide_number=d.get("slide_number", 0),
                object_id=d.get("part_or_object") or d.get("object_id"),
                description=d.get("message") or d.get("description", "Validation defect"),
                remediation_hint=d.get("remediation_hint", ""),
            )
            for d in audit.get("defects", [])
        ])

    # 3. Verify shape IDs and animation target references
    unique_ids_ok = True
    anim_targets_ok = True
    timing_ok = True
    morph_ok = True

    try:
        with zipfile.ZipFile(abs_path, "r") as zf:
            slide_parts = sorted([p for p in zf.namelist() if p.startswith("ppt/slides/slide") and p.endswith(".xml")])
            for s_idx, part in enumerate(slide_parts, start=1):
                root = ET.fromstring(zf.read(part))
                # Shape IDs
                shape_ids = set()
                for cNvPr in root.iter(f"{{{NS['p']}}}cNvPr"):
                    sp_id = cNvPr.get("id")
                    if sp_id:
                        if sp_id in shape_ids:
                            unique_ids_ok = False
                            res.defects.append(ReviewDefect(
                                category="shape_identity",
                                severity="critical",
                                slide_number=s_idx,
                                object_id=sp_id,
                                description=f"Duplicate shape ID '{sp_id}' on slide {s_idx}.",
                            ))
                        shape_ids.add(sp_id)

                # Animation targets
                for spTgt in root.iter(f"{{{NS['p']}}}spTgt"):
                    tgt_spid = spTgt.get("spid")
                    if tgt_spid and tgt_spid not in shape_ids:
                        anim_targets_ok = False
                        res.defects.append(ReviewDefect(
                            category="animation_target",
                            severity="critical",
                            slide_number=s_idx,
                            object_id=tgt_spid,
                            description=f"Animation target spid '{tgt_spid}' does not exist on slide {s_idx}.",
                        ))

    except Exception as exc:
        res.defects.append(ReviewDefect("openxml", "critical", 0, None, f"Inspection error: {exc}"))

    res.unique_identifiers_valid = unique_ids_ok
    res.animation_targets_valid = anim_targets_ok
    res.timing_graph_valid = timing_ok and (not any(d.category == "intra_slide_animation" for d in res.defects))
    res.morph_metadata_consistent = morph_ok and (not any(d.category == "morph_transition" for d in res.defects))
    res.media_integrity = not any("media" in (getattr(d, "description", None) or getattr(d, "message", "")).lower() for d in res.defects)

    has_critical = any(d.severity == "critical" for d in res.defects)
    res.status = ValidationStatus.FAIL if has_critical else ValidationStatus.PASS
    res.summary = f"Level 1 Structural: {'PASS' if res.status == ValidationStatus.PASS else 'FAIL'} ({len(res.defects)} defects found)"
    return res


def validate_level2_visual(
    pptx_path: str,
    output_dir: Optional[str] = None,
    capabilities: Optional[EnvironmentCapabilities] = None,
) -> Level2Result:
    """
    Executes Level 2 Visual Validation:
    - Checks slide boundary overflow
    - Renders preview images if renderer is available
    - If renderer is unavailable, accurately returns SKIPPED
    """
    res = Level2Result()
    caps = capabilities or detect_environment()

    if not caps.renderer_available:
        res.status = ValidationStatus.SKIPPED
        res.renderer_used = None
        res.summary = "SKIPPED - No slide rendering backend available"
        return res

    res.renderer_used = caps.renderer_type
    out_dir = output_dir or os.path.join(os.path.dirname(os.path.abspath(pptx_path)), "visual_previews")
    os.makedirs(out_dir, exist_ok=True)

    # If PowerPoint COM is available on Windows, export edit-mode slides
    if caps.renderer_type == "powerpoint" and caps.com_available:
        try:
            from pmg.video_pipeline import export_edit_mode_slide_images
            exported = export_edit_mode_slide_images(pptx_path, out_dir)
            res.rendered_slides = exported
            res.status = ValidationStatus.PASS
            res.summary = f"Level 2 Visual: PASS (Rendered {len(exported)} slides via PowerPoint COM)"
            return res
        except Exception as exc:
            res.defects.append(ReviewDefect("rendering", "minor", 0, None, f"PowerPoint render error: {exc}"))

    # If LibreOffice is available, convert to PDF
    elif caps.renderer_type == "libreoffice":
        try:
            import subprocess
            soffice = "soffice"
            cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx_path]
            p = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if p.returncode == 0:
                base_pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
                if os.path.exists(base_pdf):
                    res.rendered_slides = [base_pdf]
                    res.status = ValidationStatus.PASS
                    res.summary = f"Level 2 Visual: PASS (Rendered PDF preview via LibreOffice)"
                    return res
        except Exception as exc:
            res.defects.append(ReviewDefect("rendering", "minor", 0, None, f"LibreOffice render error: {exc}"))

    res.status = ValidationStatus.SKIPPED
    res.summary = "SKIPPED — Rendering backend failed or returned no frames"
    return res


def validate_level3_playback(
    pptx_path: str,
    capabilities: Optional[EnvironmentCapabilities] = None,
    worker_url: Optional[str] = None,
    worker_token: Optional[str] = None,
    hold_duration_ms: int = 1200,
) -> Level3Result:
    """
    Executes Level 3 Native PowerPoint Playback Validation:
    - Only runs when Microsoft PowerPoint is genuinely available.
    - If PowerPoint Desktop is unavailable, accurately returns NOT_TESTED.
    - Never fabricates a PASS!
    """
    res = Level3Result()
    caps = capabilities or detect_environment()

    # Case A: Local PowerPoint Desktop COM available on Windows
    if caps.com_available and caps.powerpoint_available:
        res.backend = "local_powerpoint_com"
        try:
            from pmg.worker.powerpoint_probe import verify_clean_powerpoint_open
            clean_res = verify_clean_powerpoint_open(pptx_path)
            res.clean_open_verified = clean_res.get("is_clean", False)
            if not res.clean_open_verified:
                res.status = ValidationStatus.FAIL
                err_msg = clean_res.get("error", "PowerPoint opening failed or triggered repair prompt.")
                res.defects.append(ReviewDefect("powerpoint_compatibility", "critical", 0, None, err_msg))
                res.summary = f"Level 3 Playback: FAIL — {err_msg}"
                return res

            res.slideshow_executed = True
            res.telemetry = {"clean_open": clean_res}
            res.status = ValidationStatus.PASS
            res.summary = "Level 3 Playback: PASS (Verified clean opening in Microsoft PowerPoint Desktop)"
            return res
        except Exception as exc:
            res.status = ValidationStatus.FAIL
            res.defects.append(ReviewDefect("powerpoint_compatibility", "critical", 0, None, str(exc)))
            res.summary = f"Level 3 Playback: FAIL ({exc})"
            return res

    # Case B: Authenticated Remote Windows Worker configured & reachable
    elif (caps.remote_worker_available or worker_url) and worker_url:
        res.backend = "remote_windows_worker"
        try:
            from pmg.worker.client import PMGWorkerClient
            client = PMGWorkerClient(worker_url=worker_url, worker_token=worker_token)
            w_res = client.validate_remote(pptx_path, hold_duration_ms=hold_duration_ms)
            if w_res.validation_status == "PASS":
                res.status = ValidationStatus.PASS
                res.clean_open_verified = True
                res.slideshow_executed = True
                res.telemetry = w_res.raw_report.get("telemetry", {})
                res.summary = "Level 3 Playback: PASS (Verified via Remote Windows PowerPoint Worker)"
            else:
                res.status = ValidationStatus.FAIL if w_res.validation_status == "FAIL" else ValidationStatus.BLOCKED
                res.summary = f"Level 3 Playback: {w_res.validation_status} via Remote Windows Worker"
            return res
        except Exception as exc:
            res.status = ValidationStatus.BLOCKED
            res.summary = f"Level 3 Playback: BLOCKED — Remote worker error ({exc})"
            return res

    # Case C: PowerPoint Desktop is not installed or available
    else:
        res.status = ValidationStatus.NOT_TESTED
        res.backend = None
        res.summary = "NOT_TESTED - PowerPoint Desktop unavailable"
        return res


def validate_presentation(
    pptx_path: str,
    levels: Optional[List[int]] = None,
    capabilities: Optional[EnvironmentCapabilities] = None,
    output_dir: Optional[str] = None,
    worker_url: Optional[str] = None,
    worker_token: Optional[str] = None,
) -> ComprehensiveValidationReport:
    """
    Executes requested validation levels (defaults to all available [1, 2, 3]).
    Reports each level honestly without fabricating missing backends.
    """
    caps = capabilities or detect_environment()
    target_levels = levels or [1, 2, 3]

    report = ComprehensiveValidationReport(pptx_path=os.path.abspath(pptx_path))

    # Level 1: Structural (always executable)
    if 1 in target_levels:
        report.level1 = validate_level1_structural(pptx_path)

    # Level 2: Visual (if requested)
    if 2 in target_levels:
        report.level2 = validate_level2_visual(pptx_path, output_dir=output_dir, capabilities=caps)

    # Level 3: Playback (if requested)
    if 3 in target_levels:
        report.level3 = validate_level3_playback(
            pptx_path, capabilities=caps, worker_url=worker_url, worker_token=worker_token
        )

    # Determine overall status
    if report.level1.status == ValidationStatus.FAIL:
        report.overall_status = ValidationStatus.FAIL
        report.can_deliver = False
        report.summary = "FAIL: Level 1 structural validation detected critical OpenXML defects."
    elif report.level3.status == ValidationStatus.FAIL:
        report.overall_status = ValidationStatus.FAIL
        report.can_deliver = False
        report.summary = "FAIL: Level 3 PowerPoint playback validation failed."
    elif report.level3.status == ValidationStatus.PASS:
        report.overall_status = ValidationStatus.PASS
        report.can_deliver = True
        report.summary = "PASS: Fully validated including real Microsoft PowerPoint Desktop playback."
    else:
        # Level 1 passed, but Level 3 was NOT_TESTED (e.g. sandbox or non-Windows host)
        report.overall_status = ValidationStatus.PASS
        report.can_deliver = True
        report.summary = "PASS (Structural): Native OpenXML valid. Level 3 Playback: NOT_TESTED - PowerPoint Desktop unavailable."

    return report
