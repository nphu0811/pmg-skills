"""
pmg.diagnostics
===============
Structured Diagnostic Engine & Error Classification for PMG v2.4.0.

Provides standardized error codes, severity ratings, evidence tracking,
suggested remediations, and repair safety classifications:
- SAFE_AUTOMATIC: Well-understood, deterministic, reversible repairs (no visual/semantic loss).
- REQUIRES_REVIEW: Repairs that alter transitions, layouts, or visual appearance.
- UNSUPPORTED: Irreparable defects that require user intervention or specification redesign.
"""

from __future__ import annotations
from enum import Enum
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional

from pmg.models import ReviewDefect


class ErrorCode(str, Enum):
    """Standardized error codes for PPTX package and playback diagnostics."""
    PKG_INVALID_ZIP = "PKG_INVALID_ZIP"
    XML_PARSE_FAILURE = "XML_PARSE_FAILURE"
    REL_BROKEN_TARGET = "REL_BROKEN_TARGET"
    OPC_MISSING_CONTENT_TYPE = "OPC_MISSING_CONTENT_TYPE"
    SHP_DUPLICATE_ID = "SHP_DUPLICATE_ID"
    ANIM_MISSING_TARGET = "ANIM_MISSING_TARGET"
    TIMING_INVALID_REFERENCE = "TIMING_INVALID_REFERENCE"
    MORPH_IDENTITY_COLLISION = "MORPH_IDENTITY_COLLISION"
    MEDIA_MISSING_RESOURCE = "MEDIA_MISSING_RESOURCE"
    ANIM_UNSUPPORTED_FEATURE = "ANIM_UNSUPPORTED_FEATURE"
    RENDER_FAILURE = "RENDER_FAILURE"
    PPT_OPEN_FAILURE = "PPT_OPEN_FAILURE"
    PPT_PLAYBACK_FAILURE = "PPT_PLAYBACK_FAILURE"
    UNKNOWN_FAILURE = "UNKNOWN_FAILURE"


class RepairSafety(str, Enum):
    """Safety classification governing automated healing."""
    SAFE_AUTOMATIC = "SAFE_AUTOMATIC"      # Completely safe to apply without user intervention
    REQUIRES_REVIEW = "REQUIRES_REVIEW"    # Changes visible motion, transitions, or geometry
    UNSUPPORTED = "UNSUPPORTED"            # Unambiguous failure; cannot be safely auto-healed


@dataclass
class StructuredDiagnostic:
    """Detailed machine-readable diagnostic record."""
    error_code: ErrorCode
    severity: str  # "critical", "major", "warning", "info"
    affected_component: str
    slide_number: int = 0
    object_id: Optional[str] = None
    evidence: str = ""
    suggested_repair: str = ""
    repair_safety: RepairSafety = RepairSafety.SAFE_AUTOMATIC

    def to_dict(self) -> Dict[str, Any]:
        return {
            "error_code": self.error_code.value,
            "severity": self.severity,
            "affected_component": self.affected_component,
            "slide_number": self.slide_number,
            "object_id": self.object_id,
            "evidence": self.evidence,
            "suggested_repair": self.suggested_repair,
            "repair_safety": self.repair_safety.value,
        }


class DiagnosticClassifier:
    """Classifies raw validation defects into structured diagnostics with safety ratings."""

    @classmethod
    def classify_defect(cls, defect: ReviewDefect) -> StructuredDiagnostic:
        desc = (defect.description or "").lower()
        cat = (defect.category or "").lower()
        remediation = defect.remediation_hint or ""

        # 1. OPC ContentType missing
        if "has no registered contenttype" in desc or "contenttype" in desc or remediation == "declare_part_content_type":
            return StructuredDiagnostic(
                error_code=ErrorCode.OPC_MISSING_CONTENT_TYPE,
                severity=defect.severity,
                affected_component=defect.object_id or "[Content_Types].xml",
                slide_number=defect.slide_number,
                object_id=defect.object_id,
                evidence=defect.description,
                suggested_repair="repair_missing_content_types",
                repair_safety=RepairSafety.SAFE_AUTOMATIC,
            )

        # 2. Morph identity collision
        if cat == "morph" or "morph collision" in desc or "traveling entity" in desc:
            return StructuredDiagnostic(
                error_code=ErrorCode.MORPH_IDENTITY_COLLISION,
                severity=defect.severity,
                affected_component=defect.object_id or f"Slide_{defect.slide_number}",
                slide_number=defect.slide_number,
                object_id=defect.object_id,
                evidence=defect.description,
                suggested_repair="disambiguate_morph_identifiers",
                repair_safety=RepairSafety.REQUIRES_REVIEW,
            )

        # 3. Duplicate shape identifier
        if "duplicate shape" in desc or "duplicate" in desc and ("shape id" in desc or "spid" in desc or cat == "shape_identity"):
            return StructuredDiagnostic(
                error_code=ErrorCode.SHP_DUPLICATE_ID,
                severity=defect.severity,
                affected_component=defect.object_id or f"Slide_{defect.slide_number}",
                slide_number=defect.slide_number,
                object_id=defect.object_id,
                evidence=defect.description,
                suggested_repair="repair_duplicate_shape_ids",
                repair_safety=RepairSafety.SAFE_AUTOMATIC,
            )

        # 3. Broken animation target
        if "animation target" in desc or "targets non-existent" in desc or "sptgt" in desc or remediation in ("repair_animation_target", "rebuild_timing_tree"):
            return StructuredDiagnostic(
                error_code=ErrorCode.ANIM_MISSING_TARGET,
                severity=defect.severity,
                affected_component=defect.object_id or f"Slide_{defect.slide_number}",
                slide_number=defect.slide_number,
                object_id=defect.object_id,
                evidence=defect.description,
                suggested_repair="repair_broken_animation_targets",
                repair_safety=RepairSafety.SAFE_AUTOMATIC,
            )

        # 4. Morph identity collision
        if "morph collision" in desc or "duplicate traveling entity" in desc:
            return StructuredDiagnostic(
                error_code=ErrorCode.MORPH_IDENTITY_COLLISION,
                severity=defect.severity,
                affected_component=defect.object_id or f"Slide_{defect.slide_number}",
                slide_number=defect.slide_number,
                object_id=defect.object_id,
                evidence=defect.description,
                suggested_repair="disambiguate_morph_identifiers",
                repair_safety=RepairSafety.REQUIRES_REVIEW,
            )

        # 5. Broken relationships
        if "relationship" in desc or "references missing" in desc:
            return StructuredDiagnostic(
                error_code=ErrorCode.REL_BROKEN_TARGET,
                severity=defect.severity,
                affected_component=defect.object_id or "relationships",
                slide_number=defect.slide_number,
                object_id=defect.object_id,
                evidence=defect.description,
                suggested_repair="rebuild_relationships",
                repair_safety=RepairSafety.SAFE_AUTOMATIC,
            )

        # 6. XML Parse failure
        if "xml" in desc and ("parse" in desc or "syntax" in desc or "invalid" in desc):
            return StructuredDiagnostic(
                error_code=ErrorCode.XML_PARSE_FAILURE,
                severity="critical",
                affected_component=defect.object_id or "xml_parts",
                slide_number=defect.slide_number,
                object_id=defect.object_id,
                evidence=defect.description,
                suggested_repair="regenerate_corrupted_xml",
                repair_safety=RepairSafety.REQUIRES_REVIEW,
            )

        # 7. ZIP package integrity
        if "zip" in desc or "package" in desc or "not a valid zip" in desc:
            return StructuredDiagnostic(
                error_code=ErrorCode.PKG_INVALID_ZIP,
                severity="critical",
                affected_component=defect.object_id or "package_zip",
                slide_number=defect.slide_number,
                object_id=defect.object_id,
                evidence=defect.description,
                suggested_repair="rebuild_package_archive",
                repair_safety=RepairSafety.UNSUPPORTED,
            )

        # 8. PowerPoint opening or playback failure
        if "clean_open" in desc or "repair presentation" in desc or "powerpoint" in desc:
            return StructuredDiagnostic(
                error_code=ErrorCode.PPT_OPEN_FAILURE,
                severity=defect.severity,
                affected_component="powerpoint_host",
                slide_number=defect.slide_number,
                object_id=defect.object_id,
                evidence=defect.description,
                suggested_repair="clean_open_fallback",
                repair_safety=RepairSafety.REQUIRES_REVIEW,
            )

        # Default fallback
        return StructuredDiagnostic(
            error_code=ErrorCode.UNKNOWN_FAILURE,
            severity=defect.severity,
            affected_component=defect.object_id or "unknown",
            slide_number=defect.slide_number,
            object_id=defect.object_id,
            evidence=defect.description,
            suggested_repair=remediation or "investigate_defect",
            repair_safety=RepairSafety.REQUIRES_REVIEW,
        )

    @classmethod
    def classify_all(cls, defects: List[ReviewDefect]) -> List[StructuredDiagnostic]:
        return [cls.classify_defect(d) for d in defects]
