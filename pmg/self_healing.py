"""
pmg.self_healing
================
Unified 3-Loop Self-Healing Engine extending scripts.motion_validator to repair
both static OpenXML deficiencies and visual/temporal playback defects detected
during real PowerPoint slideshow review.
"""

import os
import shutil
from typing import List, Dict, Any, Tuple
import pptx
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

from scripts.motion_validator import (
    repair_xml_transitions,
    repair_traveling_anchor_gap,
    substitute_with_fallback_skill,
)
from pmg.models import ReviewDefect


def _interpolate_anchor_across_gaps(pptx_path: str, anchor_tag: str, gap_slides: List[int]) -> List[str]:
    """
    Structural repair (Loop 2): Finds the nearest preceding and succeeding slides
    containing `anchor_tag` and injects an interpolated native shape on `gap_slides`
    so Morph continuity is preserved across the entire trajectory.
    """
    actions = []
    prs = pptx.Presentation(pptx_path)
    total = len(prs.slides)

    # Map existing occurrences of anchor_tag
    found_coords: Dict[int, Tuple[int, int, int, int]] = {}
    for idx, slide in enumerate(prs.slides, 1):
        for shp in slide.shapes:
            if shp.name == anchor_tag:
                found_coords[idx] = (shp.left, shp.top, shp.width, shp.height)
                break

    if not found_coords:
        repair_traveling_anchor_gap(pptx_path, anchor_tag, gap_slides)
        actions.append(f"Bridged {anchor_tag} on slides {gap_slides} via proxy anchor.")
        return actions

    sorted_present = sorted(found_coords.keys())
    modified = False

    for g_idx in gap_slides:
        if not (1 <= g_idx <= total):
            continue
        slide = prs.slides[g_idx - 1]
        if any(s.name == anchor_tag for s in slide.shapes):
            continue

        prev_slides = [s for s in sorted_present if s < g_idx]
        next_slides = [s for s in sorted_present if s > g_idx]
        s_prev = prev_slides[-1] if prev_slides else sorted_present[0]
        s_next = next_slides[0] if next_slides else sorted_present[-1]

        l1, t1, w1, h1 = found_coords[s_prev]
        l2, t2, w2, h2 = found_coords[s_next]
        span = max(1, s_next - s_prev)
        alpha = (g_idx - s_prev) / span

        interp_l = int(l1 + (l2 - l1) * alpha)
        interp_t = int(t1 + (t2 - t1) * alpha)
        interp_w = max(int(Inches(0.4)), int(w1 + (w2 - w1) * alpha))
        interp_h = max(int(Inches(0.4)), int(h1 + (h2 - h1) * alpha))

        shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, interp_l, interp_t, interp_w, interp_h)
        shp.name = anchor_tag
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor(2, 132, 199)
        shp.line.color.rgb = RGBColor(255, 255, 255)
        modified = True
        actions.append(
            f"Interpolated '{anchor_tag}' on slide {g_idx} between slides {s_prev} and {s_next}."
        )

    if modified:
        prs.save(pptx_path)
    return actions


def execute_repair_loop(
    pptx_path: str,
    loop_number: int,
    defects: List[ReviewDefect],
    static_validation: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Executes one self-healing loop according to the 3-Loop Repair Architecture:
    - Loop 1: Targeted repair (missing transitions, invalid durations, direct targets)
    - Loop 2: Structural repair (Morph anchor interpolation, geometry & hierarchy healing)
    - Loop 3: Compatible fallback repair (fallback native transition / effect substitution)
    """
    actions_taken: List[str] = []
    missing_trans = static_validation.get("missing_transitions", [])
    discontinuous = static_validation.get("discontinuous_anchors", {})

    # Also collect any slide numbers flagged in visual/animation/temporal defects
    defect_slides = sorted(
        list({d.slide_number for d in defects if d.slide_number > 1 and d.severity in ("critical", "major")})
    )

    if loop_number == 1:
        # LOOP 1: TARGETED REPAIR
        if missing_trans:
            repair_xml_transitions(pptx_path, missing_trans, trans_type="morph", dur_ms=1600)
            actions_taken.append(f"Loop 1 (Targeted): Injected Morph transitions (1600ms) on slides {missing_trans}.")

        # Check for abnormal transition durations
        dur_slides = [
            d.slide_number
            for d in defects
            if d.remediation_hint == "normalize_duration" and d.slide_number > 1
        ]
        if dur_slides:
            repair_xml_transitions(pptx_path, dur_slides, trans_type="morph", dur_ms=1600)
            actions_taken.append(f"Loop 1 (Targeted): Normalized transition durations to 1600ms on slides {dur_slides}.")

        if not actions_taken and defect_slides:
            repair_xml_transitions(pptx_path, defect_slides, trans_type="morph", dur_ms=1600)
            actions_taken.append(f"Loop 1 (Targeted): Refreshed OpenXML transition timing on slides {defect_slides}.")

    elif loop_number == 2:
        # LOOP 2: STRUCTURAL REPAIR
        if missing_trans:
            repair_xml_transitions(pptx_path, missing_trans, trans_type="morph", dur_ms=1600)
            actions_taken.append(f"Loop 2 (Structural): Re-injected Morph transitions on slides {missing_trans}.")

        if discontinuous:
            for tag, gap_slides in discontinuous.items():
                 acts = _interpolate_anchor_across_gaps(pptx_path, tag, gap_slides)
                 actions_taken.extend([f"Loop 2 (Structural): {a}" for a in acts])

        if not actions_taken and defect_slides:
            repair_xml_transitions(pptx_path, defect_slides, trans_type="morph", dur_ms=1600)
            actions_taken.append(f"Loop 2 (Structural): Rebuilt transition XML nodes on slides {defect_slides}.")

    else:
        # LOOP 3: COMPATIBLE FALLBACK REPAIR
        target_fallback_slides = sorted(list(set(missing_trans + defect_slides)))
        if not target_fallback_slides:
            prs = pptx.Presentation(pptx_path)
            target_fallback_slides = list(range(2, len(prs.slides) + 1))

        if discontinuous:
            for tag, gap_slides in discontinuous.items():
                repair_traveling_anchor_gap(pptx_path, tag, gap_slides)
                actions_taken.append(f"Loop 3 (Fallback): Bridged anchor '{tag}' on slides {gap_slides}.")

        if target_fallback_slides:
            sub_skill = substitute_with_fallback_skill(
                pptx_path,
                failed_slides=target_fallback_slides,
                reason="Fallback to compatible native Circle Wipe / Spotlight transition on loop 3.",
            )
            actions_taken.append(
                f"Loop 3 (Fallback): Activated '{sub_skill}' native transition fallback on slides {target_fallback_slides}."
            )

    return {
        "loop_number": loop_number,
        "loop_type": (
            "targeted_repair"
            if loop_number == 1
            else ("structural_repair" if loop_number == 2 else "compatible_fallback_repair")
        ),
        "actions_taken": actions_taken,
    }

