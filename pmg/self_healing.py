"""
pmg.self_healing
================
Unified 3-Loop Self-Healing Engine extending scripts.motion_validator to repair
both static OpenXML deficiencies and visual/temporal playback defects detected
during real PowerPoint slideshow review (including intra-slide <p:timing> animations).
"""

import os
import shutil
from typing import List, Dict, Any, Tuple
import pptx
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from lxml import etree

from scripts.motion_validator import (
    repair_xml_transitions,
    repair_traveling_anchor_gap,
    substitute_with_fallback_skill,
)
from pmg.models import ReviewDefect
from pmg.animations import P_NS, NSMAP, SlideAnimationTimeline


def _repair_slide_timing_targets_and_paths(pptx_path: str, loop_number: int) -> List[str]:
    """
    Repairs broken <p:spTgt> references, invalid <p:animMotion> path syntax,
    or missing mainSeq structures in intra-slide <p:timing> trees.
    """
    actions: List[str] = []
    prs = pptx.Presentation(pptx_path)
    modified = False

    for idx, slide in enumerate(prs.slides, 1):
        valid_spids = [str(s.shape_id) for s in slide.shapes]
        if not valid_spids:
            continue
        fallback_spid = valid_spids[-1]

        sld = slide._element
        timing = sld.find(f"{{{P_NS}}}timing")
        if timing is None:
            continue

        # Check if mainSeq is missing (Loop 2 structural rebuild)
        main_seq = timing.xpath('.//p:cTn[@nodeType="mainSeq"]')
        if not main_seq and loop_number >= 2:
            tl = SlideAnimationTimeline(slide_number=idx)
            for s in slide.shapes[:4]:
                tl.add_effect(s, "fade", trigger="with_previous", duration_ms=600)
            tl.compile_to_openxml(slide)
            modified = True
            actions.append(f"Rebuilt structural <p:timing> mainSeq tree on slide {idx}.")
            continue

        # Repair invalid spTgt references
        for sp_tgt in timing.xpath(".//p:spTgt"):
            spid = sp_tgt.get("spid", "")
            if spid not in valid_spids:
                sp_tgt.set("spid", fallback_spid)
                modified = True
                actions.append(f"Rebound broken animation target spid={spid} -> spid={fallback_spid} on slide {idx}.")

        # Repair invalid or fallback motion paths
        for am in timing.xpath(".//p:animMotion"):
            path_val = (am.get("path") or "").strip()
            if not path_val or not path_val.upper().startswith("M"):
                am.set("path", "M 0 0 C 0.08 -0.1 0.18 0.1 0.25 0 E" if loop_number < 3 else "M 0 0 L 0.25 0 E")
                modified = True
                actions.append(f"Repaired invalid motion path syntax on slide {idx}.")
            elif loop_number == 3 and "C" in path_val and len(path_val.split()) < 8:
                am.set("path", "M 0 0 L 0.25 0 E")
                modified = True
                actions.append(f"Substituted malformed Bezier motion path with compatible linear native path on slide {idx}.")

    if modified:
        prs.save(pptx_path)
    return actions


def _interpolate_anchor_across_gaps(pptx_path: str, anchor_tag: str, gap_slides: List[int]) -> List[str]:
    """
    Structural repair (Loop 2): Finds the nearest preceding and succeeding slides
    containing `anchor_tag` and injects an interpolated native shape on `gap_slides`
    so Morph continuity is preserved across the entire trajectory.
    """
    actions = []
    prs = pptx.Presentation(pptx_path)
    total = len(prs.slides)

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
    - Loop 1: Targeted repair (missing transitions, invalid durations, direct animation targets/paths)
    - Loop 2: Structural repair (Morph anchor interpolation, <p:timing> tree rebuild, geometry healing)
    - Loop 3: Compatible fallback repair (fallback native transition / simpler native motion path substitution)
    """
    actions_taken: List[str] = []
    missing_trans = static_validation.get("missing_transitions", [])
    discontinuous = static_validation.get("discontinuous_anchors", {})

    defect_slides = sorted(
        list({d.slide_number for d in defects if d.slide_number > 1 and d.severity in ("critical", "major")})
    )

    # Check for intra-slide <p:timing> defects
    timing_hints = {"repair_animation_target", "repair_motion_path", "rebuild_timing_tree"}
    if any(d.remediation_hint in timing_hints for d in defects):
        t_acts = _repair_slide_timing_targets_and_paths(pptx_path, loop_number)
        actions_taken.extend([f"Loop {loop_number}: {a}" for a in t_acts])

    if loop_number == 1:
        # LOOP 1: TARGETED REPAIR
        if missing_trans:
            repair_xml_transitions(pptx_path, missing_trans, trans_type="morph", dur_ms=1600)
            actions_taken.append(f"Loop 1 (Targeted): Injected Morph transitions (1600ms) on slides {missing_trans}.")

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
            if len(prs.slides) > 1:
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
