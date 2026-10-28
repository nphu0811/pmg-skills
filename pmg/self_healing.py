"""
pmg.self_healing
================
Shared single-attempt repair operations called by PMGOrchestrator. Legacy
scripts.motion_validator delegates here. Repairs address
both static OpenXML deficiencies and visual/temporal/visibility playback defects
detected during real PowerPoint slideshow review (including intra-slide <p:timing>
animations and disappearing hero images/titles).
"""

import os
import shutil
import zipfile
from typing import List, Dict, Any, Tuple
import pptx
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from lxml import etree

from scripts.motion_graph_kit import apply_transition
from pmg.models import ReviewDefect
from pmg.animations import P_NS, NSMAP, SlideAnimationTimeline, iter_native_shapes


def repair_manifest_content_type(pptx_path: str) -> List[str]:
    """Repair only the known manifest mapping; never guess arbitrary MIME types."""
    temporary = pptx_path + '.opc.tmp'
    namespace = 'http://schemas.openxmlformats.org/package/2006/content-types'
    try:
        with zipfile.ZipFile(pptx_path) as source:
            if 'ppt/pmg_execution_manifest.json' not in source.namelist():
                return []
            root = etree.fromstring(source.read('[Content_Types].xml'))
            for node in list(root):
                if node.get('PartName') == '/ppt/pmg_execution_manifest.json' or node.get('Extension', '').lower() == 'json':
                    root.remove(node)
            etree.SubElement(root, f'{{{namespace}}}Default', Extension='json', ContentType='application/json')
            with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as output:
                for item in source.infolist():
                    output.writestr(item, etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                                    if item.filename == '[Content_Types].xml' else source.read(item.filename))
        os.replace(temporary, pptx_path)
        return ['Registered execution manifest ContentType application/json.']
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)


def repair_xml_transitions(pptx_path: str, target_slides: List[int], trans_type: str = "morph", dur_ms: int = 1600):
    """Repair transition parameters through the existing shared compiler."""
    presentation = pptx.Presentation(pptx_path)
    changed = False
    for number, slide in enumerate(presentation.slides, 1):
        if number in target_slides and number > 1:
            apply_transition(slide, trans_type=trans_type, duration_ms=dur_ms)
            changed = True
    if changed:
        presentation.save(pptx_path)


def repair_traveling_anchor_gap(pptx_path: str, anchor_tag: str, gap_slides: List[int]):
    """Bridges missing traveling anchors on intermediate slides to prevent morph jump cuts."""
    prs = pptx.Presentation(pptx_path)
    modified = False

    for s_idx in gap_slides:
        if 1 <= s_idx <= len(prs.slides):
            slide = prs.slides[s_idx - 1]
            # Check if anchor is already there
            has_tag = any(shape.name == anchor_tag for shape in slide.shapes)
            if not has_tag:
                # Add an offscreen or subtle proxy shape with this name so morph stays continuous
                shape = slide.shapes.add_shape(
                    pptx.enum.shapes.MSO_SHAPE.RECTANGLE,
                    pptx.util.Inches(-2.5), pptx.util.Inches(3.5),
                    pptx.util.Inches(0.5), pptx.util.Inches(0.5)
                )
                shape.name = anchor_tag
                # Make proxy invisible
                shape.fill.background()
                shape.line.fill.background()
                modified = True

    if modified:
        prs.save(pptx_path)


def substitute_with_fallback_skill(pptx_path: str, failed_slides: List[int], reason: str) -> str:
    """
    Substitutes an intractable failure with an alternative skill from the PMG suite:
    - Failed Morph / 3D Dolly -> Fallback to pmg-choreography (Circle Wipe Spotlight or Split Reveal)
    - Broken complex motion -> Fallback to pmg-camera-parallax or pmg-glassmorphism card transition.
    """
    print(f"[*] Activating Fallback Skill Substitution for Slide(s) {failed_slides}...")
    print(f"    Reason: {reason}")
    print("    Alternative Skill Activated: 'pmg-choreography' (Iris Spotlight / Circle Wipe)")

    # Substitute failed slide transitions with high-impact Circle Spotlight or Split Curtain
    repair_xml_transitions(pptx_path, failed_slides, trans_type="circle", dur_ms=1400)
    return "pmg-choreography"



def _xp(el: Any, expr: str) -> List[Any]:
    """Safe XPath helper compatible with both python-pptx BaseOxmlElement and raw lxml.etree._Element."""
    try:
        return el.xpath(expr)
    except TypeError:
        return el.xpath(expr, namespaces=NSMAP)
    except etree.XPathEvalError:
        return el.xpath(expr, namespaces=NSMAP)


def _repair_disappearing_object_timing(pptx_path: str, loop_number: int) -> List[str]:
    """
    Repairs OpenXML <p:timing> defects that cause hero images, titles, or animated
    objects to disappear during Slide Show while preserving legitimate Fade In,
    Wipe, Fly In, Emphasis, and Motion Path animations:
    - Loop 1 (Targeted):
      Extracts all intended animations on each affected slide, deduplicates conflicting
      multiple entrance effects (`presetClass="entr"`) on the same `spid`, injects
      mandatory `<p:set>` (`style.visibility = "visible"`) transitions, binds `<p:tn val="2"/>`
      to `evt="onBegin"`, and synchronizes 3-level `mainSeq` `<p:par>` hierarchy and `<p:bldLst>`
      via `SlideAnimationTimeline.compile_to_openxml`.
    - Loop 2 (Structural):
      Rebuilds `<p:timing>` with normalized stagger delays and durations (`600ms-800ms`) so all
      entrance animations complete well before slide steady state.
    - Loop 3 (Compatible Fallback):
      Rebuilds `<p:timing>` substituting any problematic entrance filter with a verified
      compatible native entrance (`fade`).
    """
    actions: List[str] = []
    prs = pptx.Presentation(pptx_path)
    modified = False

    for idx, slide in enumerate(prs.slides, 1):
        sld = slide._element
        timing = sld.find(f"{{{P_NS}}}timing")
        if timing is None:
            continue

        shape_by_id = {str(s.shape_id): s for s in iter_native_shapes(slide.shapes)}
        if not shape_by_id:
            continue

        effect_ctns = _xp(
            timing,
            './/p:cTn[@presetClass="entr" or @presetClass="emph" or @presetClass="exit" or @presetClass="path"]',
        )
        if not effect_ctns:
            continue

        tl = SlideAnimationTimeline(slide_number=idx)
        seen_entr_spids = set()
        extracted_count = 0
        dedup_count = 0

        for step_i, ctn in enumerate(effect_ctns):
            p_class = ctn.get("presetClass", "entr")
            p_id = ctn.get("presetID", "10")
            node_type = ctn.get("nodeType", "withEffect")
            sp_tgts = _xp(ctn, ".//p:spTgt")
            if not sp_tgts:
                continue
            spid = sp_tgts[0].get("spid", "")
            if spid not in shape_by_id:
                continue

            if p_class == "entr":
                if spid in seen_entr_spids:
                    dedup_count += 1
                    continue
                seen_entr_spids.add(spid)

            # Determine delay & duration
            delay_ms = 0
            st_conds = _xp(ctn, "./p:stCondLst/p:cond")
            if st_conds:
                d_str = st_conds[0].get("delay", "0")
                if d_str and d_str.isdigit():
                    delay_ms = int(d_str)

            dur_ms = 650
            beh_ctns = _xp(
                ctn,
                "./p:childTnLst/*[self::p:animEffect or self::p:anim or self::p:animScale or self::p:animRot or self::p:animMotion]/p:cBhvr/p:cTn",
            )
            if beh_ctns:
                b_dur = beh_ctns[0].get("dur", "650")
                if b_dur and b_dur.isdigit():
                    dur_ms = max(150, int(b_dur))

            if loop_number >= 2:
                # Normalize timing window so animations finish well inside slide hold window
                delay_ms = min(400, step_i * 80)
                dur_ms = min(650, dur_ms)
            else:
                delay_ms = min(500, delay_ms)
                dur_ms = min(800, dur_ms)

            trig_map = {
                "clickEffect": "on_click",
                "withEffect": "with_previous",
                "afterEffect": "after_previous",
            }
            # In broken decks where separate top-level <p:par> siblings caused unwanted waits,
            # ensure automatic entrance effects use with_previous so they auto-start on slide entry
            trig = trig_map.get(node_type, "with_previous")
            if p_class == "entr" and trig == "on_click" and loop_number >= 2:
                trig = "with_previous"

            shp = shape_by_id[spid]

            if p_class == "path":
                am = _xp(ctn, ".//p:animMotion")
                path_str = am[0].get("path", "M 0 0 L 0.15 0 E") if am else "M 0 0 L 0.15 0 E"
                tl.add_effect(
                    shp,
                    "motion_path",
                    trigger=trig,
                    duration_ms=dur_ms,
                    delay_ms=delay_ms,
                    motion_path=path_str,
                )
            elif p_class == "exit":
                tl.add_effect(shp, "fade_out", trigger=trig, duration_ms=dur_ms, delay_ms=delay_ms)
            elif p_class == "emph":
                eff_type = "spin" if p_id == "8" else ("pulse" if p_id == "26" else "grow_shrink")
                tl.add_effect(shp, eff_type, trigger=trig, duration_ms=dur_ms, delay_ms=delay_ms)
            else:
                if loop_number >= 3:
                    eff_type = "fade"
                else:
                    eff_type = (
                        "fly_in"
                        if p_id == "2"
                        else ("wipe" if p_id == "22" else ("zoom_in" if p_id == "53" else "fade"))
                    )
                tl.add_effect(shp, eff_type, trigger=trig, duration_ms=dur_ms, delay_ms=delay_ms)
            extracted_count += 1

        if extracted_count > 0:
            tl.compile_to_openxml(slide)
            modified = True
            actions.append(
                f"Repaired <p:timing> on slide {idx}: injected <p:set> visibility='visible', "
                f"synchronized grpId/<p:bldLst>, removed {dedup_count} conflicting duplicate entrance(s), "
                f"and preserved {extracted_count} native animation(s)."
            )

    if modified:
        prs.save(pptx_path)
    return actions


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
        main_seq = _xp(timing, './/p:cTn[@nodeType="mainSeq"]')
        if not main_seq and loop_number >= 2:
            tl = SlideAnimationTimeline(slide_number=idx)
            for s in slide.shapes[:4]:
                tl.add_effect(s, "fade", trigger="with_previous", duration_ms=600)
            tl.compile_to_openxml(slide)
            modified = True
            actions.append(f"Rebuilt structural <p:timing> mainSeq tree on slide {idx}.")
            continue

        # Repair invalid spTgt references
        entr_spids = set()
        for ctn in _xp(timing, './/p:cTn[@presetClass="entr"]'):
            for st in _xp(ctn, ".//p:spTgt"):
                sid = st.get("spid")
                if sid in valid_spids:
                    entr_spids.add(sid)

        for sp_tgt in _xp(timing, ".//p:spTgt"):
            spid = sp_tgt.get("spid", "")
            if spid not in valid_spids:
                # Find parent cTn presetClass
                p_ctn = sp_tgt
                while p_ctn is not None and p_ctn.tag != f"{{{P_NS}}}cTn":
                    p_ctn = p_ctn.getparent()
                is_entr = p_ctn is not None and p_ctn.get("presetClass") == "entr"

                chosen_spid = None
                if is_entr:
                    for s in valid_spids:
                        if s not in entr_spids:
                            chosen_spid = s
                            break
                if not chosen_spid:
                    chosen_spid = valid_spids[0]

                sp_tgt.set("spid", chosen_spid)
                entr_spids.add(chosen_spid)
                modified = True
                actions.append(f"Rebound broken animation target spid={spid} -> spid={chosen_spid} on slide {idx}.")

        # Repair invalid or fallback motion paths
        for am in _xp(timing, ".//p:animMotion"):
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
    allow_substitution: bool = True,
) -> Dict[str, Any]:
    """
    Executes one self-healing loop according to the 3-Loop Repair Architecture:
    - Loop 1: Targeted repair (missing transitions, invalid durations, disappearing object <p:set>/grpId timing, direct animation targets/paths)
    - Loop 2: Structural repair (Morph anchor interpolation, canonical <p:timing> tree rebuild via SlideAnimationTimeline, geometry healing)
    - Loop 3: Compatible fallback repair (fallback native transition / simpler native entrance & motion path substitution)
    """
    if loop_number not in (1, 2, 3):
        raise ValueError('Repair attempt must be 1, 2 or 3.')
    actions_taken: List[str] = []
    if any(d.remediation_hint == 'declare_part_content_type' and
           d.object_id == 'ppt/pmg_execution_manifest.json' for d in defects):
        actions_taken.extend(repair_manifest_content_type(pptx_path))
    missing_trans = static_validation.get("missing_transitions", [])
    discontinuous = static_validation.get("discontinuous_anchors", {})

    defect_slides = sorted(
        list({d.slide_number for d in defects if d.slide_number > 1 and d.severity in ("critical", "major")})
    )

    # Check for disappearing object timing defects or timing tree rebuilds (Upgrade D)
    if any(d.remediation_hint in ("repair_disappearing_object_timing", "rebuild_timing_tree") for d in defects):
        d_acts = _repair_disappearing_object_timing(pptx_path, loop_number)
        actions_taken.extend([f"Loop {loop_number}: {a}" for a in d_acts])

    # Check for intra-slide <p:timing> target/path defects
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

        if target_fallback_slides and allow_substitution:
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
