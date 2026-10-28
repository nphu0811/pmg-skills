"""
pmg.validator
=============
Deep OpenXML static validator and ExpectedTimeline extractor for PMG presentations.
Verifies:
- Valid ZIP / OpenXML package structure
- Unique shape identifiers per slide
- Valid transition XML and duration attributes
- Valid intra-slide <p:timing> animation trees, spTgt references, and motion paths (v1.2+)
- Continuous '!!' Morph anchors across active sequences
- Valid media relationships (ppt/media/*)
- Native editable PowerPoint objects check (blocks flattened screenshot shortcuts)
- Object geometry and ExpectedTimeline generation
"""

import os
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Any, Tuple, Optional

from pptx import Presentation

from pmg.version import VERSION
from pmg.models import (
    ValidationStatus,
    ObjectMotionSpec,
    SlideTimelineSpec,
    ExpectedTimeline,
    ReviewDefect,
)
from pmg.animations import parse_slide_timing_openxml

EMU_PER_INCH = 914400.0

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
}


def _parse_shape_node(elem: ET.Element, obj_type: str) -> Optional[Dict[str, Any]]:
    """Extracts id, name, position, size, rotation, alpha, and text from an OpenXML shape node."""
    cNvPr = elem.find(".//p:cNvPr", NS)
    if cNvPr is None:
        return None
    shp_id = cNvPr.attrib.get("id", "0")
    name = cNvPr.attrib.get("name", "")

    xfrm = elem.find(".//a:xfrm", NS)
    x_in, y_in, w_in, h_in, rot_deg = 0.0, 0.0, 0.0, 0.0, 0.0
    if xfrm is not None:
        rot_str = xfrm.attrib.get("rot", "0")
        try:
            rot_deg = int(rot_str) / 60000.0
        except ValueError:
            rot_deg = 0.0
        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        if off is not None:
            x_in = int(off.attrib.get("x", "0")) / EMU_PER_INCH
            y_in = int(off.attrib.get("y", "0")) / EMU_PER_INCH
        if ext is not None:
            w_in = int(ext.attrib.get("cx", "0")) / EMU_PER_INCH
            h_in = int(ext.attrib.get("cy", "0")) / EMU_PER_INCH

    # Parse alpha if present
    alpha_elem = elem.find(".//a:solidFill//a:alpha", NS)
    alpha_val = 1.0
    if alpha_elem is not None:
        try:
            alpha_val = int(alpha_elem.attrib.get("val", "100000")) / 100000.0
        except ValueError:
            alpha_val = 1.0

    # Parse text
    texts = [t.text for t in elem.findall(".//a:t", NS) if t.text]
    full_text = " ".join(texts).strip()

    return {
        "id": shp_id,
        "name": name,
        "type": obj_type,
        "x": round(x_in, 4),
        "y": round(y_in, 4),
        "w": round(w_in, 4),
        "h": round(h_in, 4),
        "rot_deg": round(rot_deg, 2),
        "alpha": round(alpha_val, 3),
        "text": full_text,
    }


def extract_slide_objects(xml_bytes: bytes) -> List[Dict[str, Any]]:
    """Parses all top-level and grouped shapes from slide XML."""
    root = ET.fromstring(xml_bytes)
    spTree = root.find(".//p:cSld/p:spTree", NS)
    if spTree is None:
        return []

    objects: List[Dict[str, Any]] = []
    tag_map = {
        f"{{{NS['p']}}}sp": "shape",
        f"{{{NS['p']}}}pic": "picture",
        f"{{{NS['p']}}}grpSp": "group",
        f"{{{NS['p']}}}cxnSp": "connector",
        f"{{{NS['p']}}}graphicFrame": "diagram",
    }

    for child in spTree:
        obj_type = tag_map.get(child.tag)
        if obj_type:
            parsed = _parse_shape_node(child, obj_type)
            if parsed:
                if obj_type == "shape" and parsed["text"]:
                    parsed["type"] = "textbox" if "TextBox" in parsed["name"] else "shape"
                objects.append(parsed)
    return objects


def parse_slide_transition(xml_text: str, slide_idx: int) -> Tuple[str, int, bool]:
    """Returns (transition_type, duration_ms, has_explicit_transition)."""
    trans_match = re.search(r"<p:transition([^>]*)>(.*?)</p:transition>", xml_text, re.DOTALL)
    if not trans_match:
        return ("none" if slide_idx == 1 else "missing", 0, False)

    t_attrs = trans_match.group(1)
    t_inner = trans_match.group(2)
    dur_m = re.search(r'dur="(\d+)"', t_attrs)
    dur_ms = int(dur_m.group(1)) if dur_m else 1600

    if "morph" in t_inner:
        return ("morph", dur_ms, True)
    elif "<p:circle" in t_inner:
        return ("circle", dur_ms, True)
    elif "<p:split" in t_inner:
        return ("split", dur_ms, True)
    elif "<p:fade" in t_inner:
        return ("fade", dur_ms, True)
    elif "<p:push" in t_inner:
        return ("push", dur_ms, True)
    elif "<p:wipe" in t_inner:
        return ("wipe", dur_ms, True)
    else:
        return ("native", dur_ms, True)


def extract_expected_timeline(pptx_path: str, hold_duration_ms: int = 2000) -> ExpectedTimeline:
    """
    Constructs a deterministic ExpectedTimeline from a PPTX file, capturing:
    - Inter-slide Morph object movements (initial/final pos, size, rotation, opacity)
    - Intra-slide native animations (in <p:timing>)
    - Absolute, slide-relative, and click-triggered playback script
    """
    deck_title = os.path.splitext(os.path.basename(pptx_path))[0]
    slides_spec: List[SlideTimelineSpec] = []
    playback_script: List[Dict[str, Any]] = []

    prs = Presentation(pptx_path)

    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_files = [n for n in z.namelist() if re.match(r"^ppt/slides/slide\d+\.xml$", n)]
        slide_files.sort(key=lambda x: int(re.search(r"\d+", x).group()))

        prev_objects_by_name: Dict[str, Dict[str, Any]] = {}
        current_time_ms = 0

        playback_script.append({
            "event": "start_slideshow",
            "slide_number": 1,
            "time_ms": 0,
        })

        for idx, s_file in enumerate(slide_files, 1):
            xml_bytes = z.read(s_file)
            xml_text = xml_bytes.decode("utf-8", errors="ignore")
            objects = extract_slide_objects(xml_bytes)
            trans_type, trans_dur, _ = parse_slide_transition(xml_text, idx)
            if idx == 1:
                trans_dur = 0

            slide_start_ms = current_time_ms
            motion_specs: List[ObjectMotionSpec] = []
            seq_idx = 0

            curr_objects_by_name = {o["name"]: o for o in objects if o["name"]}

            # 1. Track Morph / Cross-slide object movements
            for obj in objects:
                name = obj["name"]
                if idx > 1 and name.startswith("!!") and name in prev_objects_by_name:
                    prev_obj = prev_objects_by_name[name]
                    seq_idx += 1
                    motion_specs.append(
                        ObjectMotionSpec(
                            object_id=name,
                            slide_number=idx,
                            object_type=obj["type"],
                            initial_position=(prev_obj["x"], prev_obj["y"]),
                            final_position=(obj["x"], obj["y"]),
                            initial_size=(prev_obj["w"], prev_obj["h"]),
                            final_size=(obj["w"], obj["h"]),
                            rotation_params={
                                "start_deg": prev_obj["rot_deg"],
                                "end_deg": obj["rot_deg"],
                            },
                            opacity_params={
                                "start_alpha": prev_obj["alpha"],
                                "end_alpha": obj["alpha"],
                            },
                            motion_path=f"morph_linear:({prev_obj['x']},{prev_obj['y']})->({obj['x']},{obj['y']})",
                            start_time_ms=0,
                            duration_ms=trans_dur if trans_dur > 0 else 1600,
                            delay_ms=0,
                            trigger_type="slide_transition",
                            animation_sequence=seq_idx,
                            easing="smooth_morph",
                            time_reference="slide_relative",
                            absolute_start_ms=slide_start_ms,
                            absolute_end_ms=slide_start_ms + (trans_dur if trans_dur > 0 else 1600),
                        )
                    )
                else:
                    motion_specs.append(
                        ObjectMotionSpec(
                            object_id=name or f"shape_{obj['id']}",
                            slide_number=idx,
                            object_type=obj["type"],
                            initial_position=(obj["x"], obj["y"]),
                            final_position=(obj["x"], obj["y"]),
                            initial_size=(obj["w"], obj["h"]),
                            final_size=(obj["w"], obj["h"]),
                            rotation_params={"start_deg": obj["rot_deg"], "end_deg": obj["rot_deg"]},
                            opacity_params={"start_alpha": obj["alpha"], "end_alpha": obj["alpha"]},
                            motion_path=None,
                            start_time_ms=0,
                            duration_ms=0,
                            delay_ms=0,
                            trigger_type="static",
                            animation_sequence=0,
                            easing="none",
                            time_reference="slide_relative",
                            absolute_start_ms=slide_start_ms,
                            absolute_end_ms=slide_start_ms + trans_dur + hold_duration_ms,
                        )
                    )

            # 2. Parse intra-slide <p:timing> animations using pmg.animations
            clicks: List[Dict[str, Any]] = []
            max_anim_end_rel = trans_dur
            if idx <= len(prs.slides):
                intra_specs, intra_clicks = parse_slide_timing_openxml(
                    prs.slides[idx - 1],
                    slide_number=idx,
                    slide_abs_start_ms=slide_start_ms,
                    transition_offset_ms=trans_dur,
                )
                for ispec in intra_specs:
                    seq_idx += 1
                    ispec.animation_sequence = seq_idx
                    motion_specs.append(ispec)
                    rel_end = ispec.absolute_end_ms - slide_start_ms
                    if rel_end > max_anim_end_rel:
                        max_anim_end_rel = rel_end
                clicks.extend(intra_clicks)

            slide_end_ms = slide_start_ms + max_anim_end_rel + hold_duration_ms

            if idx > 1:
                playback_script.append({
                    "event": "advance_slide",
                    "from_slide": idx - 1,
                    "to_slide": idx,
                    "time_ms": slide_start_ms,
                    "transition_type": trans_type,
                    "transition_duration_ms": trans_dur,
                })

            for cl in clicks:
                playback_script.append({
                    "event": "click_trigger",
                    "slide_number": idx,
                    "sequence_index": cl["sequence_index"],
                    "target_object": cl["target_object"],
                    "time_ms": cl["absolute_ms"],
                })

            playback_script.append({
                "event": "hold_slide",
                "slide_number": idx,
                "start_ms": slide_start_ms + max_anim_end_rel,
                "end_ms": slide_end_ms,
                "hold_duration_ms": hold_duration_ms,
            })

            slides_spec.append(
                SlideTimelineSpec(
                    slide_number=idx,
                    transition_type=trans_type,
                    transition_duration_ms=trans_dur,
                    hold_duration_ms=hold_duration_ms,
                    absolute_start_ms=slide_start_ms,
                    absolute_end_ms=slide_end_ms,
                    objects=motion_specs,
                    clicks=clicks,
                )
            )

            prev_objects_by_name = curr_objects_by_name
            current_time_ms = slide_end_ms

        playback_script.append({
            "event": "end_slideshow",
            "time_ms": current_time_ms,
        })

    return ExpectedTimeline(
        deck_title=deck_title,
        engine_version=VERSION,
        total_slides=len(slides_spec),
        total_duration_ms=current_time_ms,
        slides=slides_spec,
        playback_script=playback_script,
    )


class StaticValidator:
    """
    Performs static OpenXML, geometry, media, transition, intra-slide animation timing,
    and anti-shortcut validation.
    """

    def validate(self, pptx_path: str) -> Dict[str, Any]:
        defects: List[ReviewDefect] = []
        slide_summaries: List[Dict[str, Any]] = []
        traveling_anchors: Dict[str, List[int]] = {}
        discontinuous_anchors: Dict[str, List[int]] = {}
        missing_transitions: List[int] = []
        total_native_animations = 0

        if not os.path.exists(pptx_path):
            defects.append(
                ReviewDefect(
                    category="openxml_structure",
                    severity="critical",
                    slide_number=0,
                    object_id=None,
                    description=f"Presentation file does not exist: {pptx_path}",
                )
            )
            return {
                "is_valid": False,
                "total_slides": 0,
                "defects": [d.to_dict() for d in defects],
            }

        try:
            with zipfile.ZipFile(pptx_path, "r") as z:
                namelist = set(z.namelist())
                if "ppt/presentation.xml" not in namelist:
                    defects.append(
                        ReviewDefect(
                            category="openxml_structure",
                            severity="critical",
                            slide_number=0,
                            object_id=None,
                            description="Missing ppt/presentation.xml in package.",
                        )
                    )
                    return {
                        "is_valid": False,
                        "total_slides": 0,
                        "defects": [d.to_dict() for d in defects],
                    }

                # Check 16:9 widescreen dimensions in presentation.xml
                pres_xml = z.read("ppt/presentation.xml").decode("utf-8", errors="ignore")
                sldSz = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', pres_xml)
                if sldSz:
                    cx, cy = int(sldSz.group(1)), int(sldSz.group(2))
                    ratio = cx / cy if cy > 0 else 0
                    if abs(ratio - (16.0 / 9.0)) > 0.05:
                        defects.append(
                            ReviewDefect(
                                category="visual_quality",
                                severity="major",
                                slide_number=0,
                                object_id=None,
                                description=f"Canvas aspect ratio is not 16:9 (cx={cx}, cy={cy}).",
                            )
                        )

                slide_files = [n for n in namelist if re.match(r"^ppt/slides/slide\d+\.xml$", n)]
                slide_files.sort(key=lambda x: int(re.search(r"\d+", x).group()))
                total_slides = len(slide_files)

                if total_slides == 0:
                    defects.append(
                        ReviewDefect(
                            category="openxml_structure",
                            severity="critical",
                            slide_number=0,
                            object_id=None,
                            description="Presentation contains 0 slides.",
                        )
                    )

                for idx, s_file in enumerate(slide_files, 1):
                    xml_bytes = z.read(s_file)
                    xml_text = xml_bytes.decode("utf-8", errors="ignore")
                    objects = extract_slide_objects(xml_bytes)

                    # 1. Check shape ID uniqueness
                    seen_ids = set()
                    for o in objects:
                        if o["id"] in seen_ids:
                            defects.append(
                                ReviewDefect(
                                    category="openxml_structure",
                                    severity="major",
                                    slide_number=idx,
                                    object_id=o["name"],
                                    description=f"Duplicate shape ID {o['id']} on slide {idx}.",
                                )
                            )
                        seen_ids.add(o["id"])

                    # 2. Check Native Editable Objects (Anti-Shortcut Rule)
                    editable_objs = [o for o in objects if o["type"] in ("shape", "textbox", "group", "diagram", "connector")]
                    pictures = [o for o in objects if o["type"] == "picture"]
                    if len(editable_objs) == 0 and len(pictures) >= 1:
                        defects.append(
                            ReviewDefect(
                                category="openxml_structure",
                                severity="critical",
                                slide_number=idx,
                                object_id=pictures[0]["name"],
                                description=f"Slide {idx} contains only image(s) and zero editable shapes/text (flattened slide shortcut is forbidden).",
                            )
                        )
                    elif len(objects) < 2:
                        defects.append(
                            ReviewDefect(
                                category="visual_quality",
                                severity="major",
                                slide_number=idx,
                                object_id=None,
                                description=f"Slide {idx} has fewer than 2 independent objects ({len(objects)} found).",
                            )
                        )

                    # 3. Check Transition
                    trans_type, trans_dur, has_trans = parse_slide_transition(xml_text, idx)
                    if idx > 1 and not has_trans:
                        missing_transitions.append(idx)
                        defects.append(
                            ReviewDefect(
                                category="animation_behavior",
                                severity="major",
                                slide_number=idx,
                                object_id=None,
                                description=f"Slide {idx} is missing a native PowerPoint transition.",
                                remediation_hint="inject_transition",
                            )
                        )
                    elif idx > 1 and (trans_dur < 300 or trans_dur > 5000):
                        defects.append(
                            ReviewDefect(
                                category="temporal_quality",
                                severity="minor",
                                slide_number=idx,
                                object_id=None,
                                description=f"Slide {idx} transition duration ({trans_dur}ms) is outside recommended bounds (300ms-5000ms).",
                                remediation_hint="normalize_duration",
                            )
                        )

                    # 4. Validate Intra-Slide <p:timing> Native Animations
                    root = ET.fromstring(xml_bytes)
                    timing = root.find("p:timing", NS)
                    slide_anim_count = 0
                    if timing is not None:
                        main_seq = timing.find('.//p:cTn[@nodeType="mainSeq"]', NS)
                        if main_seq is None:
                            defects.append(
                                ReviewDefect(
                                    category="openxml_structure",
                                    severity="major",
                                    slide_number=idx,
                                    object_id=None,
                                    description=f"Slide {idx} has <p:timing> without a valid mainSeq node.",
                                    remediation_hint="rebuild_timing_tree",
                                )
                            )
                        for ctn in timing.findall(".//p:cTn", NS):
                            p_cls = ctn.attrib.get("presetClass")
                            if p_cls in ("entr", "emph", "path", "exit"):
                                slide_anim_count += 1
                                sp_tgts = ctn.findall(".//p:spTgt", NS)
                                if not sp_tgts:
                                    defects.append(
                                        ReviewDefect(
                                            category="animation_behavior",
                                            severity="major",
                                            slide_number=idx,
                                            object_id=None,
                                            description=f"Animation effect on slide {idx} is missing a <p:spTgt> target.",
                                            remediation_hint="repair_animation_target",
                                        )
                                    )
                                else:
                                    for st in sp_tgts:
                                        tgt_spid = st.attrib.get("spid", "")
                                        if tgt_spid not in seen_ids:
                                            defects.append(
                                                ReviewDefect(
                                                    category="animation_behavior",
                                                    severity="major",
                                                    slide_number=idx,
                                                    object_id=f"spid_{tgt_spid}",
                                                    description=f"Animation effect on slide {idx} targets nonexistent shape spid={tgt_spid}.",
                                                    remediation_hint="repair_animation_target",
                                                )
                                            )
                                for am in ctn.findall(".//p:animMotion", NS):
                                    path_attr = am.attrib.get("path", "").strip()
                                    if not path_attr or not path_attr.upper().startswith("M"):
                                        defects.append(
                                            ReviewDefect(
                                                category="animation_behavior",
                                                severity="major",
                                                slide_number=idx,
                                                object_id=None,
                                                description=f"Invalid motion path syntax '{path_attr}' on slide {idx}.",
                                                remediation_hint="repair_motion_path",
                                            )
                                        )
                    total_native_animations += slide_anim_count

                    # 5. Track !! Morph anchors
                    for o in objects:
                        if o["name"].startswith("!!"):
                            traveling_anchors.setdefault(o["name"], []).append(idx)

                    # 6. Check Media Relationships
                    rels_path = f"ppt/slides/_rels/slide{idx}.xml.rels"
                    if rels_path in namelist:
                        rels_xml = z.read(rels_path).decode("utf-8", errors="ignore")
                        targets = re.findall(r'Target="\.\./media/([^"]+)"', rels_xml)
                        for tgt in targets:
                            media_file = f"ppt/media/{tgt}"
                            if media_file not in namelist:
                                defects.append(
                                    ReviewDefect(
                                        category="openxml_structure",
                                        severity="critical",
                                        slide_number=idx,
                                        object_id=tgt,
                                        description=f"Slide {idx} references missing media file {media_file}.",
                                    )
                                )

                    # 7. Word count
                    word_count = sum(len(o["text"].split()) for o in objects if o["text"])

                    slide_summaries.append({
                        "slide_number": idx,
                        "object_count": len(objects),
                        "editable_object_count": len(editable_objs),
                        "picture_count": len(pictures),
                        "transition_type": trans_type,
                        "transition_duration_ms": trans_dur,
                        "native_animation_count": slide_anim_count,
                        "word_count": word_count,
                        "morph_anchors": [o["name"] for o in objects if o["name"].startswith("!!")],
                    })

                # Check Traveling Entity Continuity for primary traveling entities
                for tag, s_list in traveling_anchors.items():
                    s_list.sort()
                    if len(s_list) > 1 and ("Traveling" in tag or "prog_" in tag):
                        gaps = []
                        for s_num in range(min(s_list), max(s_list) + 1):
                            if s_num not in s_list:
                                gaps.append(s_num)
                        if gaps:
                            discontinuous_anchors[tag] = gaps
                            defects.append(
                                ReviewDefect(
                                    category="animation_behavior",
                                    severity="major",
                                    slide_number=gaps[0],
                                    object_id=tag,
                                    description=f"Traveling entity '{tag}' has continuity gap on slide(s) {gaps} (present on {s_list}).",
                                    remediation_hint="bridge_anchor_gap",
                                )
                            )

        except Exception as e:
            defects.append(
                ReviewDefect(
                    category="openxml_structure",
                    severity="critical",
                    slide_number=0,
                    object_id=None,
                    description=f"Failed to parse PPTX package: {str(e)}",
                )
            )
            return {
                "is_valid": False,
                "total_slides": 0,
                "defects": [d.to_dict() for d in defects],
            }

        blocking = [d for d in defects if d.severity in ("critical", "major")]
        return {
            "is_valid": len(blocking) == 0,
            "total_slides": len(slide_summaries),
            "total_native_animations": total_native_animations,
            "missing_transitions": missing_transitions,
            "traveling_anchors": traveling_anchors,
            "discontinuous_anchors": discontinuous_anchors,
            "slides": slide_summaries,
            "defects": [d.to_dict() for d in defects],
        }
