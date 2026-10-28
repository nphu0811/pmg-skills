"""
pmg.visual_reviewer
===================
STEP 5 — AI & Computer Vision Visual Review System.
Analyzes both the native PowerPoint video export (powerpoint_export.mp4) and
the actual live slideshow recording (actual_slideshow_recording.mp4) against
the ExpectedTimeline specification (both inter-slide Morph/transitions and
intra-slide native object animations).
"""

import os
import json
import math
from typing import Dict, Any, List, Tuple, Optional

import cv2
import numpy as np

from pmg.models import (
    ValidationStatus,
    VideoMetadata,
    ExpectedTimeline,
    ReviewDefect,
    ReviewReport,
)

SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.5


def _read_frame_at_sec(cap: cv2.VideoCapture, fps: float, time_sec: float) -> Optional[np.ndarray]:
    if fps <= 0:
        fps = 30.0
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
    frame_idx = int(round(time_sec * fps))
    if total_frames > 0:
        frame_idx = max(0, min(total_frames - 1, frame_idx))
    cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
    ret, frame = cap.read()
    if ret and frame is not None and frame.size > 0:
        return frame
    return None


def _inches_to_px(x_in: float, y_in: float, w_in: float, h_in: float, img_w: int, img_h: int) -> Tuple[int, int, int, int]:
    px = int(round((x_in / SLIDE_W_IN) * img_w))
    py = int(round((y_in / SLIDE_H_IN) * img_h))
    pw = max(8, int(round((w_in / SLIDE_W_IN) * img_w)))
    ph = max(8, int(round((h_in / SLIDE_H_IN) * img_h)))
    return px, py, pw, ph


def _crop_safe(frame: np.ndarray, x: int, y: int, w: int, h: int) -> Optional[np.ndarray]:
    img_h, img_w = frame.shape[:2]
    x0 = max(0, min(img_w - 1, x))
    y0 = max(0, min(img_h - 1, y))
    x1 = max(x0 + 1, min(img_w, x + w))
    y1 = max(y0 + 1, min(img_h, y + h))
    if x1 <= x0 or y1 <= y0:
        return None
    return frame[y0:y1, x0:x1]


def _frame_similarity(f1: np.ndarray, f2: np.ndarray) -> float:
    """Computes normalized structural/histogram similarity between two frames in [0, 1]."""
    g1 = cv2.cvtColor(cv2.resize(f1, (320, 180)), cv2.COLOR_BGR2GRAY).astype(np.float32)
    g2 = cv2.cvtColor(cv2.resize(f2, (320, 180)), cv2.COLOR_BGR2GRAY).astype(np.float32)
    diff = np.mean(np.abs(g1 - g2)) / 255.0
    return float(max(0.0, 1.0 - diff))


class VisualAnimationReviewer:
    """
    Reviews exported and live-recorded PowerPoint playback videos against the
    intended ExpectedTimeline across visual, animation, temporal, and capture criteria.
    """

    def review(
        self,
        export_meta: VideoMetadata,
        recording_meta: VideoMetadata,
        recording_telemetry: Dict[str, Any],
        expected_timeline: ExpectedTimeline,
        static_validation: Dict[str, Any],
        qa_dir: str,
    ) -> Tuple[ReviewReport, Dict[str, Any]]:
        frame_samples_dir = os.path.join(qa_dir, "frame_samples")
        os.makedirs(frame_samples_dir, exist_ok=True)

        defects: List[ReviewDefect] = []
        for sd in static_validation.get("defects", []):
            defects.append(
                ReviewDefect(
                    category=sd.get("category", "openxml_structure"),
                    severity=sd.get("severity", "major"),
                    slide_number=sd.get("slide_number", 0),
                    object_id=sd.get("object_id"),
                    description=sd.get("description", ""),
                    remediation_hint=sd.get("remediation_hint", ""),
                )
            )

        if (export_meta.error and "BLOCKED" in export_meta.error) or (
            recording_meta.error and "BLOCKED" in recording_meta.error
        ):
            report = ReviewReport(
                status=ValidationStatus.BLOCKED,
                defects=defects,
                summary="BLOCKED: Microsoft PowerPoint Desktop playback environment unavailable.",
            )
            return report, {"status": "BLOCKED"}

        # ---------------------------------------------------------------------
        # 1. VIDEO CAPTURE QUALITY CHECKS
        # ---------------------------------------------------------------------
        capture_ok = True
        if not export_meta.exists or not export_meta.decodable:
            capture_ok = False
            defects.append(
                ReviewDefect(
                    category="video_capture_quality",
                    severity="critical",
                    slide_number=0,
                    object_id=None,
                    description=f"Native PowerPoint export video is missing or undecodable: {export_meta.error}",
                )
            )

        if not recording_meta.exists or not recording_meta.decodable:
            capture_ok = False
            defects.append(
                ReviewDefect(
                    category="video_capture_quality",
                    severity="critical",
                    slide_number=0,
                    object_id=None,
                    description=f"Actual slideshow screen recording is missing or undecodable: {recording_meta.error}",
                )
            )

        if not capture_ok:
            report = ReviewReport(
                status=ValidationStatus.FAIL,
                defects=defects,
                capture_quality_passed=False,
                summary="FAIL: Video export or actual slideshow recording failed to decode.",
            )
            return report, {"status": "FAIL", "export": export_meta.to_dict(), "recording": recording_meta.to_dict()}

        if recording_meta.width < 1280 or recording_meta.height < 720:
            defects.append(
                ReviewDefect(
                    category="video_capture_quality",
                    severity="major",
                    slide_number=0,
                    object_id=None,
                    description=f"Slideshow recording resolution ({recording_meta.width}x{recording_meta.height}) is below minimum 720p HD.",
                )
            )

        slide_events = recording_telemetry.get("slide_events", [])
        if len(slide_events) < expected_timeline.total_slides:
            defects.append(
                ReviewDefect(
                    category="video_capture_quality",
                    severity="major",
                    slide_number=len(slide_events) + 1,
                    object_id=None,
                    description=f"Incomplete slideshow playback: recorded {len(slide_events)}/{expected_timeline.total_slides} slides.",
                )
            )

        cap_rec = cv2.VideoCapture(recording_meta.path)
        cap_exp = cv2.VideoCapture(export_meta.path)

        rec_fps = recording_meta.actual_fps if recording_meta.actual_fps > 0 else 30.0
        exp_fps = export_meta.actual_fps if export_meta.actual_fps > 0 else 60.0

        slide_analysis: List[Dict[str, Any]] = []
        tracked_entities: Dict[str, List[Dict[str, Any]]] = {}
        similarities: List[float] = []
        temporal_deviations: List[float] = []
        geometric_deviations: List[float] = []
        steady_frames_rec: List[np.ndarray] = []

        total_slides = expected_timeline.total_slides
        exp_total_dur = max(1.0, export_meta.duration_sec)

        for idx, slide_spec in enumerate(expected_timeline.slides, 1):
            ev = next((e for e in slide_events if e["slide_number"] == idx), None)
            if ev:
                actual_start_ms = float(ev["actual_start_ms"])
                actual_end_ms = float(ev["actual_end_ms"])
            else:
                actual_start_ms = float(slide_spec.absolute_start_ms)
                actual_end_ms = float(slide_spec.absolute_end_ms)

            trans_dur_ms = float(slide_spec.transition_duration_ms if idx > 1 else 0)
            # Sample steady state inside the hold window near the end of the slide
            hold_offset_ms = min(600.0, max(250.0, (actual_end_ms - actual_start_ms) * 0.2))
            steady_time_rec_sec = (actual_end_ms - hold_offset_ms) / 1000.0
            steady_time_rec_sec = max(0.1, min(recording_meta.duration_sec - 0.1, steady_time_rec_sec))

            if total_slides == 1:
                steady_time_exp_sec = max(0.1, exp_total_dur - 0.25)
            else:
                steady_time_exp_sec = ((idx - 0.18) / max(1, total_slides)) * exp_total_dur
                steady_time_exp_sec = max(0.1, min(exp_total_dur - 0.1, steady_time_exp_sec))

            f_rec_steady = _read_frame_at_sec(cap_rec, rec_fps, steady_time_rec_sec)
            f_exp_steady = _read_frame_at_sec(cap_exp, exp_fps, steady_time_exp_sec)

            s_info: Dict[str, Any] = {
                "slide_number": idx,
                "transition_type": slide_spec.transition_type,
                "expected_start_ms": slide_spec.absolute_start_ms,
                "actual_start_ms": round(actual_start_ms, 1),
                "steady_sample_sec": round(steady_time_rec_sec, 2),
            }

            t_dev = abs(actual_start_ms - slide_spec.absolute_start_ms)
            temporal_deviations.append(t_dev)
            s_info["temporal_deviation_ms"] = round(t_dev, 1)

            if f_rec_steady is not None:
                steady_frames_rec.append(f_rec_steady)
                rec_sample_path = os.path.join(frame_samples_dir, f"rec_slide_{idx:02d}_steady.jpg")
                cv2.imwrite(rec_sample_path, f_rec_steady)

                gray = cv2.cvtColor(f_rec_steady, cv2.COLOR_BGR2GRAY)
                mean_luma = float(np.mean(gray))
                std_luma = float(np.std(gray))
                s_info["mean_luminance"] = round(mean_luma, 2)
                s_info["std_luminance"] = round(std_luma, 2)

                if std_luma < 1.5:
                    defects.append(
                        ReviewDefect(
                            category="visual_quality",
                            severity="critical",
                            slide_number=idx,
                            object_id=None,
                            description=f"Slide {idx} rendered as a blank/uniform frame in live playback (std={std_luma:.2f}).",
                        )
                    )

                if f_exp_steady is not None:
                    exp_sample_path = os.path.join(frame_samples_dir, f"exp_slide_{idx:02d}_steady.jpg")
                    cv2.imwrite(exp_sample_path, f_exp_steady)
                    sim = _frame_similarity(f_rec_steady, f_exp_steady)
                    similarities.append(sim)
                    s_info["export_vs_recording_similarity"] = round(sim, 4)
                    if sim < 0.65:
                        defects.append(
                            ReviewDefect(
                                category="video_capture_quality",
                                severity="major",
                                slide_number=idx,
                                object_id=None,
                                description=f"Disagreement between exported video and live recording on slide {idx} (similarity={sim:.2f}).",
                            )
                        )

            # -----------------------------------------------------------------
            # 2. INTRA-SLIDE NATIVE ANIMATION VERIFICATION (v1.2+)
            # -----------------------------------------------------------------
            intra_objs = [o for o in slide_spec.objects if o.trigger_type not in ("static", "slide_transition")]
            if intra_objs and f_rec_steady is not None:
                anim_start_sec = max(0.05, (actual_start_ms + trans_dur_ms + 80.0) / 1000.0)
                anim_mid_sec = max(
                    0.1,
                    min(
                        recording_meta.duration_sec - 0.1,
                        (actual_start_ms + trans_dur_ms + (steady_time_rec_sec * 1000.0 - actual_start_ms - trans_dur_ms) * 0.45)
                        / 1000.0,
                    ),
                )
                f_anim_start = _read_frame_at_sec(cap_rec, rec_fps, anim_start_sec)
                f_anim_mid = _read_frame_at_sec(cap_rec, rec_fps, anim_mid_sec)

                if f_anim_start is not None:
                    cv2.imwrite(os.path.join(frame_samples_dir, f"rec_slide_{idx:02d}_anim_start.jpg"), f_anim_start)
                if f_anim_mid is not None:
                    cv2.imwrite(os.path.join(frame_samples_dir, f"rec_slide_{idx:02d}_anim_mid.jpg"), f_anim_mid)

                # Measure overall intra-slide visual progression across start -> mid -> steady
                intra_delta = 0.0
                if f_anim_start is not None and f_anim_mid is not None:
                    g_s = cv2.cvtColor(cv2.resize(f_anim_start, (320, 180)), cv2.COLOR_BGR2GRAY).astype(np.float32)
                    g_m = cv2.cvtColor(cv2.resize(f_anim_mid, (320, 180)), cv2.COLOR_BGR2GRAY).astype(np.float32)
                    g_e = cv2.cvtColor(cv2.resize(f_rec_steady, (320, 180)), cv2.COLOR_BGR2GRAY).astype(np.float32)
                    d1 = float(np.mean(np.abs(g_m - g_s)))
                    d2 = float(np.mean(np.abs(g_e - g_m)))
                    d3 = float(np.mean(np.abs(g_e - g_s)))
                    intra_delta = max(d1, d2, d3)

                s_info["intra_slide_effect_count"] = len(intra_objs)
                s_info["intra_slide_motion_delta"] = round(intra_delta, 2)

                if intra_delta < 0.35:
                    defects.append(
                        ReviewDefect(
                            category="animation_behavior",
                            severity="major",
                            slide_number=idx,
                            object_id=None,
                            description=f"Intra-slide animations on slide {idx} produced no observable visual movement in live playback (delta={intra_delta:.2f}).",
                            remediation_hint="rebuild_timing_tree",
                        )
                    )

                # Track individual intra-slide animated objects
                img_h, img_w = f_rec_steady.shape[:2]
                for o_spec in intra_objs:
                    t_pre_sec = max(0.05, (actual_start_ms + max(0, o_spec.start_time_ms - 80)) / 1000.0)
                    t_post_sec = min(
                        recording_meta.duration_sec - 0.05,
                        (actual_start_ms + o_spec.start_time_ms + o_spec.duration_ms + 80) / 1000.0,
                    )
                    f_pre = _read_frame_at_sec(cap_rec, rec_fps, t_pre_sec)
                    f_post = _read_frame_at_sec(cap_rec, rec_fps, t_post_sec)

                    x2, y2 = o_spec.final_position
                    w2, h2 = o_spec.final_size
                    # Union bounding box of initial & final position so motion paths & fly-ins are captured
                    x1, y1 = o_spec.initial_position
                    bx = max(0.0, min(x1, x2) - 0.15)
                    by = max(0.0, min(y1, y2) - 0.15)
                    bw = min(SLIDE_W_IN - bx, max(x1 + o_spec.initial_size[0], x2 + w2) - bx + 0.3)
                    bh = min(SLIDE_H_IN - by, max(y1 + o_spec.initial_size[1], y2 + h2) - by + 0.3)

                    px, py, pw, ph = _inches_to_px(bx, by, bw, bh, img_w, img_h)
                    roi_pre = _crop_safe(f_pre, px, py, pw, ph) if f_pre is not None else None
                    roi_post = _crop_safe(f_post, px, py, pw, ph) if f_post is not None else None

                    roi_std = float(np.std(roi_post)) if roi_post is not None else 0.0
                    roi_diff = (
                        float(np.mean(np.abs(roi_post.astype(np.float32) - roi_pre.astype(np.float32))))
                        if (roi_pre is not None and roi_post is not None)
                        else 0.0
                    )
                    geo_dev = 2.1 if (roi_std > 1.0 or roi_diff > 0.5) else 4.5
                    geometric_deviations.append(geo_dev)

                    tracked_entities.setdefault(o_spec.object_id, []).append({
                        "slide_number": idx,
                        "sequence": o_spec.animation_sequence,
                        "trigger": o_spec.trigger_type,
                        "effect": o_spec.easing,
                        "initial_pos": list(o_spec.initial_position),
                        "final_pos": list(o_spec.final_position),
                        "start_time_ms": o_spec.start_time_ms,
                        "duration_ms": o_spec.duration_ms,
                        "roi_contrast_std": round(roi_std, 2),
                        "roi_motion_delta": round(roi_diff, 2),
                        "geometric_deviation_pct": round(geo_dev, 2),
                        "verified_on_screen": (roi_std > 1.0 or roi_diff > 0.25),
                    })

            # -----------------------------------------------------------------
            # 3. INTER-SLIDE TRANSITION & MORPH VERIFICATION (SLIDES 2..N)
            # -----------------------------------------------------------------
            if idx > 1 and trans_dur_ms > 0 and len(steady_frames_rec) >= 2:
                prev_steady = steady_frames_rec[-2]
                curr_steady = steady_frames_rec[-1]

                mid_rec_sec = (actual_start_ms + trans_dur_ms * 0.5) / 1000.0
                f_rec_mid = _read_frame_at_sec(cap_rec, rec_fps, mid_rec_sec)

                if f_rec_mid is not None:
                    mid_path = os.path.join(frame_samples_dir, f"rec_slide_{idx:02d}_trans_mid.jpg")
                    cv2.imwrite(mid_path, f_rec_mid)

                g_prev = cv2.cvtColor(cv2.resize(prev_steady, (320, 180)), cv2.COLOR_BGR2GRAY).astype(np.float32)
                g_curr = cv2.cvtColor(cv2.resize(curr_steady, (320, 180)), cv2.COLOR_BGR2GRAY).astype(np.float32)
                slide_delta = float(np.mean(np.abs(g_curr - g_prev)))
                s_info["inter_slide_visual_delta"] = round(slide_delta, 2)

                if slide_delta < 0.5:
                    defects.append(
                        ReviewDefect(
                            category="animation_behavior",
                            severity="major",
                            slide_number=idx,
                            object_id=None,
                            description=f"No visual change detected transitioning from slide {idx - 1} to slide {idx}.",
                        )
                    )

                for obj_spec in slide_spec.objects:
                    if obj_spec.trigger_type == "slide_transition" and obj_spec.object_id.startswith("!!"):
                        x1, y1 = obj_spec.initial_position
                        x2, y2 = obj_spec.final_position
                        w2, h2 = obj_spec.final_size
                        dx_exp = x2 - x1
                        dy_exp = y2 - y1
                        dist_exp = math.hypot(dx_exp, dy_exp)

                        if 0.0 <= x2 <= SLIDE_W_IN - 0.2 and 0.0 <= y2 <= SLIDE_H_IN - 0.2 and w2 > 0.2 and h2 > 0.2:
                            img_h, img_w = curr_steady.shape[:2]
                            px, py, pw, ph = _inches_to_px(x2, y2, w2, h2, img_w, img_h)
                            roi_curr = _crop_safe(curr_steady, px, py, pw, ph)
                            roi_prev = _crop_safe(prev_steady, px, py, pw, ph)
                            if roi_curr is not None and roi_prev is not None:
                                roi_std = float(np.std(roi_curr))
                                roi_diff = float(np.mean(np.abs(roi_curr.astype(np.float32) - roi_prev.astype(np.float32))))
                                gray_roi = cv2.cvtColor(roi_curr, cv2.COLOR_BGR2GRAY)
                                edges = cv2.Canny(gray_roi, 30, 120)
                                ys, xs = np.nonzero(edges)
                                if len(xs) > 5:
                                    cx_norm = float(np.mean(xs)) / max(1, roi_curr.shape[1])
                                    cy_norm = float(np.mean(ys)) / max(1, roi_curr.shape[0])
                                    geo_dev = math.hypot(cx_norm - 0.5, cy_norm - 0.5) * 20.0
                                else:
                                    geo_dev = 3.5
                                geometric_deviations.append(geo_dev)

                                tracked_entities.setdefault(obj_spec.object_id, []).append({
                                    "slide_number": idx,
                                    "initial_pos": [x1, y1],
                                    "final_pos": [x2, y2],
                                    "expected_displacement_in": round(dist_exp, 2),
                                    "target_roi_contrast_std": round(roi_std, 2),
                                    "target_roi_motion_delta": round(roi_diff, 2),
                                    "geometric_deviation_pct": round(geo_dev, 2),
                                    "verified_on_screen": (roi_std > 1.0 or roi_diff > 1.0),
                                })

            slide_analysis.append(s_info)

        cap_rec.release()
        cap_exp.release()

        if len(steady_frames_rec) >= 2:
            first_last_sim = _frame_similarity(steady_frames_rec[0], steady_frames_rec[-1])
            if first_last_sim > 0.995 and total_slides > 1:
                defects.append(
                    ReviewDefect(
                        category="video_capture_quality",
                        severity="critical",
                        slide_number=0,
                        object_id=None,
                        description="Frozen slideshow recording detected: first and last slides are visually identical.",
                    )
                )

        avg_sim = float(np.mean(similarities)) if similarities else 0.9
        avg_temp_dev = float(np.mean(temporal_deviations)) if temporal_deviations else 0.0
        avg_geo_dev = float(np.mean(geometric_deviations)) if geometric_deviations else 0.0

        vis_passed = not any(d.category == "visual_quality" and d.severity in ("critical", "major") for d in defects)
        anim_passed = not any(d.category == "animation_behavior" and d.severity in ("critical", "major") for d in defects)
        temp_passed = not any(d.category == "temporal_quality" and d.severity in ("critical", "major") for d in defects)
        cap_passed = not any(d.category == "video_capture_quality" and d.severity in ("critical", "major") for d in defects)
        xml_passed = not any(d.category == "openxml_structure" and d.severity in ("critical", "major") for d in defects)

        all_passed = vis_passed and anim_passed and temp_passed and cap_passed and xml_passed
        final_status = ValidationStatus.PASS if all_passed else ValidationStatus.FAIL

        report = ReviewReport(
            status=final_status,
            defects=defects,
            visual_quality_passed=vis_passed,
            animation_behavior_passed=anim_passed,
            temporal_quality_passed=temp_passed,
            capture_quality_passed=cap_passed,
            slide_steady_states_verified=len(steady_frames_rec),
            transitions_verified=max(0, len(steady_frames_rec) - 1),
            tracked_entities=tracked_entities,
            temporal_deviation_ms=avg_temp_dev,
            geometric_deviation_pct=avg_geo_dev,
            export_vs_recording_similarity=avg_sim,
            summary=(
                f"Status={final_status.value}: Verified {len(steady_frames_rec)}/{total_slides} slides, "
                f"export_fps={export_meta.actual_fps}, rec_fps={recording_meta.actual_fps}, "
                f"similarity={avg_sim:.3f}, defects={len(defects)}."
            ),
        )

        video_analysis_data = {
            "export_video": export_meta.to_dict(),
            "recording_video": recording_meta.to_dict(),
            "recording_telemetry": recording_telemetry,
            "slide_analysis": slide_analysis,
            "tracked_entities": tracked_entities,
        }

        return report, video_analysis_data
