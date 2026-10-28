"""
Automated tests for PMG Skills v1.3.0:
- CinematicMotionBuilder & MotionPresetRegistry
- Purposeful MotionChoreographyDecision logging
- 10-Slide Environmental Protection Cinematic Presentation (Morph + Parallax + Camera Zoom +
  Animated Infographics + 3D Orbiting Carousel + UI Mockup + Intra-Slide Native Animations)
"""

import os
import pytest

from pmg import (
    MotionPresetRegistry,
    CinematicMotionBuilder,
    build_v1_3_cinematic_environmental_deck,
    StaticValidator,
    extract_expected_timeline,
)


def test_motion_preset_registry_and_choreography_decisions():
    assert "cinematic_title_reveal" in MotionPresetRegistry.PRESETS
    assert "multi_plane_parallax" in MotionPresetRegistry.PRESETS
    assert "camera_zoom_inspection" in MotionPresetRegistry.PRESETS
    assert "animated_infographic_bars" in MotionPresetRegistry.PRESETS
    assert "orbiting_carousel_3d" in MotionPresetRegistry.PRESETS
    assert "animated_interface_mockup" in MotionPresetRegistry.PRESETS

    builder = CinematicMotionBuilder()
    builder.record_decision(
        slide_number=1,
        object_id="!!Traveling_Kite",
        scope="inter_slide_morph",
        motion_preset="multi_plane_parallax",
        narrative_purpose="Guide viewer attention across landscape horizon.",
        trajectory_desc="x=5.9->6.6, y=0.65->0.95",
        duration_ms=1400,
        execution_mode="simultaneous",
    )
    assert len(builder.choreography_log) == 1
    assert builder.choreography_log[0].to_dict()["motion_preset"] == "multi_plane_parallax"


def test_v1_3_cinematic_10_slide_environmental_presentation(tmp_path):
    out_pptx = str(tmp_path / "v1_3_cinematic_env.pptx")
    build_v1_3_cinematic_environmental_deck(out_pptx, num_slides=10)
    assert os.path.exists(out_pptx)

    validator = StaticValidator()
    val = validator.validate(out_pptx)
    assert val["is_valid"] is True
    assert val["total_slides"] == 10
    assert len(val["missing_transitions"]) == 0
    assert val["total_native_animations"] >= 50

    # Verify Earth, Meadow, Blue Sky, and Flying Kite across all 10 slides
    anchors = val["traveling_anchors"]
    for required_anchor in ("!!Traveling_Kite", "!!Earth_Globe", "!!Nature_Sky", "!!Meadow_Back", "!!Meadow_Front"):
        assert required_anchor in anchors
        assert anchors[required_anchor] == list(range(1, 11))

    # Every slide must have both Morph anchors AND intra-slide native animations
    for s_summary in val["slides"]:
        assert s_summary["editable_object_count"] >= 10
        assert s_summary["native_animation_count"] >= 4

    # Verify ExpectedTimeline captures Camera Zoom Sequence, Parallax, and Animated Infographics
    timeline = extract_expected_timeline(out_pptx, hold_duration_ms=1000)
    assert timeline.total_slides == 10

    # Check Camera Zoom Sequence on Slide 3 & Slide 7 (Earth diameter expands significantly)
    s1_earth = next(o for o in timeline.slides[0].objects if o.object_id == "!!Earth_Globe")
    s3_earth = next(o for o in timeline.slides[2].objects if o.object_id == "!!Earth_Globe" and o.trigger_type == "slide_transition")
    s7_earth = next(o for o in timeline.slides[6].objects if o.object_id == "!!Earth_Globe" and o.trigger_type == "slide_transition")
    assert s3_earth.final_size[0] > s1_earth.final_size[0] * 1.4
    assert s7_earth.final_size[0] > s1_earth.final_size[0] * 1.6

    # Check Animated Infographic on Slide 4 (bars + curved motion path trend node)
    s4_ids = {o.object_id for o in timeline.slides[3].objects}
    assert "CleanEnergy_Info_Bar_1" in s4_ids
    assert "CleanEnergy_Info_Bar_4" in s4_ids
    assert "CleanEnergy_Info_Trend_Node" in s4_ids

    # Check 3D Orbiting Carousel on Slide 5 & Telemetry UI Mockup on Slide 6
    s5_ids = {o.object_id for o in timeline.slides[4].objects}
    assert "Carousel_Node_1" in s5_ids
    assert "Carousel_Node_2" in s5_ids
    s6_ids = {o.object_id for o in timeline.slides[5].objects}
    assert "UI_Mockup_Shell" in s6_ids
    assert "UI_Radar_Spinner" in s6_ids
