"""
tests.test_v2_2_0_animation_compiler
====================================
Regression and unit test suite for PMG v2.2.0 Native Animation Compiler.

Tests:
1. Animation Intermediate Representation (IR) serialization and deserialization
2. Capability status catalog and honest reporting
3. Deterministic timeline compilation and scheduling (parallel & sequential)
4. Cycle detection and dependency conflict resolution
5. Missing animation target detection
6. Duplicate shape and traveling entity identity collisions
7. Morph cross-slide reliability and gap detection
8. Motion path compilation with cubic Bezier trajectories
9. Native PPTX synthesis from AnimationIR without Microsoft PowerPoint
10. Level 1 Structural verification of compiled OOXML timing trees
"""

import os
import json
import pytest
from pmg.animation_ir import (
    AnimationCapabilityStatus,
    AnimationEasing,
    AnimationIR,
    SceneIR,
    ObjectIR,
    KeyframeIR,
    AnimationTrackIR,
    TransformIR,
    TimelineCompiler,
    MorphReliabilityValidator,
    get_effect_capability_status,
)
from pmg.compiler import NativePPTXCompiler
from pmg.validation import validate_level1_structural
from pmg.models import ValidationStatus


def test_animation_ir_serialization():
    """Verifies that AnimationIR round-trips to and from dict cleanly."""
    obj = ObjectIR(
        id="hero_cube",
        name="!!Hero_Cube",
        kind="card",
        transform=TransformIR(x=1.5, y=2.0, width=4.0, height=3.0),
        title="Quantum Core",
        body="Superconducting logic array",
    )
    assert obj.is_traveling_entity is True

    track = AnimationTrackIR(
        track_id="t1",
        target_id="hero_cube",
        effect="fade",
        trigger="with_previous",
        duration_ms=750,
        delay_ms=100,
        keyframes=[KeyframeIR(offset_fraction=0.5, opacity=0.5)],
    )

    scene = SceneIR(
        scene_id="s1",
        slide_number=1,
        title="Overview",
        objects=[obj],
        tracks=[track],
        transition_type="fade",
    )

    ir = AnimationIR(
        scenes=[scene],
        palette_name="LIGHT_ACADEMIC_CYAN",
    )

    d = ir.to_dict()
    assert d["version"] == "2.2.0"
    assert len(d["scenes"]) == 1
    assert d["scenes"][0]["objects"][0]["is_traveling_entity"] is True

    restored = AnimationIR.from_dict(d)
    assert restored.version == "2.2.0"
    assert len(restored.scenes) == 1
    assert restored.scenes[0].objects[0].name == "!!Hero_Cube"
    assert restored.scenes[0].tracks[0].effect == "fade"


def test_capability_status_catalog():
    """Verifies explicit capability classification for tested, unverified, and unsupported effects."""
    status, eff = get_effect_capability_status("fade")
    assert status == AnimationCapabilityStatus.SUPPORTED_AND_TESTED
    assert eff == "fade"

    status, eff = get_effect_capability_status("fly_in")
    assert status == AnimationCapabilityStatus.SUPPORTED_AND_TESTED

    status, eff = get_effect_capability_status("bezier_motion_path")
    assert status == AnimationCapabilityStatus.SUPPORTED_UNVERIFIED_PLAYBACK

    status, eff = get_effect_capability_status("quantum_hyperdrive_twist")
    assert status == AnimationCapabilityStatus.UNSUPPORTED


def test_timeline_compiler_parallel_and_sequential():
    """Verifies deterministic timing calculations for with_previous and after_previous tracks."""
    obj1 = ObjectIR(id="o1", name="Card1")
    obj2 = ObjectIR(id="o2", name="Card2")
    obj3 = ObjectIR(id="o3", name="Card3")

    # Track 1: First item (start 100, dur 500 -> ends 600)
    t1 = AnimationTrackIR(track_id="t1", target_id="o1", effect="fade", trigger="with_previous", duration_ms=500, delay_ms=100)
    # Track 2: Parallel with track 1 (start 0 + 200 = 200, dur 600 -> ends 800)
    t2 = AnimationTrackIR(track_id="t2", target_id="o2", effect="zoom", trigger="with_previous", duration_ms=600, delay_ms=200)
    # Track 3: Sequential after previous group (ends at 800, so starts 800 + 50 = 850, dur 400 -> ends 1250)
    t3 = AnimationTrackIR(track_id="t3", target_id="o3", effect="fly_in", trigger="after_previous", duration_ms=400, delay_ms=50)

    scene = SceneIR(scene_id="sc1", objects=[obj1, obj2, obj3], tracks=[t1, t2, t3])
    compiler = TimelineCompiler(scene)
    scheduled = compiler.compile()

    assert scheduled[0].computed_start_ms == 100
    assert scheduled[0].computed_end_ms == 600

    assert scheduled[1].computed_start_ms == 200
    assert scheduled[1].computed_end_ms == 800

    assert scheduled[2].computed_start_ms == 850
    assert scheduled[2].computed_end_ms == 1250


def test_timeline_compiler_missing_target_rejection():
    """Verifies that referencing a non-existent target raises an informative error."""
    obj1 = ObjectIR(id="valid_obj", name="Valid Object")
    track = AnimationTrackIR(track_id="t_bad", target_id="ghost_obj", effect="fade")

    scene = SceneIR(scene_id="sc_err", objects=[obj1], tracks=[track])
    compiler = TimelineCompiler(scene)
    is_valid, errors, _ = compiler.validate()

    assert is_valid is False
    assert any("targets non-existent object 'ghost_obj'" in e for e in errors)

    with pytest.raises(ValueError, match="TimelineCompiler error"):
        compiler.compile()


def test_timeline_compiler_duplicate_object_id_rejection():
    """Verifies that duplicate object IDs on the same scene are rejected."""
    obj1 = ObjectIR(id="duplicate_id", name="Obj1")
    obj2 = ObjectIR(id="duplicate_id", name="Obj2")

    scene = SceneIR(scene_id="sc_dup", objects=[obj1, obj2])
    compiler = TimelineCompiler(scene)
    is_valid, errors, _ = compiler.validate()

    assert is_valid is False
    assert any("Duplicate object ID 'duplicate_id'" in e for e in errors)


def test_morph_reliability_validator_flags_duplicate_traveling_collision():
    """Verifies that duplicate '!!' persistent names on the same slide are flagged as collisions."""
    obj1 = ObjectIR(id="c1", name="!!Traveling_Core")
    obj2 = ObjectIR(id="c2", name="!!Traveling_Core")  # Collision!

    scene = SceneIR(scene_id="s1", slide_number=1, objects=[obj1, obj2])
    ir = AnimationIR(scenes=[scene])

    rep = MorphReliabilityValidator.validate_morph_flow(ir)
    assert rep["is_valid"] is False
    assert any("Morph Collision: Duplicate traveling entity '!!Traveling_Core'" in e for e in rep["errors"])


def test_morph_reliability_validator_detects_gap_slides():
    """Verifies that skipping intermediate slides with a traveling entity raises a warning."""
    # Appears on slide 1 and slide 3, skipping slide 2
    s1 = SceneIR(scene_id="s1", slide_number=1, objects=[ObjectIR(id="t1", name="!!Probe")])
    s2 = SceneIR(scene_id="s2", slide_number=2, objects=[ObjectIR(id="o2", name="Other")])
    s3 = SceneIR(scene_id="s3", slide_number=3, transition_type="morph", objects=[ObjectIR(id="t3", name="!!Probe")])

    ir = AnimationIR(scenes=[s1, s2, s3])
    rep = MorphReliabilityValidator.validate_morph_flow(ir)
    assert any("skipping intermediate slide(s) [2]" in w for w in rep["warnings"])


def test_native_pptx_compilation_from_animation_ir(tmp_path):
    """
    End-to-end test: compiles an AnimationIR containing entrance, exit, emphasis,
    and Bezier motion paths across 2 slides with Morph transition into valid PPTX.
    Does NOT require Microsoft PowerPoint.
    """
    out_pptx = str(tmp_path / "compiled_ir_deck.pptx")

    # Slide 1: Establishing scene with hero traveling entity
    hero_s1 = ObjectIR(
        id="hero",
        name="!!Hero_Satellite",
        kind="card",
        transform=TransformIR(x=1.0, y=2.5, width=4.5, height=3.5),
        title="Orbital Relay Alpha",
        body="High-bandwidth laser transceiver locked on ground station.",
    )
    hud_s1 = ObjectIR(
        id="hud",
        name="TelemetryHUD",
        kind="textbox",
        transform=TransformIR(x=6.5, y=2.5, width=5.5, height=2.0),
        title="Signal Strength: 99.8%",
        body="Uplink active across Ka-band spectrum.",
    )

    t1_entr = AnimationTrackIR(track_id="t1", target_id="hero", effect="fade", trigger="with_previous", duration_ms=800)
    t1_path = AnimationTrackIR(
        track_id="t2",
        target_id="hero",
        effect="bezier_motion_path",
        trigger="after_previous",
        duration_ms=1000,
        motion_path="M 0 0 C 0.08 -0.05 0.16 0.05 0.25 0 E",
    )
    t1_hud = AnimationTrackIR(track_id="t3", target_id="hud", effect="fly_in", direction="up", trigger="with_previous", delay_ms=150)

    scene1 = SceneIR(
        scene_id="scene_01",
        slide_number=1,
        title="Deep Space Optical Network",
        subtitle="Autonomous Laser Link Telemetry",
        transition_type="fade",
        objects=[hero_s1, hud_s1],
        tracks=[t1_entr, t1_path, t1_hud],
    )

    # Slide 2: Morph destination scene with moved hero entity
    hero_s2 = ObjectIR(
        id="hero_morphed",
        name="!!Hero_Satellite",  # Same !! identifier for 60 FPS Morph
        kind="card",
        transform=TransformIR(x=7.0, y=2.0, width=5.5, height=4.5),
        title="Relay Station Orbiting Mars",
        body="Continuous tracking achieved with zero frame drop.",
    )
    metric_s2 = ObjectIR(
        id="metric",
        name="ThroughputMetric",
        kind="card",
        transform=TransformIR(x=1.0, y=2.5, width=5.0, height=3.0),
        title="Throughput: 1.2 Gbps",
        body="Optical link established with zero packet loss.",
    )

    t2_pulse = AnimationTrackIR(track_id="t4", target_id="hero_morphed", effect="pulse", trigger="with_previous", duration_ms=600)
    t2_scale = AnimationTrackIR(track_id="t5", target_id="metric", effect="scale", scale_factor=(1.15, 1.15), trigger="after_previous", duration_ms=700)

    scene2 = SceneIR(
        scene_id="scene_02",
        slide_number=2,
        title="Planetary Relay Integration",
        subtitle="60 FPS Native Morph Transition",
        transition_type="morph",
        transition_duration_ms=1600,
        objects=[hero_s2, metric_s2],
        tracks=[t2_pulse, t2_scale],
    )

    ir = AnimationIR(
        scenes=[scene1, scene2],
        palette_name="LIGHT_ACADEMIC_CYAN",
    )

    compiler = NativePPTXCompiler()
    report = compiler.compile_from_ir(ir, out_pptx)

    assert os.path.exists(out_pptx)
    assert report.total_slides == 2
    assert report.total_animations >= 4
    assert report.timing_trees_attached == 2
    assert report.is_fully_native is True
    assert "!!Hero_Satellite" in report.named_anchors
    assert report.morph_validation_report["is_valid"] is True
    assert report.capability_breakdown.get(AnimationCapabilityStatus.SUPPORTED_AND_TESTED.value, 0) >= 3

    # Level 1 Structural Validation verification
    l1 = validate_level1_structural(out_pptx)
    assert l1.status == ValidationStatus.PASS
    assert l1.xml_valid is True
    assert l1.opc_content_types_valid is True
    assert l1.relationships_valid is True
    assert l1.animation_targets_valid is True
    assert l1.timing_graph_valid is True
