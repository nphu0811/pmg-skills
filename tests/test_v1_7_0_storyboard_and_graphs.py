"""
Acceptance, Regression, and Integration Tests for PMG Skills v1.7.0:
- Canonical Scene Graph data model (pmg.scene_graph)
- Canonical Motion Graph data model (pmg.motion_graph)
- Storyboard Intelligence Engine (pmg-storyboard)
- AI Motion Director (pmg.motion_director)
- Acceptance Demo: 5-Slide Diverse Storyboard Presentation compiled and validated with clean PowerPoint COM open.
"""
import os
import pytest

from pmg.version import VERSION
from pmg.scene_graph import SceneGraph, SceneNode, Transform
from pmg.motion_graph import MotionGraph, AnimationAction
from pmg.storyboard import StoryboardEngine, StoryboardPlan
from pmg.motion_director import MotionDirector
from pmg.validator import StaticValidator, verify_clean_powerpoint_open
from scripts.motion_graph_kit import MotionDeck, Themes


def test_v1_7_version():
    assert VERSION in ("1.7.0", "1.8.0", "2.0.1", "2.1.0")


def test_scene_graph_world_transform_accumulation():
    sg = SceneGraph("scene_test")
    parent = SceneNode(
        node_id="ParentGroup",
        kind="group",
        transform=Transform(x=1.0, y=2.0, scale_x=2.0, scale_y=2.0, rotation_deg=45.0, opacity=0.8),
    )
    child = SceneNode(
        node_id="ChildRect",
        kind="rect",
        transform=Transform(x=0.5, y=0.5, width=2.0, height=1.0, scale_x=1.5, rotation_deg=15.0, opacity=0.5),
    )
    sg.add_node(parent)
    sg.add_node(child, parent_id="ParentGroup")

    wt = sg.get_world_transform("ChildRect")
    assert wt.x == 2.0  # 1.0 + 0.5 * 2.0
    assert wt.y == 3.0  # 2.0 + 0.5 * 2.0
    assert wt.scale_x == 3.0  # 2.0 * 1.5
    assert wt.rotation_deg == 60.0  # 45.0 + 15.0
    assert wt.opacity == 0.4  # 0.8 * 0.5


def test_scene_graph_cycle_detection():
    sg = SceneGraph("cycle_test")
    n1 = SceneNode(node_id="Node1", kind="group")
    n2 = SceneNode(node_id="Node2", kind="group")
    sg.add_node(n1)
    sg.add_node(n2, parent_id="Node1")
    # Force cycle
    sg.nodes["Node1"].parent_id = "Node2"
    sg.nodes["Node2"].children.append("Node1")

    errors = sg.validate_graph()
    assert any("Cycle detected" in err for err in errors)


def test_scene_graph_intentional_overlap_vs_readable_collision():
    sg = SceneGraph("overlap_test")
    # 1. Intentional overlap (allowed)
    bg_card = SceneNode(node_id="BGCard", kind="rect", transform=Transform(x=1.0, y=1.0, width=5.0, height=4.0),
                        overlap_policy="allow_overlap")
    overlay = SceneNode(node_id="Overlay", kind="rect", transform=Transform(x=2.0, y=2.0, width=3.0, height=2.0),
                        overlap_policy="allow_overlap")
    sg.add_node(bg_card)
    sg.add_node(overlay)
    assert len(sg.validate_graph()) == 0

    # 2. Protected readable text collision (unintended)
    t1 = SceneNode(node_id="Text1", kind="textbox", text="Header One",
                   transform=Transform(x=1.0, y=1.0, width=3.0, height=1.0),
                   overlap_policy="protected_readable")
    t2 = SceneNode(node_id="Text2", kind="textbox", text="Header Two",
                   transform=Transform(x=2.0, y=1.5, width=3.0, height=1.0),
                   overlap_policy="protected_readable")
    sg.add_node(t1)
    sg.add_node(t2)
    errors = sg.validate_graph()
    assert any("Accidental collision between protected readable" in err for err in errors)


def test_motion_graph_dependencies_and_topological_compilation():
    mg = MotionGraph("motion_test")
    a1 = AnimationAction("act_1", "Node1", "zoom", duration_ms=500)
    a2 = AnimationAction("act_2", "Node2", "fly_in", duration_ms=600)
    a3 = AnimationAction("act_3", "Node3", "fade", duration_ms=400)

    mg.add_action(a1)
    mg.add_action(a2)
    mg.add_action(a3)

    mg.add_dependency("act_1", "act_2", relationship="after_previous", lag_ms=100)
    mg.add_dependency("act_2", "act_3", relationship="with_previous", lag_ms=50)

    # Compile with dummy shape mapping
    class DummyShape:
        def __init__(self, spid, name):
            self.shape_id = spid
            self.name = name

    dummy_mapping = {
        "Node1": DummyShape(10, "Node1"),
        "Node2": DummyShape(20, "Node2"),
        "Node3": DummyShape(30, "Node3"),
    }

    timeline = mg.compile_to_timeline(slide_number=1, shape_mapping=dummy_mapping)
    assert len(timeline.steps) == 3
    assert timeline.steps[0].target_spid == 10
    assert timeline.steps[1].target_spid == 20
    assert timeline.steps[1].trigger == "after_previous"
    assert timeline.steps[2].target_spid == 30
    assert timeline.steps[2].trigger == "with_previous"


def test_storyboard_engine_plan_diversity():
    engine = StoryboardEngine()
    plan = engine.plan_presentation("Quantum Computing Exploration", num_slides=6)
    assert len(plan.scenes) == 6

    # Verify diverse composition styles across scenes
    comp_styles = {s.composition_style for s in plan.scenes}
    assert len(comp_styles) >= 4
    assert "cinematic_hero" in comp_styles
    assert "editorial_layout" in comp_styles
    assert "radial_infographic" in comp_styles


def test_motion_director_acceptance_presentation(tmp_path):
    engine = StoryboardEngine()
    plan = engine.plan_presentation("Interstellar Navigation Systems", num_slides=5)

    director = MotionDirector(Themes.LIGHT_ACADEMIC_CYAN)
    output_pptx = str(tmp_path / "interstellar_director_test.pptx")
    deck_path = director.compile_plan_to_deck(plan, output_pptx)

    assert os.path.exists(deck_path)
    val = StaticValidator().validate(deck_path)
    assert val["is_valid"] is True, f"Validation defects: {val.get('defects')}"
    assert len(val["defects"]) == 0
    assert val["total_slides"] == 5
    assert val["total_native_animations"] > 10

    # Verify traveling hero orb across slides
    anchors = val.get("traveling_anchors", {})
    assert any("Traveling_Hero" in name for name in anchors)

    # Strict clean open check via PowerPoint COM
    clean = verify_clean_powerpoint_open(deck_path)
    assert clean["is_clean"] is True, f"Clean open failed: {clean.get('error')}"
