"""
tests.test_v2_5_0_object_motion_intelligence
=============================================
Comprehensive unit and integration test suite for PMG v2.5.0 Universal Object-Specific Motion Intelligence:
- Semantic object classification
- Specialized planners (Chart, Kinetic Typography, Image, Diagram, Map, Infographic)
- Explainable Motion Decision System (8 core cognitive/pedagogical questions)
- Motion Quality, Accessibility & Reduced Motion Validation
- Cross-Object Choreography & Scene Graph Compilation
"""

import pytest
from pptx.util import Inches, Pt

from pmg.semantic_classifier import (
    SemanticClassifier,
    ObjectType,
    SemanticRole,
    NarrativeImportance,
    AudienceFocus,
    ObjectSemanticProfile,
)
from pmg.motion_planners import (
    ChartMotionPlanner,
    KineticTypographyPlanner,
    ImageMotionPlanner,
    DiagramMotionPlanner,
    MapMotionPlanner,
    InfographicMotionPlanner,
    MotionIntent,
    SharedMotionPrimitive,
)
from pmg.motion_decision import (
    MotionDecision,
    MotionDecisionReport,
    ExplainableDecisionEngine,
)
from pmg.motion_quality import (
    MotionQualityValidator,
    MotionQualityReport,
    ReducedMotionProfile,
)
from pmg.scene_graph import SceneGraph, SceneNode, Transform
from pmg.motion_director import MotionDirector
from pmg.version import VERSION


def test_v2_5_0_version_check():
    assert VERSION == "2.5.0"


def test_semantic_classifier_categories():
    # 1. Traveling Hero
    hero = SemanticClassifier.classify({"id": "HeroOrb", "name": "!!Hero_Orb", "kind": "ellipse"})
    assert hero.object_type == ObjectType.TRAVELING_HERO
    assert hero.narrative_importance == NarrativeImportance.HERO
    assert hero.audience_focus == AudienceFocus.CONTINUOUS

    # 2. Typography: Title vs Subtitle vs Body
    title = SemanticClassifier.classify({"id": "Title", "name": "Title_1", "kind": "textbox", "text": "Deep Learning", "font_size_pt": 32.0})
    assert title.object_type == ObjectType.TYPOGRAPHY
    assert title.semantic_role == SemanticRole.TITLE
    assert title.visual_hierarchy_level == 1

    sub = SemanticClassifier.classify({"id": "Sub", "name": "Sub_1", "kind": "textbox", "text": "Overview", "font_size_pt": 18.0})
    assert sub.semantic_role == SemanticRole.SUBTITLE

    # 3. Chart
    chart_bar = SemanticClassifier.classify({"id": "Bar_Q1", "name": "bar_q1", "kind": "rect", "metadata": {"type": "bar_chart"}})
    assert chart_bar.object_type == ObjectType.CHART
    assert chart_bar.semantic_role == SemanticRole.DATA_BAR

    # 4. Diagram
    node = SemanticClassifier.classify({"id": "Step1", "name": "node_ingest", "kind": "rect", "metadata": {"type": "diagram_node"}})
    assert node.object_type == ObjectType.DIAGRAM
    assert node.semantic_role == SemanticRole.PROCESS_STEP

    # 5. Map
    route = SemanticClassifier.classify({"id": "Route1", "name": "route_hn_dn", "kind": "line", "metadata": {"type": "geo_route"}})
    assert route.object_type == ObjectType.MAP
    assert route.semantic_role == SemanticRole.GEO_ROUTE

    # 6. Infographic
    timeline = SemanticClassifier.classify({"id": "Timeline", "name": "timeline_spine", "kind": "line", "metadata": {"type": "timeline"}})
    assert timeline.object_type == ObjectType.INFOGRAPHIC
    assert timeline.semantic_role == SemanticRole.TIMELINE_SPINE

    # 7. Image
    img = SemanticClassifier.classify({"id": "Photo", "name": "img_subject", "kind": "picture"})
    assert img.object_type == ObjectType.IMAGE
    assert img.semantic_role == SemanticRole.HERO_SUBJECT


def test_chart_motion_planner_patterns():
    planner = ChartMotionPlanner()
    intents = planner.plan_motion({
        "chart_type": "bar",
        "axis_id": "axis_main",
        "items": [
            {"id": "b1", "value": "100", "label_id": "lbl1"},
            {"id": "b2", "value": "250", "label_id": "lbl2", "is_peak": True},
        ]
    })
    # Must have: axis DRAW -> b1 GROW -> lbl1 COUNT -> b2 GROW -> lbl2 COUNT -> b2 EMPHASIZE
    assert len(intents) == 6
    assert intents[0].target_id == "axis_main"
    assert intents[0].primitive == SharedMotionPrimitive.DRAW
    assert intents[1].target_id == "b1"
    assert intents[1].primitive == SharedMotionPrimitive.GROW
    assert intents[1].direction == "up"
    assert intents[2].target_id == "lbl1"
    assert intents[2].primitive == SharedMotionPrimitive.COUNT
    assert intents[5].target_id == "b2"
    assert intents[5].primitive == SharedMotionPrimitive.EMPHASIZE
    assert intents[5].scale_factor > 1.0


def test_kinetic_typography_planner_patterns():
    planner = KineticTypographyPlanner()
    
    # Title Unveiling
    intents = planner.plan_motion({
        "style": "title_unveiling",
        "target_id": "title_head",
        "subtitle_id": "title_sub",
    })
    assert len(intents) == 2
    assert intents[0].primitive == SharedMotionPrimitive.REVEAL
    assert intents[1].primitive == SharedMotionPrimitive.REVEAL

    # Keyword Punch
    kw_intents = planner.plan_motion({
        "style": "keyword_punch",
        "target_id": "phrase",
        "keyword_ids": ["kw_deep", "kw_neural"],
    })
    assert len(kw_intents) == 3
    assert kw_intents[1].primitive == SharedMotionPrimitive.EMPHASIZE
    assert kw_intents[1].scale_factor >= 1.14


def test_image_motion_planner_patterns():
    planner = ImageMotionPlanner()

    # Layered Parallax
    intents = planner.plan_motion({
        "pattern": "layered_parallax",
        "background_id": "bg_layer",
        "target_id": "mid_subject",
        "foreground_id": "fg_scrim",
    })
    assert len(intents) == 3
    assert intents[0].target_id == "bg_layer"
    assert intents[0].parameters.get("speed_factor") == 0.7
    assert intents[1].target_id == "mid_subject"
    assert intents[1].parameters.get("speed_factor") == 1.0
    assert intents[2].target_id == "fg_scrim"
    assert intents[2].parameters.get("speed_factor") == 1.4

    # Controlled Macro Zoom clamp
    zoom_intents = planner.plan_motion({
        "pattern": "controlled_zoom",
        "target_id": "mid_subject",
        "scale_factor": 1.50,  # Exceeds max
    })
    assert len(zoom_intents) == 1
    assert zoom_intents[0].scale_factor <= 1.25  # Clamped


def test_diagram_motion_planner_causal_topology():
    planner = DiagramMotionPlanner()
    intents = planner.plan_motion({
        "pattern": "process_flow",
        "steps": [{"id": "step_a"}, {"id": "step_b"}, {"id": "step_c"}],
        "connectors": [
            {"id": "conn_ab", "source": "step_a", "direction": "right"},
            {"id": "conn_bc", "source": "step_b", "direction": "right"},
        ]
    })
    # Order must strictly be: step_a -> conn_ab -> step_b -> conn_bc -> step_c
    expected_order = ["step_a", "conn_ab", "step_b", "conn_bc", "step_c"]
    assert [i.target_id for i in intents] == expected_order
    assert intents[1].primitive == SharedMotionPrimitive.DRAW
    assert intents[1].direction == "right"


def test_map_motion_planner_patterns():
    planner = MapMotionPlanner()
    intents = planner.plan_motion({
        "pattern": "route_progression",
        "base_map_id": "vietnam_map",
        "locations": [
            {"id": "pin_hanoi", "label": "Hanoi"},
            {"id": "pin_hcmc", "label": "HCMC", "pulse_id": "pulse_hcmc"},
        ],
        "routes": [{"id": "route_hn_hcm", "direction": "down"}],
    })
    assert intents[0].target_id == "vietnam_map"
    assert intents[1].target_id == "pin_hanoi"
    assert intents[1].primitive == SharedMotionPrimitive.MOVE
    assert intents[1].direction == "down"
    assert intents[2].target_id == "route_hn_hcm"
    assert intents[2].primitive == SharedMotionPrimitive.DRAW
    assert intents[3].target_id == "pin_hcmc"
    assert intents[4].target_id == "pulse_hcmc"
    assert intents[4].primitive == SharedMotionPrimitive.HIGHLIGHT


def test_infographic_motion_planner_timeline():
    planner = InfographicMotionPlanner()
    intents = planner.plan_motion({
        "pattern": "timeline_progression",
        "spine_id": "time_spine",
        "milestones": [
            {"node_id": "m1", "card_id": "c1", "year": "2024"},
            {"node_id": "m2", "card_id": "c2", "year": "2025"},
        ]
    })
    # spine -> m1 -> c1 -> m2 -> c2
    assert intents[0].target_id == "time_spine"
    assert intents[0].primitive == SharedMotionPrimitive.DRAW
    assert intents[1].target_id == "m1"
    assert intents[2].target_id == "c1"
    assert intents[3].target_id == "m2"
    assert intents[4].target_id == "c2"


def test_explainable_motion_decision_system():
    profile = ObjectSemanticProfile(
        object_id="bar_revenue",
        object_type=ObjectType.CHART,
        semantic_role=SemanticRole.DATA_BAR,
        narrative_importance=NarrativeImportance.PRIMARY,
        data_context={"series_name": "Q3 Revenue"},
    )
    intent = MotionIntent(
        target_id="bar_revenue",
        primitive=SharedMotionPrimitive.GROW,
        semantic_purpose="Bar grows upward to quantitative target",
        duration_ms=750,
        direction="up",
    )
    decision = ExplainableDecisionEngine.evaluate(profile, intent)
    assert decision.target_id == "bar_revenue"
    assert decision.shared_primitive == SharedMotionPrimitive.GROW
    assert len(decision.answers) == 8
    assert "Chart" in decision.answers["1. What is this object?"]
    assert "Quantitative" in decision.answers["2. What information does it communicate?"]
    assert "native" in decision.answers["7. Native compilation status?"].lower()


def test_motion_quality_validator_and_reduced_motion():
    # 1. Normal valid intents
    valid_intents = [
        MotionIntent(target_id="n1", primitive=SharedMotionPrimitive.REVEAL, semantic_purpose="Reveal 1", duration_ms=600),
        MotionIntent(target_id="n2", primitive=SharedMotionPrimitive.REVEAL, semantic_purpose="Reveal 2", duration_ms=600),
    ]
    report = MotionQualityValidator.validate(valid_intents, known_target_ids={"n1", "n2"})
    assert report.passed is True
    assert report.score == 100.0

    # 2. Missing target detection
    bad_report = MotionQualityValidator.validate(valid_intents, known_target_ids={"n1"})  # n2 missing
    assert bad_report.passed is False
    assert any("non-existent" in err for err in bad_report.errors)

    # 3. Excessive simultaneous actions warning (>4)
    heavy_intents = [
        MotionIntent(target_id=f"item_{k}", primitive=SharedMotionPrimitive.REVEAL, semantic_purpose="test", duration_ms=500, trigger="with_previous")
        for k in range(6)
    ]
    heavy_report = MotionQualityValidator.validate(heavy_intents, known_target_ids={f"item_{k}" for k in range(6)})
    assert any("cognitive load" in w.lower() for w in heavy_report.warnings)

    # 4. Reduced Motion Profile softens primitives
    intense_intents = [
        MotionIntent(target_id="spin_item", primitive=SharedMotionPrimitive.ROTATE, semantic_purpose="Rapid Spin", duration_ms=2000),
        MotionIntent(target_id="zoom_item", primitive=SharedMotionPrimitive.SCALE, semantic_purpose="Zoom", duration_ms=1000),
    ]
    adapted = ReducedMotionProfile.apply(intense_intents)
    assert adapted[0].primitive == SharedMotionPrimitive.REVEAL
    assert adapted[1].primitive == SharedMotionPrimitive.REVEAL


def test_cross_object_choreography_end_to_end():
    sg = SceneGraph(scene_id="scene_test")
    # Add Title
    sg.add_node(SceneNode(node_id="Title", name="Title_Text", kind="textbox", text="Main Title", font_size_pt=32.0))
    # Add Bar
    sg.add_node(SceneNode(node_id="Bar", name="Bar_Q1", kind="rect", semantic_role="chart"))
    # Add Card
    sg.add_node(SceneNode(node_id="Card", name="GlassCard", kind="rect", semantic_role="card"))

    director = MotionDirector()
    mg, dec_report, qual_report = director.choreograph_scene(sg)

    assert len(mg.actions) >= 3
    assert len(dec_report.decisions) >= 3
    assert qual_report.passed is True
    assert dec_report.total_duration_ms > 0
    # Check explain text generation
    summary_text = dec_report.explain_text()
    assert "Motion Decision Report" in summary_text
    assert "Title" in summary_text
