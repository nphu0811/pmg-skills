"""Acceptance, Regression, and Integration Tests for PMG Skills v2.0.1:
Creative Motion Intelligence & Autonomous Motion Director 2.0.
"""

import os
from pathlib import Path
import pytest
import pptx

from pmg.version import VERSION
from pmg.creative_director import (
    CreativeEvaluationAxis,
    CreativeConcept,
    CreativeCritique,
    ModelPlanningAdapter,
    CreativeDirector,
    build_v2_0_oneshot_space_scene,
    build_v2_0_agribank_oneshot_deck,
    build_v2_0_creative_concept_selection_deck,
)
from pmg.storyboard import StoryboardEngine
from pmg.animations import parse_slide_timing_openxml
from pmg.validator import StaticValidator, verify_clean_powerpoint_open


def test_v2_0_version():
    assert VERSION == "2.0.1"


def test_creative_concept_generation_and_diversity():
    director = CreativeDirector()
    concepts = director.generate_creative_concepts(
        topic="Neural Quantum Networks",
        audience="academic_keynote",
        tone="cinematic",
        num_concepts=3,
    )
    assert len(concepts) == 3
    ids = [c.concept_id for c in concepts]
    assert len(set(ids)) == 3

    # Verify all 6 axes are present and scored for every candidate
    required_axes = {
        CreativeEvaluationAxis.CLARITY,
        CreativeEvaluationAxis.ORIGINALITY,
        CreativeEvaluationAxis.VISUAL_HIERARCHY,
        CreativeEvaluationAxis.CINEMATIC_COHERENCE,
        CreativeEvaluationAxis.RHYTHM_AND_PACING,
        CreativeEvaluationAxis.NATIVE_EDITABILITY,
    }
    for c in concepts:
        assert set(c.evaluation_scores.keys()) == required_axes
        assert c.total_score > 0.0
        assert len(c.narrative_angle) > 0
        assert len(c.motion_grammar) > 0
        assert len(c.visual_metaphor) > 0


def test_creative_concept_evaluation_and_selection():
    director = CreativeDirector()
    concepts = director.generate_creative_concepts(topic="Autonomous Mars Rover", num_concepts=3)
    winner, matrix = director.evaluate_and_select_concept(concepts)

    assert isinstance(winner, CreativeConcept)
    assert winner.concept_id == matrix["selected_concept_id"]
    assert len(matrix["candidates"]) == 3
    assert "selection_reason" in matrix
    assert winner.total_score >= min(c.total_score for c in concepts)


def test_bounded_creative_critique_and_refinement():
    director = CreativeDirector()
    concepts = director.generate_creative_concepts(topic="Renewable Fusion Energy", num_concepts=3)
    winner, _ = director.evaluate_and_select_concept(concepts)

    engine = StoryboardEngine()
    initial_plan = engine.plan_presentation(topic="Renewable Fusion Energy", num_slides=3, tone="cinematic")
    # Artificially compress pacing to test refinement triggering
    for s in initial_plan.scenes:
        s.pacing_seconds = 3.0

    refined_plan, critiques = director.critique_and_refine(winner, initial_plan, max_iterations=2)
    assert len(critiques) == 2
    assert all(isinstance(cr, CreativeCritique) for cr in critiques)
    assert critiques[0].is_approved is True
    # Verify refinements expanded durations
    assert sum(s.pacing_seconds for s in refined_plan.scenes) > 9.0


def test_acceptance_demo1_oneshot_space_scene(tmp_path):
    out_pptx = str(tmp_path / "v2_0_oneshot_space.pptx")
    build_v2_0_oneshot_space_scene(out_pptx, duration=25.0)

    # 1. Static OpenXML validation
    val = StaticValidator().validate(out_pptx)
    assert val["is_valid"] is True, f"Validation defects: {val['defects']}"
    assert val["total_slides"] == 1

    # 2. Inspect OpenXML timing tree & duration
    prs = pptx.Presentation(out_pptx)
    specs, clicks = parse_slide_timing_openxml(prs.slides[0], slide_number=1)
    assert len(specs) >= 15
    # Verify continuous long timeline duration >= 20 seconds (20,000 ms)
    max_duration_ms = max(s.absolute_end_ms for s in specs)
    assert max_duration_ms >= 20000, f"Expected duration >= 20,000 ms, got {max_duration_ms}"
    # Verify presence of multiple simultaneous effect classes: motion paths, rotations, scales
    assert any(s.motion_path is not None for s in specs)
    assert any(s.rotation_params is not None for s in specs)

    # 3. Clean PowerPoint Desktop COM open
    clean = verify_clean_powerpoint_open(out_pptx)
    assert clean["is_clean"] is True, f"Clean open failed: {clean}"


def test_acceptance_demo2_agribank_oneshot_deck(tmp_path):
    out_pptx = str(tmp_path / "v2_0_agribank_oneshot.pptx")
    build_v2_0_agribank_oneshot_deck(out_pptx)

    # 1. Static OpenXML validation
    val = StaticValidator().validate(out_pptx)
    assert val["is_valid"] is True, f"Validation defects: {val['defects']}"
    assert val["total_slides"] == 1

    # 2. Inspect shape hierarchy for native tiles and typography
    prs = pptx.Presentation(out_pptx)
    slide = prs.slides[0]
    shape_names = [s.name for s in slide.shapes]
    assert any("Agribank_Tile" in n for n in shape_names)
    assert any("Wheat_Grain" in n for n in shape_names)
    assert any("Pillar_Card" in n for n in shape_names)

    # 3. Clean PowerPoint Desktop COM open
    clean = verify_clean_powerpoint_open(out_pptx)
    assert clean["is_clean"] is True, f"Clean open failed: {clean}"


def test_acceptance_demo3_creative_concept_selection_deck(tmp_path):
    out_pptx = str(tmp_path / "v2_0_concept_selection.pptx")
    res = build_v2_0_creative_concept_selection_deck("Artificial Intelligence & Robotics", out_pptx)

    assert "selected" in res
    assert "selection_matrix" in res
    assert len(res["concepts"]) == 3
    assert len(res["critiques"]) >= 1

    # 1. Static OpenXML validation
    val = StaticValidator().validate(out_pptx)
    assert val["is_valid"] is True, f"Validation defects: {val['defects']}"
    assert val["total_slides"] == 4

    # 2. Clean PowerPoint Desktop COM open
    clean = verify_clean_powerpoint_open(out_pptx)
    assert clean["is_clean"] is True, f"Clean open failed: {clean}"
