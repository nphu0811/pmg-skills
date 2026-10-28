---
name: pmg-storyboard
description: "Storyboard intelligence and narrative presentation planner for PMG. Generates structured, diverse scene plans (cinematic hero, editorial, radial infographics, continuous one-shot, dashboards) with camera intentions, asset requirements, and pacing before constructing slides."
---

# PMG Storyboard Intelligence Engine

`pmg-storyboard` plans presentation narratives, scene diversity, camera movements, and motion strategies before individual slides are constructed.
Integrated via `pmg.storyboard.StoryboardEngine` and consumed by `pmg.motion_director.MotionDirector`.

## Core Responsibilities & Invariants

1. **Anti-Repetition & Scene Diversity**:
   - Rejects monotonous, repetitive two-card layouts across presentations.
   - Intentionally selects composition archetypes based on narrative purpose:
     - `cinematic_hero`: Grand keynote opening and hero focus.
     - `editorial_layout`: Clean asymmetrical editorial multi-column structure.
     - `radial_infographic`: Coordinated orbiting nodes, HUD wheels, and concentric rings.
     - `layered_cutaway`: Multi-plane technical exploded views and annotations.
     - `photo_carousel`: Dynamic photo gallery transformations via Morph.
     - `continuous_one_shot`: Extended intra-slide continuous animation sequence.

2. **Camera & Pacing Direction**:
   - Assigns narrative-driven virtual camera intentions (`wide_establishing`, `push_in_zoom`, `tracking_pan`, `macro_detail`, `pull_out`).
   - Pairs camera intentions with native transition psychology (Morph, Circle Wipe, Split Curtain, Fade).

3. **Asset & Motion Intent Allocation**:
   - Outlines visual asset requirements per scene for `pmg-visual-assets`.
   - Flags decomposition candidates for `pmg-image-decomposer`.
   - Provides high-level motion guidance to the AI Motion Director.

4. **Boundaries**:
   - Produces structured `StoryboardPlan` and `StoryboardScene` objects.
   - Does not implement OpenXML or compile animations directly.

## Python API Usage

```python
from pmg.storyboard import StoryboardEngine
from pmg.motion_director import MotionDirector

# 1. Generate diverse presentation storyboard
engine = StoryboardEngine()
plan = engine.plan_presentation(
    topic="Autonomous Space Systems",
    num_slides=5,
    audience="technical_keynote",
    tone="modern_cinematic",
)

# 2. Hand off plan to MotionDirector
director = MotionDirector()
output_pptx = director.compile_plan_to_deck(plan, "output/space_mission.pptx")
```
