"""
pmg.motion_director
===================
Universal AI Motion Director for PMG Skills (v2.5.0).

Upgraded with Universal Object-Specific Motion Intelligence:
- Semantic object classification across charts, typography, images, diagrams, maps, and infographics.
- 6 specialized object motion planners (`pmg-chart-motion`, `pmg-kinetic-typography`, `pmg-image-motion`,
  `pmg-diagram-motion`, `pmg-map-motion`, `pmg-infographic-motion`).
- Explainable Motion Decision System answering the 8 cognitive/pedagogical questions.
- Cross-object choreography enforcing narrative synchronization.
- Motion quality, accessibility, and performance validation with reduced-motion support.
"""

from typing import Dict, Any, List, Optional, Tuple, Set
from pathlib import Path

from pmg.storyboard import StoryboardPlan, StoryboardScene
from pmg.scene_graph import SceneGraph, SceneNode, Transform
from pmg.motion_graph import MotionGraph, AnimationAction, Dependency
from pmg.animations import SlideAnimationTimeline
from scripts.motion_graph_kit import MotionDeck, Themes

from pmg.semantic_classifier import (
    SemanticClassifier,
    ObjectType,
    SemanticRole,
    NarrativeImportance,
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


class MotionDirector:
    """
    Intelligent Motion Director coordinating visual composition, semantic classification,
    object-specific motion planning, cross-object choreography, and quality audits.
    """
    def __init__(self, theme=Themes.LIGHT_ACADEMIC_CYAN, reduced_motion: bool = False):
        self.theme = theme
        self.reduced_motion = reduced_motion
        
        # Specialized Object Motion Planners
        self.chart_planner = ChartMotionPlanner()
        self.typography_planner = KineticTypographyPlanner()
        self.image_planner = ImageMotionPlanner()
        self.diagram_planner = DiagramMotionPlanner()
        self.map_planner = MapMotionPlanner()
        self.infographic_planner = InfographicMotionPlanner()

    def classify_scene_objects(self, sg: SceneGraph) -> List[ObjectSemanticProfile]:
        """Classifies all nodes in a SceneGraph into semantic profiles."""
        profiles = []
        for idx, (node_id, node) in enumerate(sg.nodes.items()):
            profile = SemanticClassifier.classify(node, index=idx)
            profiles.append(profile)
        return profiles

    def choreograph_scene(
        self,
        sg: SceneGraph,
        reduced_motion: Optional[bool] = None,
    ) -> Tuple[MotionGraph, MotionDecisionReport, MotionQualityReport]:
        """
        Executes end-to-end universal motion choreography for a single SceneGraph.
        Returns:
          1. Native MotionGraph ready for OpenXML timeline compilation
          2. Explainable MotionDecisionReport detailing why each motion was selected
          3. MotionQualityReport verifying cognitive load and timing
        """
        use_reduced = self.reduced_motion if reduced_motion is None else reduced_motion
        mg = MotionGraph(graph_id=f"motion_{sg.scene_id}")
        profiles = self.classify_scene_objects(sg)
        profile_map = {p.object_id: p for p in profiles}

        all_intents: List[MotionIntent] = []
        decisions: List[MotionDecision] = []

        # Sort profiles by reading order and visual hierarchy
        sorted_profiles = sorted(
            profiles,
            key=lambda p: (
                0 if p.narrative_importance == NarrativeImportance.HERO else
                1 if p.semantic_role in (SemanticRole.TITLE, SemanticRole.DATA_AXIS, SemanticRole.TIMELINE_SPINE) else
                2 if p.narrative_importance == NarrativeImportance.PRIMARY else
                3 if p.narrative_importance == NarrativeImportance.SECONDARY else 4,
                p.reading_order,
            )
        )

        # Plan motion per object profile using specialized planners
        for prof in sorted_profiles:
            intents: List[MotionIntent] = []
            obj_type = prof.object_type

            if obj_type == ObjectType.TRAVELING_HERO:
                # Persistent cross-slide Morph hero
                intents.append(
                    MotionIntent(
                        target_id=prof.object_id,
                        primitive=SharedMotionPrimitive.FOCUS,
                        semantic_purpose="Traveling hero orb establishes continuous visual focal point",
                        duration_ms=750,
                        delay_ms=50,
                        trigger="after_previous",
                        scale_factor=1.08,
                        narrative_importance="hero",
                    )
                )

            elif obj_type == ObjectType.TYPOGRAPHY:
                style = "title_unveiling" if prof.semantic_role == SemanticRole.TITLE else "keyword_punch" if prof.semantic_role == SemanticRole.KEYWORD else "word_stagger"
                intents = self.typography_planner.plan_motion({
                    "style": style,
                    "target_id": prof.object_id,
                })

            elif obj_type == ObjectType.CHART:
                intents = self.chart_planner.plan_motion({
                    "chart_type": "bar",
                    "items": [{"id": prof.object_id, "role": prof.semantic_role.value}],
                })

            elif obj_type == ObjectType.DIAGRAM:
                intents = self.diagram_planner.plan_motion({
                    "pattern": "process_flow",
                    "steps": [{"id": prof.object_id, "label": prof.object_id}],
                })

            elif obj_type == ObjectType.MAP:
                intents = self.map_planner.plan_motion({
                    "pattern": "route_progression",
                    "locations": [{"id": prof.object_id}],
                })

            elif obj_type == ObjectType.INFOGRAPHIC:
                intents = self.infographic_planner.plan_motion({
                    "pattern": "timeline_progression",
                    "milestones": [{"node_id": prof.object_id}],
                })

            elif obj_type == ObjectType.IMAGE:
                intents = self.image_planner.plan_motion({
                    "pattern": "directional_reveal",
                    "target_id": prof.object_id,
                })

            else:
                # Decorative / Background Cards settle softly
                intents.append(
                    MotionIntent(
                        target_id=prof.object_id,
                        primitive=SharedMotionPrimitive.REVEAL,
                        semantic_purpose="Structural container card provides background framing",
                        duration_ms=400,
                        delay_ms=0,
                        trigger="with_previous" if all_intents else "after_previous",
                        narrative_importance="background",
                    )
                )

            for intent in intents:
                all_intents.append(intent)
                decision = ExplainableDecisionEngine.evaluate(prof, intent, profiles)
                decisions.append(decision)

        # Apply reduced motion profile if requested
        if use_reduced:
            all_intents = ReducedMotionProfile.apply(all_intents)

        # Build native MotionGraph actions and dependencies
        prev_action_id: Optional[str] = None
        for idx, intent in enumerate(all_intents):
            act = intent.to_animation_action(f"act_{idx}_{intent.target_id}")
            mg.add_action(act)
            if prev_action_id and intent.trigger == "with_previous":
                mg.add_dependency(prev_action_id, act.action_id, relationship="with_previous", lag_ms=intent.delay_ms)
            prev_action_id = act.action_id

        # Quality Validation
        quality_report = MotionQualityValidator.validate(
            all_intents,
            known_target_ids=set(sg.nodes.keys()),
            reduced_motion=use_reduced,
        )

        total_duration = sum(i.duration_ms for i in all_intents if i.trigger == "after_previous")
        decision_report = MotionDecisionReport(
            scene_id=sg.scene_id,
            decisions=decisions,
            narrative_flow_summary=f"Coordinated {len(profiles)} objects across {len(all_intents)} semantic motion intents.",
            cognitive_load_level="low" if quality_report.metrics.get("max_simultaneous", 1) <= 2 else "moderate",
            total_duration_ms=total_duration,
        )

        return mg, decision_report, quality_report

    def compile_plan_to_deck(self, plan: StoryboardPlan, output_path: str) -> str:
        """Compile an entire StoryboardPlan through SceneGraph and MotionGraph into native PPTX."""
        deck = MotionDeck(theme=self.theme)

        for scene in plan.scenes:
            trans = "none" if scene.scene_number == 1 else scene.transition_to_next
            slide = deck.add_slide(transition=trans)

            # 1. Build SceneGraph for this scene
            sg = self.design_scene_graph(scene)

            # 2. Compile SceneGraph into native PowerPoint shapes
            mapping = sg.compile_to_slide(slide)

            # 3. Build MotionGraph using universal choreography
            mg = self.design_motion_graph(scene, sg)

            # 4. Compile MotionGraph into native OpenXML animation timeline
            if mg.actions:
                timeline = mg.compile_to_timeline(scene.scene_number, mapping, sg)
                timeline.compile_to_openxml(slide)

        deck.save(output_path)
        return str(Path(output_path).resolve())

    def design_scene_graph(self, scene: StoryboardScene) -> SceneGraph:
        """Synthesize a structured SceneGraph based on scene composition style."""
        sg = SceneGraph(scene_id=f"scene_{scene.scene_number}")

        # Root Typography
        title_node = SceneNode(
            node_id="Title",
            name="!!Scene_Title",
            kind="textbox",
            text=scene.title,
            font_size_pt=32.0,
            text_color="0F172A",
            bold=True,
            transform=Transform(x=1.0, y=0.8, width=11.333, height=0.85),
            semantic_role="typography",
            overlap_policy="protected_readable",
        )
        sg.add_node(title_node)

        subtitle_node = SceneNode(
            node_id="Subtitle",
            name="!!Scene_Subtitle",
            kind="textbox",
            text=scene.subtitle,
            font_size_pt=16.0,
            text_color="475569",
            transform=Transform(x=1.0, y=1.8, width=11.333, height=0.6),
            semantic_role="typography",
            overlap_policy="protected_readable",
        )
        sg.add_node(subtitle_node)

        # Thematic traveling hero element (maintains !! across all scenes)
        hero_x = 2.0 + (scene.scene_number * 1.5) % 8.0
        hero_node = SceneNode(
            node_id="TravelingHero",
            name="!!Traveling_Hero_Orb",
            kind="ellipse",
            fill="4F46E5",
            transform=Transform(x=hero_x, y=3.2, width=1.5, height=1.5, z_index=10),
            semantic_role="subject",
            overlap_policy="allow_overlap",
        )
        sg.add_node(hero_node)

        # Scene-specific composition structures
        if scene.composition_style == "cinematic_hero":
            card = SceneNode(
                node_id="HeroCard",
                name="!!Hero_Card",
                kind="rect",
                fill="F1F5F9",
                stroke="CBD5E1",
                stroke_width=1.0,
                transform=Transform(x=1.5, y=2.8, width=10.333, height=3.8, z_index=1),
                semantic_role="card",
                overlap_policy="allow_overlap",
            )
            sg.add_node(card)
        elif scene.composition_style == "editorial_layout":
            col1 = SceneNode(
                node_id="EditorialCol1",
                name="!!Editorial_Col_1",
                kind="rect",
                fill="F8FAFC",
                stroke="E2E8F0",
                stroke_width=1.0,
                transform=Transform(x=1.0, y=2.8, width=5.2, height=3.8, z_index=1),
                semantic_role="card",
            )
            col2 = SceneNode(
                node_id="EditorialCol2",
                name="!!Editorial_Col_2",
                kind="rect",
                fill="F8FAFC",
                stroke="E2E8F0",
                stroke_width=1.0,
                transform=Transform(x=6.8, y=2.8, width=5.5, height=3.8, z_index=1),
                semantic_role="card",
            )
            sg.add_node(col1)
            sg.add_node(col2)
        elif scene.composition_style == "radial_infographic":
            for k in range(4):
                rad = 1.8
                cx = 6.666 + rad * 1.8 * (1 if k in (0, 1) else -1) * (0.6 if k in (0, 3) else 1.0)
                cy = 4.2 + rad * (1 if k in (1, 2) else -1) * 0.7
                node_k = SceneNode(
                    node_id=f"RadialNode_{k}",
                    name=f"!!Radial_Node_{k}",
                    kind="ellipse",
                    fill="0EA5E9",
                    transform=Transform(x=cx, y=cy, width=1.0, height=1.0, z_index=5),
                    semantic_role="hud",
                )
                sg.add_node(node_k)
        else:
            # Multi-card layout
            for idx in range(3):
                cx = 1.0 + idx * 3.9
                card_node = SceneNode(
                    node_id=f"Card_{idx}",
                    name=f"!!Card_{idx}",
                    kind="rect",
                    fill="F8FAFC",
                    stroke="E2E8F0",
                    stroke_width=1.0,
                    transform=Transform(x=cx, y=2.8, width=3.5, height=3.8, z_index=2),
                    semantic_role="card",
                )
                sg.add_node(card_node)

        return sg

    def design_motion_graph(self, scene: StoryboardScene, sg: SceneGraph) -> MotionGraph:
        """Create coordinated MotionGraph for intra-slide native animations."""
        mg, _, _ = self.choreograph_scene(sg, reduced_motion=self.reduced_motion)
        return mg
