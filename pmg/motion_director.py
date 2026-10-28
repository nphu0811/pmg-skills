"""AI Motion Director for PMG Skills (v1.7.0).

Transforms narrative storyboard plans into concrete Scene Graphs and Motion Graphs,
and compiles them into native editable PowerPoint presentations.
"""
from typing import Dict, Any, List, Optional
from pathlib import Path

from pmg.storyboard import StoryboardPlan, StoryboardScene
from pmg.scene_graph import SceneGraph, SceneNode, Transform
from pmg.motion_graph import MotionGraph, AnimationAction, Dependency
from pmg.animations import SlideAnimationTimeline
from scripts.motion_graph_kit import MotionDeck, Themes


class MotionDirector:
    """Intelligent Motion Director coordinating visual composition and motion design."""
    def __init__(self, theme=Themes.LIGHT_ACADEMIC_CYAN):
        self.theme = theme

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

            # 3. Build MotionGraph for this scene
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
                angle = k * (360.0 / 4.0)
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
        mg = MotionGraph(graph_id=f"motion_{scene.scene_number}")

        # Entrance for traveling hero
        a1 = AnimationAction(
            action_id="hero_entrance",
            target_node_id="TravelingHero",
            effect_type="zoom",
            duration_ms=800,
            delay_ms=100,
            trigger="after_previous",
        )
        mg.add_action(a1)

        # Simultaneous pulse/rotation
        a2 = AnimationAction(
            action_id="hero_spin",
            target_node_id="TravelingHero",
            effect_type="rotate",
            duration_ms=1200,
            rotation_deg=360.0,
            trigger="with_previous",
        )
        mg.add_action(a2)
        mg.add_dependency("hero_entrance", "hero_spin", relationship="with_previous")

        # Animate secondary cards/nodes sequentially
        secondary_nodes = [nid for nid in sg.nodes if nid not in ("Title", "Subtitle", "TravelingHero")]
        prev_action = "hero_entrance"
        for idx, nid in enumerate(secondary_nodes):
            act_id = f"anim_{nid}"
            act = AnimationAction(
                action_id=act_id,
                target_node_id=nid,
                effect_type="fly_in",
                direction="down" if idx % 2 == 0 else "up",
                duration_ms=650,
                delay_ms=idx * 80,
                trigger="after_previous" if idx == 0 else "with_previous",
            )
            mg.add_action(act)
            mg.add_dependency(prev_action, act_id, relationship="with_previous" if idx > 0 else "after_previous", lag_ms=idx * 80)

        return mg
