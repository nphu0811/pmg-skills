"""
pmg.motion_planners.image_motion
================================
Specialized Image & Visual Composition Motion Planner for PMG v2.5.0 (`pmg-image-motion`).

Designs purposeful photographic and visual asset animation patterns (directional reveal,
mask unmasking, layered multi-plane parallax, controlled macro dolly zoom, pan,
before/after comparison, and foreground/background depth separation).
"""

from __future__ import annotations
from typing import Dict, List, Optional, Any, Set

from pmg.motion_planners.base import BaseMotionPlanner, MotionIntent, SharedMotionPrimitive


class ImageMotionPlanner(BaseMotionPlanner):
    """
    Intelligent motion planner for photographs, illustrations, and decomposed visual layers.
    Preserves strict aspect ratio locks, prevents image distortion, and rejects indiscriminate Ken Burns.
    """
    @property
    def planner_name(self) -> str:
        return "pmg-image-motion"

    @property
    def supported_object_types(self) -> Set[str]:
        return {"image", "photo", "picture", "illustration", "cutout"}

    def plan_motion(self, context: Dict[str, Any]) -> List[MotionIntent]:
        """
        Plans image animation based on compositional intent.
        Context keys:
          - pattern: 'directional_reveal', 'layered_parallax', 'controlled_zoom', 'pan',
                     'before_after', 'foreground_background_separation', 'mask_reveal'
          - target_id: Main image or group ID
          - background_id: Optional background layer ID
          - foreground_id: Optional foreground layer ID
          - before_id: Optional before image ID
          - after_id: Optional after image ID
          - divider_id: Optional divider bar ID for before/after comparison
        """
        intents: List[MotionIntent] = []
        pattern = context.get("pattern", "directional_reveal").lower()
        target_id = context.get("target_id", "image_main")

        # 1. Directional Reveal along Leading Lines
        if pattern == "directional_reveal":
            direction = context.get("direction", "right")
            intents.append(
                MotionIntent(
                    target_id=target_id,
                    primitive=SharedMotionPrimitive.REVEAL,
                    semantic_purpose=f"Directional reveal along photographic leading line ({direction})",
                    duration_ms=750,
                    delay_ms=80,
                    trigger="after_previous",
                    direction=direction,
                    easing="ease_out",
                    narrative_importance="primary",
                )
            )

        # 2. Layered Multi-Plane Parallax (BG 0.7x, Mid 1.0x, FG 1.4x)
        elif pattern == "layered_parallax":
            bg_id = context.get("background_id")
            mid_id = target_id
            fg_id = context.get("foreground_id")

            # Background: Slow subtle drift (0.7x speed, subtle scale or path)
            if bg_id:
                intents.append(
                    MotionIntent(
                        target_id=bg_id,
                        primitive=SharedMotionPrimitive.MOVE,
                        semantic_purpose="Background plane moves at 0.7x parallax speed establishing depth",
                        duration_ms=1800,
                        delay_ms=0,
                        trigger="after_previous",
                        direction="left",
                        parameters={"speed_factor": 0.7},
                        narrative_importance="background",
                    )
                )

            # Midground subject: Primary tracking speed (1.0x)
            intents.append(
                MotionIntent(
                    target_id=mid_id,
                    primitive=SharedMotionPrimitive.MOVE,
                    semantic_purpose="Subject plane tracks at 1.0x speed as primary focus",
                    duration_ms=1600,
                    delay_ms=50,
                    trigger="with_previous" if bg_id else "after_previous",
                    direction="left",
                    parameters={"speed_factor": 1.0},
                    narrative_importance="primary",
                )
            )

            # Foreground scrim/framing: Accelerated push (1.4x speed)
            if fg_id:
                intents.append(
                    MotionIntent(
                        target_id=fg_id,
                        primitive=SharedMotionPrimitive.MOVE,
                        semantic_purpose="Foreground framing elements clear at 1.4x parallax speed",
                        duration_ms=1300,
                        delay_ms=100,
                        trigger="with_previous",
                        direction="left",
                        parameters={"speed_factor": 1.4},
                        narrative_importance="supporting",
                    )
                )

        # 3. Controlled Macro Dolly Zoom (Push-In)
        elif pattern in ("controlled_zoom", "dolly_zoom"):
            scale_target = min(context.get("scale_factor", 1.15), 1.25)  # Guardrail against pixelation
            intents.append(
                MotionIntent(
                    target_id=target_id,
                    primitive=SharedMotionPrimitive.SCALE,
                    semantic_purpose=f"Controlled macro push-in zoom to {int(scale_target*100)}% to reveal detail",
                    duration_ms=1500,
                    delay_ms=100,
                    trigger="after_previous",
                    scale_factor=scale_target,
                    easing="ease_in_out",
                    narrative_importance="primary",
                )
            )

        # 4. Cinematic Tracking Pan
        elif pattern == "pan":
            intents.append(
                MotionIntent(
                    target_id=target_id,
                    primitive=SharedMotionPrimitive.MOVE,
                    semantic_purpose="Cinematic horizontal pan across panoramic image landscape",
                    duration_ms=2200,
                    delay_ms=50,
                    trigger="after_previous",
                    direction=context.get("direction", "left"),
                    easing="linear",
                    narrative_importance="primary",
                )
            )

        # 5. Before / After Comparison
        elif pattern in ("before_after", "comparison"):
            before_id = context.get("before_id", f"{target_id}_before")
            after_id = context.get("after_id", f"{target_id}_after")
            divider_id = context.get("divider_id")

            # Base 'before' state established
            intents.append(
                MotionIntent(
                    target_id=before_id,
                    primitive=SharedMotionPrimitive.REVEAL,
                    semantic_purpose="Baseline 'Before' state displayed for reference",
                    duration_ms=500,
                    delay_ms=0,
                    trigger="after_previous",
                    narrative_importance="secondary",
                )
            )

            # 'After' state wiped in synchronously with divider
            intents.append(
                MotionIntent(
                    target_id=after_id,
                    primitive=SharedMotionPrimitive.DRAW,
                    semantic_purpose="'After' transformation state wiped from left to right",
                    duration_ms=1000,
                    delay_ms=250,
                    trigger="after_previous",
                    direction="right",
                    easing="ease_in_out",
                    narrative_importance="hero",
                )
            )
            if divider_id:
                intents.append(
                    MotionIntent(
                        target_id=divider_id,
                        primitive=SharedMotionPrimitive.MOVE,
                        semantic_purpose="Visual split divider tracks wipe front",
                        duration_ms=1000,
                        delay_ms=250,
                        trigger="with_previous",
                        direction="right",
                        narrative_importance="supporting",
                    )
                )

        # 6. Foreground / Background Depth Separation
        elif pattern == "foreground_background_separation":
            bg_id = context.get("background_id", f"{target_id}_bg")
            cutout_id = context.get("foreground_id", f"{target_id}_cutout")

            # Atmospheric background settles softly
            intents.append(
                MotionIntent(
                    target_id=bg_id,
                    primitive=SharedMotionPrimitive.REVEAL,
                    semantic_purpose="Atmospheric context background settles into place",
                    duration_ms=600,
                    delay_ms=0,
                    trigger="after_previous",
                    narrative_importance="background",
                )
            )
            # Subject cutout lifts off canvas
            intents.append(
                MotionIntent(
                    target_id=cutout_id,
                    primitive=SharedMotionPrimitive.FOCUS,
                    semantic_purpose="Native vector cutout subject steps forward establishing 3D separation",
                    duration_ms=750,
                    delay_ms=120,
                    trigger="after_previous",
                    scale_factor=1.05,
                    easing="ease_out",
                    narrative_importance="hero",
                )
            )

        return intents

    def explain_selection(self, intent: MotionIntent, context: Dict[str, Any]) -> str:
        return (
            f"ImageMotionPlanner selected '{intent.primitive.value}' for '{intent.target_id}' "
            f"to preserve photographic framing and aspect ratio while enhancing spatial depth."
        )
