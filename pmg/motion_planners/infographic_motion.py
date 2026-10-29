"""
pmg.motion_planners.infographic_motion
======================================
Specialized Infographic & Timeline Motion Planner for PMG v2.5.0 (`pmg-infographic-motion`).

Designs structured informational animation patterns (chronological timeline progression,
milestone pin emergence, horizontal & circular progress bar growth, feature comparisons,
and layered information disclosure).
"""

from __future__ import annotations
from typing import Dict, List, Optional, Any, Set

from pmg.motion_planners.base import BaseMotionPlanner, MotionIntent, SharedMotionPrimitive


class InfographicMotionPlanner(BaseMotionPlanner):
    """
    Intelligent motion planner for timelines, progress gauges, and structured visual infographics.
    Invariants: Animation directly mirrors informational hierarchy and chronological flow.
    """
    @property
    def planner_name(self) -> str:
        return "pmg-infographic-motion"

    @property
    def supported_object_types(self) -> Set[str]:
        return {"infographic", "timeline", "progress", "milestone", "gauge", "comparison"}

    def plan_motion(self, context: Dict[str, Any]) -> List[MotionIntent]:
        """
        Plans infographic motion.
        Context keys:
          - pattern: 'timeline_progression', 'progress_bar', 'circular_gauge', 'feature_comparison'
          - spine_id: For timelines, the primary chronological axis
          - milestones: List of dicts [{'node_id': 'm1_node', 'card_id': 'm1_card', 'year': '2024'}]
          - track_id: For progress bars, the background track
          - fill_id: For progress bars, the active growing bar
          - percent_id: For progress bars, the numerical percentage label
          - items: List of comparison cards or icons
        """
        intents: List[MotionIntent] = []
        pattern = context.get("pattern", "timeline_progression").lower()

        # 1. Chronological Timeline Progression
        if pattern == "timeline_progression":
            spine_id = context.get("spine_id", "timeline_spine")
            milestones = context.get("milestones", [])
            orientation = context.get("orientation", "horizontal")

            # Timeline Spine Traces Across the Slide
            intents.append(
                MotionIntent(
                    target_id=spine_id,
                    primitive=SharedMotionPrimitive.DRAW,
                    semantic_purpose="Chronological timeline spine drawn to anchor narrative events",
                    duration_ms=800,
                    delay_ms=50,
                    trigger="after_previous",
                    direction="right" if orientation == "horizontal" else "down",
                    easing="ease_in_out",
                    narrative_importance="supporting",
                )
            )

            # Milestones and Cards Unfold Chronologically
            for idx, m in enumerate(milestones):
                node_id = m.get("node_id", f"milestone_pin_{idx}")
                card_id = m.get("card_id")

                # Milestone Pin Emerges on the Spine
                intents.append(
                    MotionIntent(
                        target_id=node_id,
                        primitive=SharedMotionPrimitive.FOCUS,
                        semantic_purpose=f"Milestone {idx+1} ({m.get('year', '')}) anchors on timeline",
                        duration_ms=380,
                        delay_ms=80 if idx > 0 else 40,
                        trigger="after_previous",
                        scale_factor=1.15,
                        easing="ease_out",
                        narrative_importance="primary",
                    )
                )

                # Accompanying Event Card Expands
                if card_id:
                    intents.append(
                        MotionIntent(
                            target_id=card_id,
                            primitive=SharedMotionPrimitive.REVEAL,
                            semantic_purpose=f"Milestone detail card expands with event briefing",
                            duration_ms=450,
                            delay_ms=60,
                            trigger="with_previous",
                            direction="up" if idx % 2 == 0 else "down",
                            narrative_importance="secondary",
                        )
                    )

        # 2. Linear Progress Bar Growth with Value Tracking
        elif pattern == "progress_bar":
            track_id = context.get("track_id", "progress_track")
            fill_id = context.get("fill_id", "progress_fill")
            percent_id = context.get("percent_id")

            # Background Track Appears
            intents.append(
                MotionIntent(
                    target_id=track_id,
                    primitive=SharedMotionPrimitive.REVEAL,
                    semantic_purpose="Progress bar track framework establishes scale",
                    duration_ms=400,
                    delay_ms=0,
                    trigger="after_previous",
                    narrative_importance="supporting",
                )
            )

            # Active Bar Expands from 0% to Target
            intents.append(
                MotionIntent(
                    target_id=fill_id,
                    primitive=SharedMotionPrimitive.GROW,
                    semantic_purpose="Active metric fill grows horizontally toward completion target",
                    duration_ms=1100,
                    delay_ms=80,
                    trigger="after_previous",
                    direction="right",
                    easing="ease_out",
                    narrative_importance="primary",
                )
            )

            # Percentage Text Follows the Bar Growth
            if percent_id:
                intents.append(
                    MotionIntent(
                        target_id=percent_id,
                        primitive=SharedMotionPrimitive.TRACK,
                        semantic_purpose="Percentage indicator synchronizes with fill advancement",
                        duration_ms=400,
                        delay_ms=250,
                        trigger="with_previous",
                        narrative_importance="hero",
                    )
                )

        # 3. Circular Gauge / Radial Progression
        elif pattern == "circular_gauge":
            dial_id = context.get("dial_id", "dial_ring")
            value_id = context.get("value_id", "dial_value")
            intents.append(
                MotionIntent(
                    target_id=dial_id,
                    primitive=SharedMotionPrimitive.GROW,
                    semantic_purpose="Radial progress dial unfolds proportionally",
                    duration_ms=1000,
                    delay_ms=50,
                    trigger="after_previous",
                    easing="ease_out",
                    narrative_importance="primary",
                )
            )
            intents.append(
                MotionIntent(
                    target_id=value_id,
                    primitive=SharedMotionPrimitive.EMPHASIZE,
                    semantic_purpose="Central percentage value locks into place",
                    duration_ms=450,
                    delay_ms=150,
                    trigger="after_previous",
                    scale_factor=1.10,
                    narrative_importance="hero",
                )
            )

        # 4. Feature / Option Comparison Matrix
        elif pattern == "feature_comparison":
            cards = context.get("items", [])
            winner_badge_id = context.get("winner_badge_id")

            for cidx, card in enumerate(cards):
                card_id = card.get("id", f"compare_card_{cidx}")
                intents.append(
                    MotionIntent(
                        target_id=card_id,
                        primitive=SharedMotionPrimitive.STAGGER,
                        semantic_purpose=f"Comparison alternative {cidx+1} revealed side-by-side",
                        duration_ms=550,
                        delay_ms=cidx * 140,
                        trigger="after_previous" if cidx == 0 else "with_previous",
                        direction="up",
                        narrative_importance="primary",
                    )
                )

            # Highlight the recommended/winning option
            if winner_badge_id:
                intents.append(
                    MotionIntent(
                        target_id=winner_badge_id,
                        primitive=SharedMotionPrimitive.EMPHASIZE,
                        semantic_purpose="Highlight badge flags winning strategic option",
                        duration_ms=500,
                        delay_ms=120,
                        trigger="after_previous",
                        scale_factor=1.15,
                        narrative_importance="hero",
                    )
                )

        return intents

    def explain_selection(self, intent: MotionIntent, context: Dict[str, Any]) -> str:
        return (
            f"InfographicMotionPlanner selected '{intent.primitive.value}' for '{intent.target_id}' "
            f"to articulate structured progression without overwhelming audience attention."
        )
