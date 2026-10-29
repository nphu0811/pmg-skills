"""
pmg.motion_planners.chart_motion
================================
Specialized Chart & Numerical Motion Planner for PMG v2.5.0 (`pmg-chart-motion`).

Designs purposeful data animation patterns (bar growth, line draw, doughnut slice reveal,
value counter synchronization, and peak metric emphasis) reflecting quantitative relationships.
"""

from __future__ import annotations
from typing import Dict, List, Optional, Any, Set

from pmg.motion_planners.base import BaseMotionPlanner, MotionIntent, SharedMotionPrimitive
from pmg.semantic_classifier import ObjectType, SemanticRole, NarrativeImportance, ObjectSemanticProfile


class ChartMotionPlanner(BaseMotionPlanner):
    """
    Intelligent motion planner for quantitative data charts and numerical indicators.
    Never uses generic Fade In; tailors motion to the data geometry and quantitative meaning.
    """
    @property
    def planner_name(self) -> str:
        return "pmg-chart-motion"

    @property
    def supported_object_types(self) -> Set[str]:
        return {"chart", "metric", "bar_chart", "line_chart", "doughnut_chart"}

    def plan_motion(self, context: Dict[str, Any]) -> List[MotionIntent]:
        """
        Plans coordinated data animation based on chart type and components.
        Context expected keys:
          - chart_type: 'bar', 'horizontal_bar', 'line', 'doughnut', 'metric_counter'
          - items: List of dicts with 'id', 'role', 'value', 'is_peak', 'label_id'
          - axis_id: Optional ID for the coordinate axis
        """
        intents: List[MotionIntent] = []
        chart_type = context.get("chart_type", "bar").lower()
        items = context.get("items", [])
        axis_id = context.get("axis_id")

        # 1. Grounding Axis Setup
        if axis_id:
            intents.append(
                MotionIntent(
                    target_id=axis_id,
                    primitive=SharedMotionPrimitive.DRAW,
                    semantic_purpose="Ground coordinate axes before data disclosure",
                    duration_ms=450,
                    delay_ms=0,
                    trigger="after_previous",
                    direction="right" if "bar" in chart_type else "up",
                    narrative_importance="supporting",
                )
            )

        # 2. Bar Chart Motion: Directional Growth from Base
        if chart_type in ("bar", "column", "vertical_bar"):
            for idx, item in enumerate(items):
                item_id = item.get("id", f"bar_{idx}")
                is_peak = item.get("is_peak", False)
                label_id = item.get("label_id")

                # Bar grows upward from baseline
                intents.append(
                    MotionIntent(
                        target_id=item_id,
                        primitive=SharedMotionPrimitive.GROW,
                        semantic_purpose=f"Bar grows vertically to represent value ({item.get('value', '')})",
                        duration_ms=750,
                        delay_ms=0 if idx == 0 else 120,
                        trigger="after_previous" if (idx == 0 and not axis_id) else ("after_previous" if idx == 0 else "with_previous"),
                        direction="up",
                        easing="ease_out",
                        narrative_importance="primary",
                    )
                )

                # Synchronized numeric label counter reveal
                if label_id:
                    intents.append(
                        MotionIntent(
                            target_id=label_id,
                            primitive=SharedMotionPrimitive.COUNT,
                            semantic_purpose=f"Synchronized value callout tracking bar height",
                            duration_ms=300,
                            delay_ms=200,
                            trigger="with_previous",
                            narrative_importance="secondary",
                        )
                    )

                # Peak metric emphasis
                if is_peak:
                    intents.append(
                        MotionIntent(
                            target_id=item_id,
                            primitive=SharedMotionPrimitive.EMPHASIZE,
                            semantic_purpose="Highlight peak metric as key takeaway",
                            duration_ms=400,
                            delay_ms=50,
                            trigger="after_previous",
                            scale_factor=1.06,
                            narrative_importance="hero",
                        )
                    )

        # 3. Horizontal Bar Chart Motion: Growth from Left to Right
        elif chart_type in ("horizontal_bar", "bar_h"):
            for idx, item in enumerate(items):
                item_id = item.get("id", f"hbar_{idx}")
                label_id = item.get("label_id")
                intents.append(
                    MotionIntent(
                        target_id=item_id,
                        primitive=SharedMotionPrimitive.GROW,
                        semantic_purpose="Horizontal bar expands rightward to quantitative target",
                        duration_ms=700,
                        delay_ms=0 if idx == 0 else 100,
                        trigger="after_previous" if idx == 0 else "with_previous",
                        direction="right",
                        easing="ease_out",
                    )
                )
                if label_id:
                    intents.append(
                        MotionIntent(
                            target_id=label_id,
                            primitive=SharedMotionPrimitive.TRACK,
                            semantic_purpose="Numeric label tracks horizontal bar tip",
                            duration_ms=300,
                            delay_ms=180,
                            trigger="with_previous",
                        )
                    )

        # 4. Line Chart: Progressive Chronological Trajectory
        elif chart_type == "line":
            for idx, item in enumerate(items):
                line_id = item.get("id", f"line_{idx}")
                intents.append(
                    MotionIntent(
                        target_id=line_id,
                        primitive=SharedMotionPrimitive.DRAW,
                        semantic_purpose="Progressive line trajectory drawn along chronological timeline",
                        duration_ms=1100,
                        delay_ms=0 if idx == 0 else 150,
                        trigger="after_previous" if idx == 0 else "with_previous",
                        direction="right",
                        easing="ease_in_out",
                        narrative_importance="primary",
                    )
                )
                # Key markers along the line
                marker_ids = item.get("marker_ids", [])
                for midx, mid in enumerate(marker_ids):
                    intents.append(
                        MotionIntent(
                            target_id=mid,
                            primitive=SharedMotionPrimitive.FOCUS,
                            semantic_purpose=f"Key inflection data point drop",
                            duration_ms=350,
                            delay_ms=150 + midx * 120,
                            trigger="with_previous",
                            narrative_importance="secondary",
                        )
                    )

        # 5. Doughnut / Pie Chart: Radial Slice Progression
        elif chart_type in ("doughnut", "pie"):
            center_kpi_id = context.get("center_kpi_id")
            for idx, item in enumerate(items):
                slice_id = item.get("id", f"slice_{idx}")
                intents.append(
                    MotionIntent(
                        target_id=slice_id,
                        primitive=SharedMotionPrimitive.GROW,
                        semantic_purpose=f"Doughnut sector {idx+1} unfolds proportionally",
                        duration_ms=600,
                        delay_ms=idx * 90,
                        trigger="after_previous" if idx == 0 else "with_previous",
                        easing="ease_out",
                        narrative_importance="primary",
                    )
                )
            if center_kpi_id:
                intents.append(
                    MotionIntent(
                        target_id=center_kpi_id,
                        primitive=SharedMotionPrimitive.EMPHASIZE,
                        semantic_purpose="Central aggregate KPI metric punch",
                        duration_ms=450,
                        delay_ms=100,
                        trigger="after_previous",
                        scale_factor=1.12,
                        narrative_importance="hero",
                    )
                )

        # 6. Standalone Metric Counter Callout
        elif chart_type == "metric_counter":
            for idx, item in enumerate(items):
                metric_id = item.get("id", f"metric_{idx}")
                label_id = item.get("label_id")
                intents.append(
                    MotionIntent(
                        target_id=metric_id,
                        primitive=SharedMotionPrimitive.COUNT,
                        semantic_purpose="Numerical stat count-up to target value",
                        duration_ms=850,
                        delay_ms=idx * 150,
                        trigger="after_previous" if idx == 0 else "with_previous",
                        narrative_importance="hero",
                    )
                )
                if label_id:
                    intents.append(
                        MotionIntent(
                            target_id=label_id,
                            primitive=SharedMotionPrimitive.REVEAL,
                            semantic_purpose="Supporting metric descriptor settles into place",
                            duration_ms=400,
                            delay_ms=150,
                            trigger="with_previous",
                            narrative_importance="secondary",
                        )
                    )

        return intents

    def explain_selection(self, intent: MotionIntent, context: Dict[str, Any]) -> str:
        return (
            f"ChartMotionPlanner selected '{intent.primitive.value}' ({intent.direction or 'neutral'}) "
            f"for '{intent.target_id}' to communicate quantitative progression faithfully without visual distraction."
        )
