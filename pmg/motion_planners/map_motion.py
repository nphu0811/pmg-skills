"""
pmg.motion_planners.map_motion
==============================
Specialized Geospatial & Map Motion Planner for PMG v2.5.0 (`pmg-map-motion`).

Designs geographic animation patterns (region highlighting, progressive route drawing,
location marker pin drops, hub radar pulses, and multi-location itinerary storytelling)
preserving strict spatial truth and geographic proportions.
"""

from __future__ import annotations
from typing import Dict, List, Optional, Any, Set

from pmg.motion_planners.base import BaseMotionPlanner, MotionIntent, SharedMotionPrimitive


class MapMotionPlanner(BaseMotionPlanner):
    """
    Intelligent motion planner for geographic maps, territorial regions, and route networks.
    Invariants: Preserves verified geographic relationships and avoids false topological distortions.
    """
    @property
    def planner_name(self) -> str:
        return "pmg-map-motion"

    @property
    def supported_object_types(self) -> Set[str]:
        return {"map", "geo_region", "geo_marker", "geo_route", "territory"}

    def plan_motion(self, context: Dict[str, Any]) -> List[MotionIntent]:
        """
        Plans geospatial animation.
        Context keys:
          - pattern: 'route_progression', 'region_highlight', 'multi_hub_network', 'geo_zoom'
          - base_map_id: Base map container ID
          - locations: List of location dicts [{'id': 'pin_hcmc', 'label': 'HCMC', 'pulse_id': '...'}]
          - routes: List of route dicts [{'id': 'route_1', 'origin': 'pin_hanoi', 'dest': 'pin_hcmc', 'direction': 'down'}]
          - regions: List of territory region IDs
          - overlay_id: Optional stat panel overlay ID
        """
        intents: List[MotionIntent] = []
        pattern = context.get("pattern", "route_progression").lower()
        base_map_id = context.get("base_map_id")

        # 1. Base Map Settles into View
        if base_map_id:
            intents.append(
                MotionIntent(
                    target_id=base_map_id,
                    primitive=SharedMotionPrimitive.REVEAL,
                    semantic_purpose="Geographic landmass establishes spatial orientation",
                    duration_ms=600,
                    delay_ms=0,
                    trigger="after_previous",
                    narrative_importance="supporting",
                )
            )

        # 2. Progressive Route Drawing & Pin Drop Itinerary
        if pattern == "route_progression":
            locations = context.get("locations", [])
            routes = context.get("routes", [])

            for idx, loc in enumerate(locations):
                loc_id = loc.get("id", f"pin_{idx}")
                pulse_id = loc.get("pulse_id")

                # Marker Pin Drops into Coordinate
                intents.append(
                    MotionIntent(
                        target_id=loc_id,
                        primitive=SharedMotionPrimitive.MOVE,
                        semantic_purpose=f"Location marker drops at verified coordinates ({loc.get('label', loc_id)})",
                        duration_ms=450,
                        delay_ms=60 if idx > 0 else 0,
                        trigger="after_previous",
                        direction="down",
                        easing="ease_out",
                        narrative_importance="primary",
                    )
                )

                # Radar pulse ripple around key hub
                if pulse_id:
                    intents.append(
                        MotionIntent(
                            target_id=pulse_id,
                            primitive=SharedMotionPrimitive.HIGHLIGHT,
                            semantic_purpose=f"Radar pulse ring ripples from hub location",
                            duration_ms=600,
                            delay_ms=40,
                            trigger="with_previous",
                            narrative_importance="secondary",
                        )
                    )

                # Route draws to the subsequent location
                if idx < len(routes):
                    route = routes[idx]
                    intents.append(
                        MotionIntent(
                            target_id=route.get("id"),
                            primitive=SharedMotionPrimitive.DRAW,
                            semantic_purpose=f"Logistics / travel route progressively traces vector path",
                            duration_ms=750,
                            delay_ms=40,
                            trigger="after_previous",
                            direction=route.get("direction", "down"),
                            easing="linear",
                            narrative_importance="primary",
                        )
                    )

        # 3. Territory Region Highlighting
        elif pattern == "region_highlight":
            regions = context.get("regions", [])
            overlay_id = context.get("overlay_id")

            for ridx, rid in enumerate(regions):
                intents.append(
                    MotionIntent(
                        target_id=rid,
                        primitive=SharedMotionPrimitive.HIGHLIGHT,
                        semantic_purpose=f"Geographic territory {ridx+1} illuminated with accent fill",
                        duration_ms=650,
                        delay_ms=ridx * 120,
                        trigger="after_previous" if ridx == 0 else "with_previous",
                        narrative_importance="primary",
                    )
                )

            # Accompanying regional demographic / business metrics overlay
            if overlay_id:
                intents.append(
                    MotionIntent(
                        target_id=overlay_id,
                        primitive=SharedMotionPrimitive.REVEAL,
                        semantic_purpose="Regional telemetry metrics panel docks adjacent to highlighted territory",
                        duration_ms=500,
                        delay_ms=100,
                        trigger="after_previous",
                        direction="right",
                        narrative_importance="hero",
                    )
                )

        # 4. Geographic Dolly Zoom Focus
        elif pattern == "geo_zoom":
            cluster_id = context.get("cluster_id", "metro_cluster")
            intents.append(
                MotionIntent(
                    target_id=cluster_id,
                    primitive=SharedMotionPrimitive.FOCUS,
                    semantic_purpose="Macro push-in zoom into high-density metropolitan corridor",
                    duration_ms=1400,
                    delay_ms=100,
                    trigger="after_previous",
                    scale_factor=1.20,
                    easing="ease_in_out",
                    narrative_importance="hero",
                )
            )

        return intents

    def explain_selection(self, intent: MotionIntent, context: Dict[str, Any]) -> str:
        return (
            f"MapMotionPlanner selected '{intent.primitive.value}' for '{intent.target_id}' "
            f"to trace verified spatial trajectories without distorting geographic truth."
        )
