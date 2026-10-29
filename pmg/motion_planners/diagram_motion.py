"""
pmg.motion_planners.diagram_motion
==================================
Specialized Diagram & Process Motion Planner for PMG v2.5.0 (`pmg-diagram-motion`).

Designs causal, logically structured animation patterns (directional process flow,
node-by-node unfolding, connector vector drawing, hierarchical expansion, and data flow)
following strict dependency order: Node A -> Connector AB -> Node B -> Connector BC -> Node C.
"""

from __future__ import annotations
from typing import Dict, List, Optional, Any, Set, Tuple

from pmg.motion_planners.base import BaseMotionPlanner, MotionIntent, SharedMotionPrimitive


class DiagramMotionPlanner(BaseMotionPlanner):
    """
    Intelligent motion planner for flowcharts, system architectures, and relationship diagrams.
    Strict invariant: Connectors and nodes must animate in logical causal sequence.
    Never reveals complex diagrams all at once or in random order.
    """
    @property
    def planner_name(self) -> str:
        return "pmg-diagram-motion"

    @property
    def supported_object_types(self) -> Set[str]:
        return {"diagram", "flowchart", "process", "architecture", "system_graph"}

    def plan_motion(self, context: Dict[str, Any]) -> List[MotionIntent]:
        """
        Plans diagram animation strictly following logical topology.
        Context keys:
          - pattern: 'process_flow', 'hierarchical_expansion', 'cycle_flow', 'data_flow'
          - steps: List of dicts representing nodes: [{'id': 'node_1', 'label': '...'}]
          - connectors: List of dicts: [{'id': 'conn_1_2', 'source': 'node_1', 'target': 'node_2', 'direction': 'right'}]
          - root_id: For hierarchical expansion, the root system node
          - branch_nodes: For hierarchical expansion, child nodes grouped by branch
        """
        intents: List[MotionIntent] = []
        pattern = context.get("pattern", "process_flow").lower()
        steps = context.get("steps", [])
        connectors = context.get("connectors", [])

        # 1. Directional Process Flow: A -> Connector -> B -> Connector -> C
        if pattern == "process_flow":
            conn_by_source = {c.get("source"): c for c in connectors}
            
            for idx, step in enumerate(steps):
                step_id = step.get("id", f"node_{idx}")
                
                # Reveal Step Node
                intents.append(
                    MotionIntent(
                        target_id=step_id,
                        primitive=SharedMotionPrimitive.SEQUENCE,
                        semantic_purpose=f"Process Step {idx+1} unfolds ({step.get('label', step_id)})",
                        duration_ms=500,
                        delay_ms=80 if idx > 0 else 0,
                        trigger="after_previous",
                        direction="up" if idx % 2 == 0 else "down",
                        easing="ease_out",
                        narrative_importance="primary",
                    )
                )

                # Trace outgoing connector to the next node
                conn = conn_by_source.get(step_id)
                if conn:
                    conn_id = conn.get("id")
                    conn_dir = conn.get("direction", "right")
                    intents.append(
                        MotionIntent(
                            target_id=conn_id,
                            primitive=SharedMotionPrimitive.DRAW,
                            semantic_purpose=f"Causal connector traces flow to next step ({conn_dir})",
                            duration_ms=400,
                            delay_ms=40,
                            trigger="after_previous",
                            direction=conn_dir,
                            easing="linear",
                            narrative_importance="supporting",
                        )
                    )

        # 2. Hierarchical Expansion (Root -> Distribution Connectors -> Child Modules)
        elif pattern == "hierarchical_expansion":
            root_id = context.get("root_id", "root_node")
            
            # Root Node Appears
            intents.append(
                MotionIntent(
                    target_id=root_id,
                    primitive=SharedMotionPrimitive.FOCUS,
                    semantic_purpose="Central core architecture root establishes system anchor",
                    duration_ms=650,
                    delay_ms=0,
                    trigger="after_previous",
                    scale_factor=1.08,
                    narrative_importance="hero",
                )
            )

            # Connectors Branch Out Simultaneously
            for conn in connectors:
                conn_id = conn.get("id")
                conn_dir = conn.get("direction", "down")
                intents.append(
                    MotionIntent(
                        target_id=conn_id,
                        primitive=SharedMotionPrimitive.DRAW,
                        semantic_purpose=f"Branch vector traces connection to module",
                        duration_ms=450,
                        delay_ms=60,
                        trigger="with_previous",
                        direction=conn_dir,
                        easing="ease_out",
                        narrative_importance="supporting",
                    )
                )

            # Child Leaves Cascade In
            children = context.get("children", [s.get("id") for s in steps if s.get("id") != root_id])
            for cidx, cid in enumerate(children):
                intents.append(
                    MotionIntent(
                        target_id=cid,
                        primitive=SharedMotionPrimitive.STAGGER,
                        semantic_purpose=f"Subsystem module {cidx+1} activates",
                        duration_ms=480,
                        delay_ms=cidx * 90,
                        trigger="after_previous" if cidx == 0 else "with_previous",
                        direction="up",
                        narrative_importance="primary",
                    )
                )

        # 3. Data-Flow Visualization (Traveling Signal along Path)
        elif pattern == "data_flow":
            # Nodes already established; signal orb travels along pipeline
            signal_id = context.get("signal_id", "data_packet")
            intents.append(
                MotionIntent(
                    target_id=signal_id,
                    primitive=SharedMotionPrimitive.MOVE,
                    semantic_purpose="Data payload travels through pipeline nodes demonstrating real-time throughput",
                    duration_ms=1800,
                    delay_ms=100,
                    trigger="after_previous",
                    easing="ease_in_out",
                    parameters={"path_type": "pipeline_linear"},
                    narrative_importance="hero",
                )
            )

        # 4. Closed Loop Cycle Flow
        elif pattern == "cycle_flow":
            for idx, step in enumerate(steps):
                step_id = step.get("id", f"cycle_node_{idx}")
                intents.append(
                    MotionIntent(
                        target_id=step_id,
                        primitive=SharedMotionPrimitive.SEQUENCE,
                        semantic_purpose=f"Cycle Phase {idx+1} reveals in continuous feedback loop",
                        duration_ms=500,
                        delay_ms=100,
                        trigger="after_previous",
                        narrative_importance="primary",
                    )
                )
                if idx < len(connectors):
                    conn_id = connectors[idx].get("id")
                    intents.append(
                        MotionIntent(
                            target_id=conn_id,
                            primitive=SharedMotionPrimitive.DRAW,
                            semantic_purpose=f"Curved cycle vector connects to next phase",
                            duration_ms=380,
                            delay_ms=30,
                            trigger="after_previous",
                            direction=connectors[idx].get("direction", "right"),
                            narrative_importance="supporting",
                        )
                    )

        return intents

    def explain_selection(self, intent: MotionIntent, context: Dict[str, Any]) -> str:
        return (
            f"DiagramMotionPlanner selected '{intent.primitive.value}' for '{intent.target_id}' "
            f"to enforce causality and topological dependency across process nodes."
        )
