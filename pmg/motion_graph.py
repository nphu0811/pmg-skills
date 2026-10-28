"""Canonical Motion Graph data model for PMG Skills.

Expresses animation events, sequential and parallel dependencies, parent-child motion inheritance,
trigger rules, and compiles directly into declarative SlideAnimationTimeline.
"""
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any, Set
from collections import defaultdict, deque

from pmg.animations import SlideAnimationTimeline, normalize_effect


@dataclass
class AnimationAction:
    """Individual animation action targeting a SceneNode."""
    action_id: str
    target_node_id: str
    effect_type: str  # 'appear', 'fade', 'fly_in', 'zoom', 'rotate', 'scale', 'motion_path', 'wipe'
    duration_ms: int = 1000
    delay_ms: int = 0
    trigger: str = "after_previous"  # 'after_previous', 'with_previous', 'on_click'
    direction: Optional[str] = None
    rotation_deg: float = 360.0
    scale_factor: float = 1.0
    motion_path: Optional[str] = None
    auto_reverse: bool = False
    repeat_count: int = 1
    parameters: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        self.effect_type = normalize_effect(self.effect_type)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class Dependency:
    """Directed temporal dependency between two animation actions."""
    source_action_id: str
    target_action_id: str
    relationship: str = "after_previous"  # 'after_previous' or 'with_previous'
    lag_ms: int = 0


class MotionGraph:
    """Graph of coordinated animations and causal dependencies."""
    def __init__(self, graph_id: str):
        self.graph_id = graph_id
        self.actions: Dict[str, AnimationAction] = {}
        self.dependencies: List[Dependency] = []

    def add_action(self, action: AnimationAction) -> AnimationAction:
        self.actions[action.action_id] = action
        return action

    def add_dependency(self, source_action_id: str, target_action_id: str,
                       relationship: str = "after_previous", lag_ms: int = 0) -> Dependency:
        if source_action_id not in self.actions:
            raise KeyError(f"Source action '{source_action_id}' not found in MotionGraph.")
        if target_action_id not in self.actions:
            raise KeyError(f"Target action '{target_action_id}' not found in MotionGraph.")
        dep = Dependency(source_action_id, target_action_id, relationship, lag_ms)
        self.dependencies.append(dep)
        # Update target action's trigger to match relationship
        self.actions[target_action_id].trigger = relationship
        if lag_ms > 0:
            self.actions[target_action_id].delay_ms = lag_ms
        return dep

    def validate_graph(self, valid_node_ids: Optional[Set[str]] = None) -> List[str]:
        """Detect cycles, missing targets, or dangling action dependencies."""
        errors = []
        adj = defaultdict(list)
        in_degree = defaultdict(int)

        for act_id in self.actions:
            in_degree[act_id] = 0

        for dep in self.dependencies:
            adj[dep.source_action_id].append(dep.target_action_id)
            in_degree[dep.target_action_id] += 1

        # Check for target node validity if provided
        if valid_node_ids is not None:
            for act_id, act in self.actions.items():
                if act.target_node_id not in valid_node_ids:
                    errors.append(f"Action '{act_id}' targets non-existent SceneNode '{act.target_node_id}'")

        # Topological sort (Kahn's algorithm) to detect cycles
        queue = deque([act_id for act_id, deg in in_degree.items() if deg == 0])
        visited_count = 0

        while queue:
            curr = queue.popleft()
            visited_count += 1
            for neighbor in adj[curr]:
                in_degree[neighbor] -= 1
                if in_degree[neighbor] == 0:
                    queue.append(neighbor)

        if visited_count < len(self.actions):
            errors.append("Cycle detected in MotionGraph dependencies")

        return errors

    def compile_to_timeline(self, slide_number: int, shape_mapping: Dict[str, Any],
                            scene_graph=None) -> SlideAnimationTimeline:
        """Compile the MotionGraph into an executable SlideAnimationTimeline."""
        valid_nodes = set(shape_mapping.keys())
        if scene_graph is not None:
            valid_nodes.update(scene_graph.nodes.keys())

        errors = self.validate_graph(valid_nodes)
        if errors:
            raise ValueError("Invalid MotionGraph: " + "; ".join(errors))

        timeline = SlideAnimationTimeline(slide_number=slide_number)

        # Topological sort for deterministic execution order
        adj = defaultdict(list)
        in_degree = defaultdict(int)
        for act_id in self.actions:
            in_degree[act_id] = 0
        for dep in self.dependencies:
            adj[dep.source_action_id].append(dep.target_action_id)
            in_degree[dep.target_action_id] += 1

        queue = deque([act_id for act_id, deg in in_degree.items() if deg == 0])
        ordered_action_ids = []
        while queue:
            curr = queue.popleft()
            ordered_action_ids.append(curr)
            for neighbor in adj[curr]:
                in_degree[neighbor] -= 1
                if in_degree[neighbor] == 0:
                    queue.append(neighbor)

        # Add remaining unconstrained actions
        for act_id in self.actions:
            if act_id not in ordered_action_ids:
                ordered_action_ids.append(act_id)

        for act_id in ordered_action_ids:
            act = self.actions[act_id]
            shape = shape_mapping.get(act.target_node_id)
            if shape is None:
                continue

            timeline.add_effect(
                shape=shape,
                effect_type=act.effect_type,
                duration_ms=act.duration_ms,
                delay_ms=act.delay_ms,
                trigger=act.trigger,
                direction=act.direction or "up",
                rotation_deg=act.rotation_deg,
                scale_factor=act.scale_factor,
                motion_path=act.motion_path,
                auto_reverse=act.auto_reverse,
                repeat_count=act.repeat_count,
            )

        return timeline
