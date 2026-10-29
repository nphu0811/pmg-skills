"""
pmg.motion_planners.base
========================
Base architecture and Shared Motion Vocabulary for PMG v2.5.0 Universal Object-Specific Motion Intelligence.

Defines the shared motion vocabulary, motion intent representation, and abstract
planner contracts used by all specialized object motion planners.
"""

from __future__ import annotations
from abc import ABC, abstractmethod
from enum import Enum
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any, Set, Tuple

from pmg.motion_graph import AnimationAction
from pmg.animation_ir import AnimationTrackIR, AnimationEasing


class SharedMotionPrimitive(str, Enum):
    """
    Universal Motion Vocabulary shared across all specialized object planners.
    Each semantic intent maps to these primitives, which are then compiled
    into native PowerPoint OOXML.
    """
    REVEAL = "reveal"          # Clean entrance, unmasking, or curtain draw
    DRAW = "draw"              # Progressive line or connector tracing along vector path
    GROW = "grow"              # Directional dimensional expansion (e.g. bar height, progress)
    TRANSFORM = "transform"    # Morph-based geometric state transformation
    MOVE = "move"              # Translation along linear or curved trajectory
    SCALE = "scale"            # Symmetrical or dimensional zoom / pop
    ROTATE = "rotate"          # Angular spin or orbiting movement
    HIGHLIGHT = "highlight"    # Luminance, tint, or color accentuation
    FOCUS = "focus"            # Camera push-in, dimming surroundings, or focal isolation
    TRACK = "track"            # Dynamic follower motion (e.g. label tracking bar peak)
    COUNT = "count"            # Numerical counter progression
    MORPH = "morph"            # Persistent cross-slide identity morphing (!! naming)
    STAGGER = "stagger"        # Cascading micro-delayed group animations
    SEQUENCE = "sequence"      # Explicit causal chronological ordering (A -> B -> C)
    EMPHASIZE = "emphasize"    # Attention grabber (pulse, scale pop)
    EXIT = "exit"              # Graceful departure from the active stage


@dataclass
class MotionIntent:
    """
    Semantic motion request produced by an object motion planner.
    Bridges high-level narrative purpose with native animation compilation.
    """
    target_id: str
    primitive: SharedMotionPrimitive
    semantic_purpose: str
    duration_ms: int = 800
    delay_ms: int = 0
    trigger: str = "after_previous"  # 'after_previous', 'with_previous', 'on_click'
    easing: str = "ease_out"         # 'linear', 'ease', 'ease_in', 'ease_out', 'ease_in_out'
    direction: Optional[str] = None  # 'up', 'down', 'left', 'right', 'in', 'out'
    scale_factor: float = 1.0
    rotation_deg: float = 0.0
    parameters: Dict[str, Any] = field(default_factory=dict)
    narrative_importance: str = "primary"  # 'hero', 'primary', 'secondary', 'supporting', 'background'

    def to_dict(self) -> Dict[str, Any]:
        res = asdict(self)
        res["primitive"] = self.primitive.value
        return res

    def to_animation_action(self, action_id: Optional[str] = None) -> AnimationAction:
        """
        Compiles the semantic MotionIntent into a native PMG AnimationAction for MotionGraph.
        """
        act_id = action_id or f"action_{self.target_id}_{self.primitive.value}"
        
        # Map shared primitive to native PowerPoint effect type
        effect_type = "appear"
        direction = self.direction
        scale_val = self.scale_factor
        rot_val = self.rotation_deg

        if self.primitive == SharedMotionPrimitive.REVEAL:
            effect_type = "fade"
        elif self.primitive == SharedMotionPrimitive.DRAW:
            effect_type = "wipe"
            if not direction:
                direction = "right"
        elif self.primitive == SharedMotionPrimitive.GROW:
            effect_type = "wipe"
            if not direction:
                direction = "up"
        elif self.primitive == SharedMotionPrimitive.MOVE:
            effect_type = "motion_path"
        elif self.primitive == SharedMotionPrimitive.SCALE:
            effect_type = "scale"
        elif self.primitive == SharedMotionPrimitive.ROTATE:
            effect_type = "rotate"
        elif self.primitive == SharedMotionPrimitive.HIGHLIGHT:
            effect_type = "pulse"
        elif self.primitive == SharedMotionPrimitive.FOCUS:
            effect_type = "zoom"
        elif self.primitive == SharedMotionPrimitive.COUNT:
            # Numeric counter compiles to crisp stepped entrance or reveal
            effect_type = "appear"
        elif self.primitive == SharedMotionPrimitive.EMPHASIZE:
            effect_type = "scale"
            scale_val = max(1.1, self.scale_factor)
        elif self.primitive == SharedMotionPrimitive.EXIT:
            effect_type = "fade_out"
        elif self.primitive in (SharedMotionPrimitive.MORPH, SharedMotionPrimitive.TRANSFORM):
            # Handled via slide-level Morph transition
            effect_type = "appear"
        else:
            effect_type = "appear"

        return AnimationAction(
            action_id=act_id,
            target_node_id=self.target_id,
            effect_type=effect_type,
            duration_ms=self.duration_ms,
            delay_ms=self.delay_ms,
            trigger=self.trigger,
            direction=direction,
            scale_factor=scale_val,
            rotation_deg=rot_val,
            parameters=self.parameters,
        )

    def to_animation_track_ir(self, track_id: Optional[str] = None) -> AnimationTrackIR:
        """
        Compiles the semantic MotionIntent into a v2.2.0 AnimationTrackIR.
        """
        tid = track_id or f"track_{self.target_id}_{self.primitive.value}"
        act = self.to_animation_action(tid)
        
        easing_enum = AnimationEasing.EASE_OUT
        try:
            easing_enum = AnimationEasing(self.easing)
        except ValueError:
            pass

        return AnimationTrackIR(
            id=tid,
            target_id=self.target_id,
            effect=act.effect_type,
            trigger=self.trigger,
            duration_ms=self.duration_ms,
            delay_ms=self.delay_ms,
            easing=easing_enum,
            direction=self.direction,
            scale_factor=self.scale_factor,
            rotation_deg=self.rotation_deg,
            parameters=self.parameters,
        )


class BaseMotionPlanner(ABC):
    """
    Abstract contract for specialized object-specific motion intelligence.
    """
    @property
    @abstractmethod
    def planner_name(self) -> str:
        """Name of the specialized planner (e.g. 'pmg-chart-motion')."""
        pass

    @property
    @abstractmethod
    def supported_object_types(self) -> Set[str]:
        """Set of supported object types or categories."""
        pass

    @abstractmethod
    def plan_motion(self, context: Dict[str, Any]) -> List[MotionIntent]:
        """
        Generates a sequence of semantic MotionIntent objects for the given context.
        """
        pass

    @abstractmethod
    def explain_selection(self, intent: MotionIntent, context: Dict[str, Any]) -> str:
        """
        Provides a transparent diagnostic explanation answering why this animation
        was selected based on semantic role, audience focus, and native compatibility.
        """
        pass
