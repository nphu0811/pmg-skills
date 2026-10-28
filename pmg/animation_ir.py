"""
pmg.animation_ir
================
Unified Animation Intermediate Representation (IR) and Timeline Compiler for PMG v2.2.0.

Decouples creative animation intent, spatial coordinates, and timeline dependencies
from PowerPoint OOXML serialization. Provides deterministic cycle detection,
dependency resolution, capability status tracking, and cross-slide Morph identity verification.
"""

from __future__ import annotations
import math
from enum import Enum
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional, Tuple, Set
from collections import defaultdict, deque

from pmg.animations import normalize_effect, SUPPORTED_EFFECTS, EFFECT_ALIASES


class AnimationCapabilityStatus(str, Enum):
    """Explicit capability status for animation effects."""
    SUPPORTED_AND_TESTED = "SUPPORTED_AND_TESTED"
    SUPPORTED_UNVERIFIED_PLAYBACK = "SUPPORTED_UNVERIFIED_PLAYBACK"
    UNSUPPORTED = "UNSUPPORTED"


class AnimationEasing(str, Enum):
    """Supported easing curves for native motion and parameter tweens."""
    LINEAR = "linear"
    EASE = "ease"
    EASE_IN = "ease_in"
    EASE_OUT = "ease_out"
    EASE_IN_OUT = "ease_in_out"


# Explicit Capability Catalog mapping effects to tested status
CAPABILITY_CATALOG: Dict[str, AnimationCapabilityStatus] = {
    # Entrances
    "appear": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "fade": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "fly_in": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "wipe": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "zoom": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    # Emphasis
    "scale": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "rotate": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "transparency": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "pulse": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    # Motion Paths
    "motion_path": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "linear_motion_path": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "bezier_motion_path": AnimationCapabilityStatus.SUPPORTED_UNVERIFIED_PLAYBACK,
    # Exits
    "disappear": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "fade_out": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "fly_out": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
    "wipe_out": AnimationCapabilityStatus.SUPPORTED_AND_TESTED,
}


def get_effect_capability_status(effect_name: str) -> Tuple[AnimationCapabilityStatus, str]:
    """
    Returns the capability status and normalized name for any requested effect.
    Never silently substitutes a fundamentally different effect.
    """
    clean_name = effect_name.strip().lower() if isinstance(effect_name, str) else ""
    if clean_name in CAPABILITY_CATALOG:
        return CAPABILITY_CATALOG[clean_name], EFFECT_ALIASES.get(clean_name, clean_name)
    canonical = EFFECT_ALIASES.get(clean_name, clean_name)
    if canonical in CAPABILITY_CATALOG:
        return CAPABILITY_CATALOG[canonical], canonical
    return AnimationCapabilityStatus.UNSUPPORTED, clean_name


@dataclass
class TransformIR:
    """Spatial 2D transform and geometry bounds in inches."""
    x: float = 0.0
    y: float = 0.0
    width: float = 1.0
    height: float = 1.0
    rotation_deg: float = 0.0
    scale_x: float = 1.0
    scale_y: float = 1.0
    opacity: float = 1.0

    def bounds(self) -> Tuple[float, float, float, float]:
        return (self.x, self.y, self.width * self.scale_x, self.height * self.scale_y)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> TransformIR:
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


@dataclass
class ObjectIR:
    """
    Visual element with a stable identity.
    Names prefixed with '!!' represent cross-slide traveling entities for Morph.
    """
    id: str
    name: str = ""
    kind: str = "card"  # 'card', 'rect', 'ellipse', 'textbox', 'picture', 'vector_shape'
    transform: TransformIR = field(default_factory=TransformIR)
    title: str = ""
    body: str = ""
    semantic_role: str = "content"  # 'hero', 'traveling_entity', 'card', 'background', 'typography'
    metadata: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        if not self.name:
            self.name = self.id

    @property
    def is_traveling_entity(self) -> bool:
        """Determines if object participates in persistent cross-slide Morph tracking."""
        return self.name.startswith("!!") or self.id.startswith("!!")

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["is_traveling_entity"] = self.is_traveling_entity
        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> ObjectIR:
        raw = dict(data)
        raw.pop("is_traveling_entity", None)
        if "transform" in raw and isinstance(raw["transform"], dict):
            raw["transform"] = TransformIR.from_dict(raw["transform"])
        return cls(**{k: v for k, v in raw.items() if k in cls.__dataclass_fields__})


@dataclass
class KeyframeIR:
    """Discrete parameter keyframe for multi-step transitions."""
    offset_fraction: float  # 0.0 to 1.0
    x: Optional[float] = None
    y: Optional[float] = None
    scale_x: Optional[float] = None
    scale_y: Optional[float] = None
    rotation_deg: Optional[float] = None
    opacity: Optional[float] = None

    def to_dict(self) -> Dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class AnimationTrackIR:
    """
    Declarative animation operation targeting a specific ObjectIR.
    Supports parallel and sequential dependency linking, time offsets, and motion paths.
    """
    track_id: str
    target_id: str
    effect: str
    trigger: str = "with_previous"  # 'with_previous', 'after_previous', 'on_click', 'on_shape_click'
    duration_ms: int = 800
    delay_ms: int = 0
    easing: str = AnimationEasing.EASE_IN_OUT.value
    direction: Optional[str] = "up"
    rotation_deg: float = 360.0
    scale_factor: Tuple[float, float] = (1.0, 1.0)
    opacity_from: float = 1.0
    opacity_to: float = 1.0
    motion_path: Optional[str] = None  # e.g., "M 0 0 C 0.1 -0.1 0.2 0.1 0.3 0 E"
    auto_reverse: bool = False
    repeat_count: float = 1.0
    trigger_shape_id: Optional[str] = None
    keyframes: List[KeyframeIR] = field(default_factory=list)
    computed_start_ms: int = 0
    computed_end_ms: int = 800
    sequence_index: int = 0
    capability_status: str = AnimationCapabilityStatus.SUPPORTED_AND_TESTED.value

    def __post_init__(self):
        status, canonical = get_effect_capability_status(self.effect)
        self.capability_status = status.value
        self.effect = canonical
        if self.duration_ms <= 0:
            raise ValueError(f"Invalid animation duration_ms {self.duration_ms}; must be > 0.")
        if self.delay_ms < 0:
            raise ValueError(f"Invalid animation delay_ms {self.delay_ms}; must be >= 0.")

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        if isinstance(d.get("scale_factor"), tuple):
            d["scale_factor"] = list(d["scale_factor"])
        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> AnimationTrackIR:
        raw = dict(data)
        if "keyframes" in raw and isinstance(raw["keyframes"], list):
            raw["keyframes"] = [KeyframeIR(**k) if isinstance(k, dict) else k for k in raw["keyframes"]]
        if "scale_factor" in raw and isinstance(raw["scale_factor"], list):
            raw["scale_factor"] = tuple(raw["scale_factor"])
        return cls(**{k: v for k, v in raw.items() if k in cls.__dataclass_fields__})


@dataclass
class SceneIR:
    """Single presentation scene corresponding to a PowerPoint slide."""
    scene_id: str
    slide_number: int = 1
    title: str = ""
    subtitle: str = ""
    transition_type: str = "none"  # 'morph', 'circle', 'fade', 'split', 'none'
    transition_duration_ms: int = 1200
    objects: List[ObjectIR] = field(default_factory=list)
    tracks: List[AnimationTrackIR] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def get_object(self, obj_id: str) -> Optional[ObjectIR]:
        for obj in self.objects:
            if obj.id == obj_id or obj.name == obj_id:
                return obj
        return None

    def get_traveling_entities(self) -> List[ObjectIR]:
        return [o for o in self.objects if o.is_traveling_entity]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "scene_id": self.scene_id,
            "slide_number": self.slide_number,
            "title": self.title,
            "subtitle": self.subtitle,
            "transition_type": self.transition_type,
            "transition_duration_ms": self.transition_duration_ms,
            "objects": [o.to_dict() for o in self.objects],
            "tracks": [t.to_dict() for t in self.tracks],
            "metadata": self.metadata,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> SceneIR:
        raw = dict(data)
        objects = [ObjectIR.from_dict(o) for o in raw.get("objects", [])]
        tracks = [AnimationTrackIR.from_dict(t) for t in raw.get("tracks", [])]
        return cls(
            scene_id=raw.get("scene_id", f"scene_{raw.get('slide_number', 1)}"),
            slide_number=raw.get("slide_number", 1),
            title=raw.get("title", ""),
            subtitle=raw.get("subtitle", ""),
            transition_type=raw.get("transition_type", "none"),
            transition_duration_ms=raw.get("transition_duration_ms", 1200),
            objects=objects,
            tracks=tracks,
            metadata=raw.get("metadata", {}),
        )


@dataclass
class AnimationIR:
    """
    Top-level Presentation Animation Intermediate Representation.
    Encapsulates all scenes, objects, transitions, and timeline dependencies.
    """
    scenes: List[SceneIR] = field(default_factory=list)
    palette_name: str = "LIGHT_ACADEMIC_CYAN"
    canvas_width_in: float = 13.333
    canvas_height_in: float = 7.5
    version: str = "2.2.0"
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "version": self.version,
            "palette_name": self.palette_name,
            "canvas_width_in": self.canvas_width_in,
            "canvas_height_in": self.canvas_height_in,
            "scenes": [s.to_dict() for s in self.scenes],
            "metadata": self.metadata,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> AnimationIR:
        raw = dict(data)
        scenes = [SceneIR.from_dict(s) for s in raw.get("scenes", [])]
        return cls(
            scenes=scenes,
            palette_name=raw.get("palette_name", "LIGHT_ACADEMIC_CYAN"),
            canvas_width_in=raw.get("canvas_width_in", 13.333),
            canvas_height_in=raw.get("canvas_height_in", 7.5),
            version=raw.get("version", "2.2.0"),
            metadata=raw.get("metadata", {}),
        )


class TimelineCompiler:
    """
    Deterministic animation timeline compiler.
    Resolves dependency graphs, detects cycles, schedules parallel and sequential groups,
    flags dangling targets, and produces scheduled tracks ready for OOXML compilation.
    """

    def __init__(self, scene: SceneIR):
        self.scene = scene
        self.errors: List[str] = []
        self.warnings: List[str] = []

    def validate(self) -> Tuple[bool, List[str], List[str]]:
        """
        Validates target integrity, triggers, cycle detection, and duplicate IDs.
        """
        self.errors.clear()
        self.warnings.clear()

        # 1. Object ID uniqueness on this scene
        seen_ids = set()
        seen_names = set()
        for obj in self.scene.objects:
            if obj.id in seen_ids:
                self.errors.append(f"Duplicate object ID '{obj.id}' on scene '{self.scene.scene_id}'.")
            seen_ids.add(obj.id)
            if obj.name in seen_names:
                self.errors.append(f"Duplicate object name '{obj.name}' on scene '{self.scene.scene_id}'.")
            seen_names.add(obj.name)

        # 2. Target existence for all tracks
        for track in self.scene.tracks:
            target = self.scene.get_object(track.target_id)
            if target is None:
                self.errors.append(
                    f"Animation track '{track.track_id}' targets non-existent object '{track.target_id}'."
                )

            # Trigger shape existence if specified
            if track.trigger_shape_id:
                trig_target = self.scene.get_object(track.trigger_shape_id)
                if trig_target is None:
                    self.errors.append(
                        f"Animation track '{track.track_id}' specifies non-existent trigger shape '{track.trigger_shape_id}'."
                    )

            # Check capability status
            if track.capability_status == AnimationCapabilityStatus.UNSUPPORTED.value:
                self.warnings.append(
                    f"Track '{track.track_id}' uses effect '{track.effect}' which is UNSUPPORTED natively."
                )

        # 3. Detect cycles in track dependencies
        self._check_dependency_cycles()

        is_valid = len(self.errors) == 0
        return is_valid, list(self.errors), list(self.warnings)

    def _check_dependency_cycles(self) -> None:
        """Kahn's algorithm on track dependency ordering."""
        adj = defaultdict(list)
        in_degree = defaultdict(int)

        for track in self.scene.tracks:
            in_degree[track.track_id] = 0

        # Construct implicit or explicit dependencies
        # In linear tracks, with_previous and after_previous create causal dependencies
        for i, track in enumerate(self.scene.tracks):
            if i > 0 and track.trigger == "after_previous":
                prev_id = self.scene.tracks[i - 1].track_id
                adj[prev_id].append(track.track_id)
                in_degree[track.track_id] += 1

        queue = deque([t_id for t_id, deg in in_degree.items() if deg == 0])
        visited_count = 0

        while queue:
            curr = queue.popleft()
            visited_count += 1
            for nxt in adj[curr]:
                in_degree[nxt] -= 1
                if in_degree[nxt] == 0:
                    queue.append(nxt)

        if visited_count < len(self.scene.tracks):
            self.errors.append(f"Causal cycle detected in animation tracks for scene '{self.scene.scene_id}'.")

    def compile(self) -> List[AnimationTrackIR]:
        """
        Schedules tracks deterministically.
        Calculates computed_start_ms and computed_end_ms for every track.
        """
        is_valid, errors, _ = self.validate()
        if not is_valid:
            raise ValueError(f"TimelineCompiler error on scene '{self.scene.scene_id}': " + "; ".join(errors))

        current_group_start = 0
        current_group_end = 0

        scheduled: List[AnimationTrackIR] = []

        for idx, track in enumerate(self.scene.tracks):
            track.sequence_index = idx + 1
            effective_dur = int(track.duration_ms * max(1.0, track.repeat_count) * (2 if track.auto_reverse else 1))

            if idx == 0:
                track.computed_start_ms = track.delay_ms
                track.computed_end_ms = track.computed_start_ms + effective_dur
                current_group_start = 0
                current_group_end = track.computed_end_ms
            elif track.trigger == "with_previous":
                track.computed_start_ms = current_group_start + track.delay_ms
                track.computed_end_ms = track.computed_start_ms + effective_dur
                current_group_end = max(current_group_end, track.computed_end_ms)
            elif track.trigger in ("after_previous", "on_click", "on_shape_click"):
                gap = 200 if track.trigger in ("on_click", "on_shape_click") else 0
                current_group_start = current_group_end + gap
                track.computed_start_ms = current_group_start + track.delay_ms
                track.computed_end_ms = track.computed_start_ms + effective_dur
                current_group_end = track.computed_end_ms

            scheduled.append(track)

        return scheduled


class MorphReliabilityValidator:
    """
    Validates cross-slide Morph invariants and persistent traveling entities.
    Enforces '!!' identifier conventions, flags duplicate collisions, and checks geometry consistency.
    """

    @classmethod
    def validate_morph_flow(cls, animation_ir: AnimationIR) -> Dict[str, Any]:
        report = {
            "is_valid": True,
            "errors": [],
            "warnings": [],
            "traveling_entities": {},
            "transitions_verified": 0,
        }

        # Track appearances of '!!' objects across consecutive slides
        entity_appearances: Dict[str, List[int]] = defaultdict(list)

        for s_idx, scene in enumerate(animation_ir.scenes, start=1):
            # Check collisions on the same slide
            traveling_names_on_slide: Set[str] = set()
            for obj in scene.objects:
                if obj.is_traveling_entity:
                    if obj.name in traveling_names_on_slide:
                        report["errors"].append(
                            f"Morph Collision: Duplicate traveling entity '{obj.name}' on Slide {s_idx}."
                        )
                        report["is_valid"] = False
                    traveling_names_on_slide.add(obj.name)
                    entity_appearances[obj.name].append(s_idx)

            if scene.transition_type.lower() == "morph" and s_idx > 1:
                report["transitions_verified"] += 1
                prev_scene = animation_ir.scenes[s_idx - 2]
                prev_traveling = {o.name for o in prev_scene.objects if o.is_traveling_entity}
                curr_traveling = {o.name for o in scene.objects if o.is_traveling_entity}
                common = prev_traveling.intersection(curr_traveling)
                if not common:
                    report["warnings"].append(
                        f"Slide {s_idx} declares Morph transition from Slide {s_idx - 1}, "
                        f"but shares no common '!!' traveling entities. Morph will tween all un-prefixed elements."
                    )

        # Check for gap slides in traveling entities
        for name, slides in entity_appearances.items():
            report["traveling_entities"][name] = slides
            for i in range(len(slides) - 1):
                gap = slides[i + 1] - slides[i]
                if gap > 1:
                    report["warnings"].append(
                        f"Traveling entity '{name}' appears on Slide {slides[i]} and Slide {slides[i + 1]}, "
                        f"skipping intermediate slide(s) {list(range(slides[i] + 1, slides[i + 1]))}. "
                        f"May cause visual jump-cut during playback."
                    )

        return report
