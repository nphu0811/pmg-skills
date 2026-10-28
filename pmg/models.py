"""
pmg.models
==========
Data models for expected animation timelines, validation status, video metadata,
AI visual review reports, repair history, and execution manifests.
"""

import json
import os
from enum import Enum
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional, Tuple


class ValidationStatus(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    BLOCKED = "BLOCKED"
    NOT_VERIFIED = "NOT_VERIFIED"


@dataclass
class ObjectMotionSpec:
    object_id: str
    slide_number: int
    object_type: str  # "shape", "textbox", "picture", "group", "connector", "svg"
    initial_position: Tuple[float, float]  # (x_inches, y_inches)
    final_position: Tuple[float, float]    # (x_inches, y_inches)
    initial_size: Tuple[float, float]      # (w_inches, h_inches)
    final_size: Tuple[float, float]        # (w_inches, h_inches)
    rotation_params: Dict[str, float] = field(default_factory=lambda: {"start_deg": 0.0, "end_deg": 0.0})
    opacity_params: Dict[str, float] = field(default_factory=lambda: {"start_alpha": 1.0, "end_alpha": 1.0})
    motion_path: Optional[str] = None
    start_time_ms: int = 0
    duration_ms: int = 1600
    delay_ms: int = 0
    trigger_type: str = "slide_transition"  # "slide_transition", "with_previous", "after_previous", "on_click"
    animation_sequence: int = 0
    easing: str = "smooth_morph"
    time_reference: str = "slide_relative"  # "slide_relative", "absolute", "event_triggered"
    absolute_start_ms: int = 0
    absolute_end_ms: int = 1600

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class SlideTimelineSpec:
    slide_number: int
    transition_type: str = "morph"
    transition_duration_ms: int = 1600
    hold_duration_ms: int = 2000
    absolute_start_ms: int = 0
    absolute_end_ms: int = 3600
    objects: List[ObjectMotionSpec] = field(default_factory=list)
    clicks: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "slide_number": self.slide_number,
            "transition_type": self.transition_type,
            "transition_duration_ms": self.transition_duration_ms,
            "hold_duration_ms": self.hold_duration_ms,
            "absolute_start_ms": self.absolute_start_ms,
            "absolute_end_ms": self.absolute_end_ms,
            "objects": [o.to_dict() for o in self.objects],
            "clicks": self.clicks,
        }


@dataclass
class ExpectedTimeline:
    deck_title: str
    engine_version: str
    total_slides: int
    total_duration_ms: int = 0
    slides: List[SlideTimelineSpec] = field(default_factory=list)
    playback_script: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "deck_title": self.deck_title,
            "engine_version": self.engine_version,
            "total_slides": self.total_slides,
            "total_duration_ms": self.total_duration_ms,
            "slides": [s.to_dict() for s in self.slides],
            "playback_script": self.playback_script,
        }

    def save(self, path: str) -> str:
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
        return path

    @classmethod
    def load(cls, path: str) -> "ExpectedTimeline":
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        slides = []
        for s in data.get("slides", []):
            objs = [
                ObjectMotionSpec(
                    object_id=o["object_id"],
                    slide_number=o["slide_number"],
                    object_type=o["object_type"],
                    initial_position=tuple(o["initial_position"]),
                    final_position=tuple(o["final_position"]),
                    initial_size=tuple(o["initial_size"]),
                    final_size=tuple(o["final_size"]),
                    rotation_params=o.get("rotation_params", {"start_deg": 0.0, "end_deg": 0.0}),
                    opacity_params=o.get("opacity_params", {"start_alpha": 1.0, "end_alpha": 1.0}),
                    motion_path=o.get("motion_path"),
                    start_time_ms=o.get("start_time_ms", 0),
                    duration_ms=o.get("duration_ms", 1600),
                    delay_ms=o.get("delay_ms", 0),
                    trigger_type=o.get("trigger_type", "slide_transition"),
                    animation_sequence=o.get("animation_sequence", 0),
                    easing=o.get("easing", "smooth_morph"),
                    time_reference=o.get("time_reference", "slide_relative"),
                    absolute_start_ms=o.get("absolute_start_ms", 0),
                    absolute_end_ms=o.get("absolute_end_ms", 1600),
                )
                for o in s.get("objects", [])
            ]
            slides.append(
                SlideTimelineSpec(
                    slide_number=s["slide_number"],
                    transition_type=s.get("transition_type", "morph"),
                    transition_duration_ms=s.get("transition_duration_ms", 1600),
                    hold_duration_ms=s.get("hold_duration_ms", 2000),
                    absolute_start_ms=s.get("absolute_start_ms", 0),
                    absolute_end_ms=s.get("absolute_end_ms", 3600),
                    objects=objs,
                    clicks=s.get("clicks", []),
                )
            )
        return cls(
            deck_title=data.get("deck_title", "PMG Presentation"),
            engine_version=data.get("engine_version", "1.1.0"),
            total_slides=data.get("total_slides", len(slides)),
            total_duration_ms=data.get("total_duration_ms", 0),
            slides=slides,
            playback_script=data.get("playback_script", []),
        )


@dataclass
class VideoMetadata:
    path: str
    exists: bool = False
    decodable: bool = False
    width: int = 0
    height: int = 0
    r_frame_rate: str = "0/1"
    avg_frame_rate: str = "0/1"
    actual_fps: float = 0.0
    requested_fps: int = 60
    nb_frames: int = 0
    duration_sec: float = 0.0
    size_bytes: int = 0
    codec_name: str = ""
    dropped_frames: int = 0
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ReviewDefect:
    category: str       # "visual_quality", "animation_behavior", "temporal_quality", "video_capture_quality", "openxml_structure"
    severity: str       # "critical", "major", "minor", "warning"
    slide_number: int
    object_id: Optional[str]
    description: str
    remediation_hint: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ReviewReport:
    status: ValidationStatus
    defects: List[ReviewDefect] = field(default_factory=list)
    visual_quality_passed: bool = False
    animation_behavior_passed: bool = False
    temporal_quality_passed: bool = False
    capture_quality_passed: bool = False
    slide_steady_states_verified: int = 0
    transitions_verified: int = 0
    tracked_entities: Dict[str, Any] = field(default_factory=dict)
    temporal_deviation_ms: float = 0.0
    geometric_deviation_pct: float = 0.0
    export_vs_recording_similarity: float = 0.0
    summary: str = ""

    def has_blocking_defects(self) -> bool:
        return any(d.severity in ("critical", "major") for d in self.defects)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": self.status.value if isinstance(self.status, ValidationStatus) else self.status,
            "visual_quality_passed": self.visual_quality_passed,
            "animation_behavior_passed": self.animation_behavior_passed,
            "temporal_quality_passed": self.temporal_quality_passed,
            "capture_quality_passed": self.capture_quality_passed,
            "slide_steady_states_verified": self.slide_steady_states_verified,
            "transitions_verified": self.transitions_verified,
            "tracked_entities": self.tracked_entities,
            "temporal_deviation_ms": round(self.temporal_deviation_ms, 2),
            "geometric_deviation_pct": round(self.geometric_deviation_pct, 2),
            "export_vs_recording_similarity": round(self.export_vs_recording_similarity, 4),
            "defects": [d.to_dict() for d in self.defects],
            "summary": self.summary,
        }

    def save(self, path: str) -> str:
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
        return path


@dataclass
class ExecutionManifest:
    run_id: str
    provenance: Dict[str, Any]
    source_pptx: str
    delivered_pptx: Optional[str]
    qa_dir: str
    status: ValidationStatus
    static_validation: Dict[str, Any] = field(default_factory=dict)
    export_video: Dict[str, Any] = field(default_factory=dict)
    recording_video: Dict[str, Any] = field(default_factory=dict)
    review_summary: Dict[str, Any] = field(default_factory=dict)
    repair_loops_executed: int = 0
    repair_history: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "run_id": self.run_id,
            "provenance": self.provenance,
            "source_pptx": self.source_pptx,
            "delivered_pptx": self.delivered_pptx,
            "qa_dir": self.qa_dir,
            "status": self.status.value if isinstance(self.status, ValidationStatus) else self.status,
            "static_validation": self.static_validation,
            "export_video": self.export_video,
            "recording_video": self.recording_video,
            "review_summary": self.review_summary,
            "repair_loops_executed": self.repair_loops_executed,
            "repair_history": self.repair_history,
        }

    def save(self, path: str) -> str:
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
        return path

