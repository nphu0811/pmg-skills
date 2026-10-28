"""Canonical Storyboard Intelligence Engine for PMG Skills.

Plans narrative structure, diverse visual compositions, camera intentions,
asset requirements, and motion strategies before constructing individual slides.
"""
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional
import json


@dataclass
class StoryboardScene:
    """Scene specification defining visual composition, pacing, camera, and motion."""
    scene_number: int
    title: str
    subtitle: str = ""
    narrative_purpose: str = "establishing"  # 'establishing', 'tension', 'exploration', 'deep_dive', 'synthesis', 'climax'
    composition_style: str = "cinematic_hero"  # 'cinematic_hero', 'editorial_layout', 'radial_infographic', 'layered_cutaway', 'photo_carousel', 'continuous_one_shot', 'dashboard'
    camera_intent: str = "wide_establishing"   # 'wide_establishing', 'push_in_zoom', 'tracking_pan', 'macro_detail', 'pull_out'
    transition_to_next: str = "morph"          # 'morph', 'circle', 'split', 'fade', 'none'
    key_elements: List[str] = field(default_factory=list)
    asset_requirements: List[str] = field(default_factory=list)
    motion_intent: str = "gentle_float"
    pacing_seconds: float = 4.0
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class StoryboardPlan:
    """Complete presentation storyboard plan across multiple scenes."""
    topic: str
    purpose: str = "educational keynote"
    target_audience: str = "general"
    tone: str = "cinematic_inspiring"
    scenes: List[StoryboardScene] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def to_json(self, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), indent=indent, ensure_ascii=False)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "StoryboardPlan":
        scenes = [StoryboardScene(**s) for s in data.get("scenes", [])]
        data_copy = dict(data)
        data_copy["scenes"] = scenes
        return cls(**data_copy)


class StoryboardEngine:
    """Synthesizes structured narrative and visual motion plans."""
    COMPOSITION_ROTATION = [
        ("cinematic_hero", "wide_establishing", "split"),
        ("editorial_layout", "tracking_pan", "morph"),
        ("radial_infographic", "push_in_zoom", "morph"),
        ("layered_cutaway", "macro_detail", "morph"),
        ("photo_carousel", "tracking_pan", "morph"),
        ("continuous_one_shot", "push_in_zoom", "circle"),
    ]

    def plan_presentation(self, topic: str, num_slides: int = 5,
                           audience: str = "executive", tone: str = "modern_cinematic") -> StoryboardPlan:
        """Create a diverse, non-repetitive presentation plan."""
        scenes: List[StoryboardScene] = []
        for i in range(1, num_slides + 1):
            comp_idx = (i - 1) % len(self.COMPOSITION_ROTATION)
            comp_style, cam_intent, trans = self.COMPOSITION_ROTATION[comp_idx]

            if i == 1:
                purpose = "establishing"
                title = f"{topic.upper()} — VISION"
                subtitle = "Autonomous Space Exploration & Discovery"
                trans = "morph"
            elif i == num_slides:
                purpose = "climax"
                title = "HORIZONS EXPANDED"
                subtitle = "The Future of Interstellar Exploration"
                trans = "circle"
            else:
                purpose = "deep_dive"
                title = f"EXPLORATION PHASE 0{i}"
                subtitle = f"Strategic Trajectory & Systems Analysis {i}"

            scene = StoryboardScene(
                scene_number=i,
                title=title,
                subtitle=subtitle,
                narrative_purpose=purpose,
                composition_style=comp_style,
                camera_intent=cam_intent,
                transition_to_next=trans,
                key_elements=[f"Hero_Object_{i}", f"Data_Module_{i}", f"Traveling_Probe"],
                asset_requirements=[f"{topic.lower()}_visual_{i}"],
                motion_intent="dynamic_coordinate",
            )
            scenes.append(scene)

        return StoryboardPlan(
            topic=topic,
            purpose="presentation",
            target_audience=audience,
            tone=tone,
            scenes=scenes,
        )
