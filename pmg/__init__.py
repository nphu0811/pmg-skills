"""
PMG Skills — Native PowerPoint Motion Graphics Engine (v1.1.0)
==============================================================
Canonical Python API and Orchestration Engine for PowerPoint Motion Graphics.
"""

from pmg.version import VERSION, get_git_commit, get_engine_provenance
from pmg.models import (
    ValidationStatus,
    ObjectMotionSpec,
    SlideTimelineSpec,
    ExpectedTimeline,
    VideoMetadata,
    ReviewDefect,
    ReviewReport,
    ExecutionManifest,
)
from pmg.engine import (
    Themes,
    CinematicColors,
    MotionDeck,
    LightThemeCraftDeck,
    IndustrialPalette,
    ForgeDeck,
    Cinematic3DDeck,
    apply_transition,
    apply_morph,
    link_morph,
    export_deck_to_video,
    add_editable_earth,
    add_editable_meadow_landscape,
    add_editable_kite,
    build_environmental_protection_deck,
)
from pmg.validator import StaticValidator, extract_expected_timeline
from pmg.video_pipeline import (
    is_powerpoint_available,
    probe_video_metadata,
    export_video_with_powerpoint,
    record_actual_slideshow,
)
from pmg.visual_reviewer import VisualAnimationReviewer
from pmg.self_healing import execute_repair_loop
from pmg.orchestrator import PMGOrchestrator

__version__ = VERSION

__all__ = [
    "__version__",
    "VERSION",
    "get_git_commit",
    "get_engine_provenance",
    "ValidationStatus",
    "ObjectMotionSpec",
    "SlideTimelineSpec",
    "ExpectedTimeline",
    "VideoMetadata",
    "ReviewDefect",
    "ReviewReport",
    "ExecutionManifest",
    "Themes",
    "CinematicColors",
    "MotionDeck",
    "LightThemeCraftDeck",
    "IndustrialPalette",
    "ForgeDeck",
    "Cinematic3DDeck",
    "apply_transition",
    "apply_morph",
    "link_morph",
    "export_deck_to_video",
    "add_editable_earth",
    "add_editable_meadow_landscape",
    "add_editable_kite",
    "build_environmental_protection_deck",
    "StaticValidator",
    "extract_expected_timeline",
    "is_powerpoint_available",
    "probe_video_metadata",
    "export_video_with_powerpoint",
    "record_actual_slideshow",
    "VisualAnimationReviewer",
    "execute_repair_loop",
    "PMGOrchestrator",
]

