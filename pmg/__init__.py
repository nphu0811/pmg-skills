"""
PMG Skills — Native PowerPoint Motion Graphics Engine (v1.7.0)
==============================================================
Canonical Python API, Windows Validation Worker, Animation Benchmark Suite,
and Orchestration Engine for PowerPoint Motion Graphics.
"""

from pmg.version import (
    VERSION,
    get_git_commit,
    get_engine_provenance,
    verify_runtime_dependencies,
)
from pmg.models import (
    ValidationStatus,
    ObjectMotionSpec,
    SlideTimelineSpec,
    ExpectedTimeline,
    VideoMetadata,
    ReviewDefect,
    AuditReport,
    ReviewReport,
    ExecutionManifest,
)
from pmg.animations import (
    AnimationStep,
    SlideAnimationTimeline,
    parse_slide_timing_openxml,
    verify_animation_pane_com,
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
    build_v1_2_single_slide_animation_showcase,
    SpotifyAppShellBuilder,
    EarthSpaceMotionBuilder,
    UniversityBloomBuilder,
)
from pmg.cinematic_builder import (
    MotionChoreographyDecision,
    MotionPresetRegistry,
    CinematicMotionBuilder,
    build_v1_3_cinematic_environmental_deck,
)
from pmg.reverse_engineer import (
    ExtractedObjectSpec,
    ExtractedAnimationSpec,
    SlideBlueprintSpec,
    PresentationMotionSpec,
    PPTXMotionReverseEngineer,
    build_v1_4_reference_smartphone_slide,
)
from pmg.video_reconstructor import (
    TrackedVideoObject,
    VideoSceneSpec,
    VideoReconstructionReport,
    VideoMotionReconstructor,
    generate_controlled_reference_video,
)
from pmg.validator import StaticValidator, extract_expected_timeline, verify_clean_powerpoint_open
from pmg.video_pipeline import (
    is_powerpoint_available,
    probe_video_metadata,
    export_edit_mode_slide_images,
    export_video_with_powerpoint,
    record_actual_slideshow,
)
from pmg.visual_reviewer import VisualAnimationReviewer
from pmg.self_healing import execute_repair_loop
from pmg.orchestrator import PMGOrchestrator
from pmg.worker import (
    WorkerHealthStatus,
    WorkerJobRequest,
    WorkerJobResult,
    PMGValidationWorkerServer,
    PMGWorkerClient,
    check_worker_health,
)
from pmg.visual_assets import (
    VisualAssetEngine,
    WikimediaProvider,
    NASAProvider,
    AssetRequirement,
    AssetRecord,
)
from pmg.layers import LayerTree, VisualObject, LayerMask
from pmg.image_decomposer import ImageDecomposer
from pmg.asset_builder import build_logo_assembly_deck, build_image_shutter_deck
from pmg.scene_graph import SceneGraph, SceneNode, Transform
from pmg.motion_graph import MotionGraph, AnimationAction, Dependency
from pmg.storyboard import StoryboardEngine, StoryboardPlan, StoryboardScene
from pmg.motion_director import MotionDirector
from pmg.compositing import (
    ParallaxPlane,
    PARALLAX_PLANES,
    build_bezier_motion_path,
    build_orbital_motion_path,
    build_waypoint_motion_path,
    generate_massive_constellation,
    build_v1_8_agribank_identity_deck,
    build_v1_8_cinematic_shutter_reveal,
    build_v1_8_deep_space_exploration,
)

__version__ = VERSION

__all__ = [
    "__version__",
    "VERSION",
    "get_git_commit",
    "get_engine_provenance",
    "verify_runtime_dependencies",
    "ValidationStatus",
    "ObjectMotionSpec",
    "SlideTimelineSpec",
    "ExpectedTimeline",
    "VideoMetadata",
    "ReviewDefect",
    "AuditReport",
    "ReviewReport",
    "ExecutionManifest",
    "AnimationStep",
    "SlideAnimationTimeline",
    "parse_slide_timing_openxml",
    "verify_animation_pane_com",
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
    "build_v1_2_single_slide_animation_showcase",
    "SpotifyAppShellBuilder",
    "EarthSpaceMotionBuilder",
    "UniversityBloomBuilder",
    "MotionChoreographyDecision",
    "MotionPresetRegistry",
    "CinematicMotionBuilder",
    "build_v1_3_cinematic_environmental_deck",
    "ExtractedObjectSpec",
    "ExtractedAnimationSpec",
    "SlideBlueprintSpec",
    "PresentationMotionSpec",
    "PPTXMotionReverseEngineer",
    "build_v1_4_reference_smartphone_slide",
    "TrackedVideoObject",
    "VideoSceneSpec",
    "VideoReconstructionReport",
    "VideoMotionReconstructor",
    "generate_controlled_reference_video",
    "StaticValidator",
    "extract_expected_timeline",
    "verify_clean_powerpoint_open",
    "is_powerpoint_available",
    "probe_video_metadata",
    "export_edit_mode_slide_images",
    "export_video_with_powerpoint",
    "record_actual_slideshow",
    "VisualAnimationReviewer",
    "execute_repair_loop",
    "PMGOrchestrator",
    "WorkerHealthStatus",
    "WorkerJobRequest",
    "WorkerJobResult",
    "PMGValidationWorkerServer",
    "PMGWorkerClient",
    "check_worker_health",
    "VisualAssetEngine",
    "WikimediaProvider",
    "NASAProvider",
    "AssetRequirement",
    "AssetRecord",
    "LayerTree",
    "VisualObject",
    "LayerMask",
    "ImageDecomposer",
    "build_logo_assembly_deck",
    "build_image_shutter_deck",
    "SceneGraph",
    "SceneNode",
    "Transform",
    "MotionGraph",
    "AnimationAction",
    "Dependency",
    "StoryboardEngine",
    "StoryboardPlan",
    "StoryboardScene",
    "MotionDirector",
    "ParallaxPlane",
    "PARALLAX_PLANES",
    "build_bezier_motion_path",
    "build_orbital_motion_path",
    "build_waypoint_motion_path",
    "generate_massive_constellation",
    "build_v1_8_agribank_identity_deck",
    "build_v1_8_cinematic_shutter_reveal",
    "build_v1_8_deep_space_exploration",
]
