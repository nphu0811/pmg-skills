"""
PMG Skills — Universal Runtime & Native PowerPoint Motion Graphics Engine (v2.1.0)
=================================================================================
Canonical Python API, Universal Capability Detection, Offline-First Architecture,
Native OOXML Compiler, Three-Level Validation & Self-Healing System.
"""

from pmg.version import (
    VERSION,
    get_git_commit,
    get_engine_provenance,
    verify_runtime_dependencies,
)
from pmg.capabilities import (
    CapabilityState,
    ExecutionMode,
    EnvironmentCapabilities,
    detect_environment,
)
from pmg.dependencies import (
    DependencyResolution,
    resolve_dependencies,
    find_compatible_offline_wheels,
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
from pmg.compiler import (
    NativePPTXCompiler,
    CompilerReport,
)
from pmg.animation_ir import (
    AnimationCapabilityStatus,
    AnimationEasing,
    AnimationIR,
    SceneIR,
    ObjectIR,
    KeyframeIR,
    AnimationTrackIR,
    TransformIR,
    TimelineCompiler,
    MorphReliabilityValidator,
    get_effect_capability_status,
)
from pmg.validation import (
    Level1Result,
    Level2Result,
    Level3Result,
    ComprehensiveValidationReport,
    validate_level1_structural,
    validate_level2_visual,
    validate_level3_playback,
    validate_presentation,
)
from pmg.runtime import (
    PMGRuntime,
    PMGBuildResult,
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
from pmg.validator import StaticValidator, extract_expected_timeline, verify_clean_powerpoint_open
from pmg.self_healing import (
    execute_repair_loop,
    repair_missing_content_types,
    repair_duplicate_shape_ids,
    repair_broken_animation_targets,
    TransactionalRepairManager,
    RepairSessionResult,
    ImmutableValidationCertificate,
    issue_validation_certificate,
    verify_certificate,
)
from pmg.diagnostics import (
    ErrorCode,
    RepairSafety,
    StructuredDiagnostic,
    DiagnosticClassifier,
)
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
from pmg.creative_director import (
    CreativeEvaluationAxis,
    CreativeConcept,
    CreativeCritique,
    ModelPlanningAdapter,
    CreativeDirector,
    build_v2_0_oneshot_space_scene,
    build_v2_0_agribank_oneshot_deck,
    build_v2_0_creative_concept_selection_deck,
)

# Optional heavy CV / Media modules guarded for web sandboxes lacking cv2/scipy
try:
    from pmg.video_reconstructor import (
        TrackedVideoObject,
        VideoSceneSpec,
        VideoReconstructionReport,
        VideoMotionReconstructor,
        generate_controlled_reference_video,
    )
except ImportError:
    TrackedVideoObject = None  # type: ignore
    VideoSceneSpec = None  # type: ignore
    VideoReconstructionReport = None  # type: ignore
    VideoMotionReconstructor = None  # type: ignore
    generate_controlled_reference_video = None  # type: ignore

try:
    from pmg.video_pipeline import (
        is_powerpoint_available,
        probe_video_metadata,
        export_edit_mode_slide_images,
        export_video_with_powerpoint,
        record_actual_slideshow,
    )
    from pmg.visual_reviewer import VisualAnimationReviewer
except ImportError:
    is_powerpoint_available = lambda: False  # type: ignore
    probe_video_metadata = None  # type: ignore
    export_edit_mode_slide_images = None  # type: ignore
    export_video_with_powerpoint = None  # type: ignore
    record_actual_slideshow = None  # type: ignore
    VisualAnimationReviewer = None  # type: ignore

try:
    from pmg.image_decomposer import ImageDecomposer
except ImportError:
    ImageDecomposer = None  # type: ignore

__version__ = VERSION

__all__ = [
    "__version__",
    "VERSION",
    "get_git_commit",
    "get_engine_provenance",
    "verify_runtime_dependencies",
    # Universal Runtime & Capabilities (v2.1.0)
    "PMGRuntime",
    "PMGBuildResult",
    "CapabilityState",
    "ExecutionMode",
    "EnvironmentCapabilities",
    "detect_environment",
    "DependencyResolution",
    "resolve_dependencies",
    "find_compatible_offline_wheels",
    "NativePPTXCompiler",
    "CompilerReport",
    "Level1Result",
    "Level2Result",
    "Level3Result",
    "ComprehensiveValidationReport",
    "validate_level1_structural",
    "validate_level2_visual",
    "validate_level3_playback",
    "validate_presentation",
    # Core Models & Animations
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
    "CreativeEvaluationAxis",
    "CreativeConcept",
    "CreativeCritique",
    "ModelPlanningAdapter",
    "CreativeDirector",
    "build_v2_0_oneshot_space_scene",
    "build_v2_0_agribank_oneshot_deck",
    "build_v2_0_creative_concept_selection_deck",
    "AnimationCapabilityStatus",
    "AnimationEasing",
    "AnimationIR",
    "SceneIR",
    "ObjectIR",
    "KeyframeIR",
    "AnimationTrackIR",
    "TransformIR",
    "TimelineCompiler",
    "MorphReliabilityValidator",
    "get_effect_capability_status",
    "ErrorCode",
    "RepairSafety",
    "StructuredDiagnostic",
    "DiagnosticClassifier",
    "repair_missing_content_types",
    "repair_duplicate_shape_ids",
    "repair_broken_animation_targets",
    "TransactionalRepairManager",
    "RepairSessionResult",
    "ImmutableValidationCertificate",
    "issue_validation_certificate",
    "verify_certificate",
]
