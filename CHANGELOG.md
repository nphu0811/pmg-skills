# Changelog — PMG Skills

## [v1.3.0] - 2026-10-28
### Added
- **Professional Cinematic Motion Graphics Builder (`pmg.cinematic_builder`):**
  - Implemented `CinematicMotionBuilder`, `MotionPresetRegistry`, and `MotionChoreographyDecision` to combine cross-slide Morph storytelling with intra-slide native PowerPoint Animation Pane timelines (`SlideAnimationTimeline`).
  - Added reusable purposeful motion presets: `cinematic_title_reveal`, `multi_plane_parallax` (0.7x BG, 1.0x Mid, 1.4x FG), `camera_zoom_inspection` (macro push-in 100% -> 200%), `animated_infographic_bars` (proportional bar wipes + curved Bezier trend orb), `orbiting_carousel_3d` (3-slot depth carousel), and `animated_interface_mockup` (satellite telemetry HUD with rotating radar).
- **10-Slide Cinematic Environmental Protection Builder (`build_v1_3_cinematic_environmental_deck`):**
  - Combines editable multi-layered Earth globe, blue sky, multi-plane green meadow landscapes, recurring flying Kite (`!!Traveling_Kite`), camera zoom sequences (Slides 3 & 7), animated infographics (Slides 4 & 8), 3D orbiting carousel (Slide 5), satellite HUD interface mockup (Slide 6), and 75 native Animation Pane effects across 10 slides.

## [v1.2.0] - 2026-10-28
### Added
- **Native Intra-Slide Object Animation & Declarative Timeline Engine (`pmg.animations`):**
  - Implemented `AnimationStep` and `SlideAnimationTimeline` supporting native PowerPoint Animation Pane effects:
    - Entrances: `appear`, `fade`, `fly_in`, `wipe`, `zoom`
    - Emphasis: `scale` / `grow_shrink` (`<p:animScale>`), `rotate` / `spin` (`<p:animRot>`), `transparency` (`style.opacity`), `pulse`
    - Motion Paths: linear and curved Bezier (`<p:animMotion path="M ... C ... E">`)
    - Exits: `disappear`, `fade_out`, `fly_out`, `wipe_out`
    - Timing & Triggers: `delay_ms`, `duration_ms`, `with_previous` (`withEffect`), `after_previous` (`afterEffect`), `on_click` (`clickEffect`), `on_shape_click` (`interactiveSeq`), `repeat_count`, `auto_reverse`, and multiple sequential/simultaneous effects targeting a single object.
  - Added `parse_slide_timing_openxml` and `verify_animation_pane_com` to verify all effects in PowerPoint's native `slide.TimeLine.MainSequence`.
- **Single-Slide Native Animation Showcase Builder (`build_v1_2_single_slide_animation_showcase`):**
  - 12 independent editable objects and 19 native PowerPoint Animation Pane effects (`pmg build --preset v1_2_showcase`).
- **Enhanced Video Pipeline, Reviewer & Self-Healing (`pmg.video_pipeline`, `pmg.visual_reviewer`, `pmg.self_healing`):**
  - Synchronized `GotoSlide(1, -1)` and timed `on_click` progression during live Slide Show capture.
  - Sub-frame intra-slide motion sampling (`anim_start`, `anim_mid`, `steady`) and per-object ROI delta tracking.
  - Automated repair for broken `<p:spTgt>` references, invalid `<p:animMotion>` paths, and missing `mainSeq` nodes.

## [v1.1.0] - 2026-10-28
### Added
- **Canonical Python Package (`pmg`) & CLI (`python -m pmg` / `pmg`):**
  - Implemented `pmg build`, `pmg validate`, `pmg render`, `pmg record`, `pmg review`, `pmg inspect`, and `pmg reconstruct`.
  - Added `pyproject.toml`, `setup.py`, and `requirements.txt`.
- **Main Skill Entrypoint (`skills/pmg-engine/SKILL.md`):**
  - Enforces Engine-First execution and final validated `.pptx` delivery policy.
- **Expected Animation Timeline (`pmg.models.ExpectedTimeline` & `pmg.validator.extract_expected_timeline`):**
  - Deterministic motion specification tracking object IDs, slide numbers, initial/final positions and sizes, rotations, opacities, durations, triggers, and click playback scripts.
- **Mandatory Real Playback & Video Validation Pipeline (`pmg.orchestrator.PMGOrchestrator`):**
  - Step 1: Canonical build with 100% native editable shapes/text.
  - Step 2: Deep OpenXML static validation (`StaticValidator`) and anti-shortcut verification (blocks flattened screenshot slides).
  - Step 3: Native PowerPoint video export (`powerpoint_export.mp4`) with `ffprobe` inspection.
  - Step 4: Actual Microsoft PowerPoint Desktop full-screen Slide Show playback recording (`actual_slideshow_recording.mp4`) synchronized with `expected_timeline.json`.
  - Step 5: AI & Computer Vision visual review (`VisualAnimationReviewer`) extracting keyframes (`frame_samples/`), verifying object movement trajectories, and checking visual, animation, temporal, and capture quality.
  - Step 6 & 7: Unified 3-Loop Self-Healing (`pmg.self_healing`) with fresh re-rendering and re-recording after every repair.
  - Step 8 & 9: Strict `PASS` delivery gate and `execution_manifest.json` provenance tracking.
- **Environmental Protection 10-Slide Engine Builder (`build_environmental_protection_deck`):**
  - Multi-layered native editable Earth globe, multi-plane green grassland/meadow landscapes, golden sun, frosted glass cards, and recurring flying Kite (`!!Traveling_Kite`).
- **Automated Test Suite (`tests/`):**
  - Comprehensive unit and integration tests for engine, OpenXML validation, timeline extraction, self-healing loops, and real playback review.

### Fixed
- Added missing `LightThemeCraftDeck` and `CinematicColors` in `scripts/motion_graph_kit.py` referenced by `README.md`.
- Added `add_traveling_entity` alias on `MotionDeck` fixing `scripts/build_showcase_deck.py` and `scripts/create_camau_deck.py`.
- Added `Cinematic3DDeck` in `scripts/forge_lib.py` referenced by `skills/pmg-camera-parallax/SKILL.md`.
- Fixed PowerPoint COM cleanup to avoid closing unrelated user presentations.
