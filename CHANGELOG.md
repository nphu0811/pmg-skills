# Changelog — PMG Skills

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

