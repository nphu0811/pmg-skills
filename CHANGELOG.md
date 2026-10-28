# Changelog — PMG Skills

## [v1.8.0] - 2026-10-29

Advanced Cinematic Compositing, Motion Reconstruction & Professional Benchmarks release. Expands benchmark suite from 24 to 44 cases and delivers three production-grade cinematic motion presentations.

### Added
- **Advanced Cinematic Compositing Engine (`pmg.compositing`):**
  - 6-layer multi-plane parallax velocity scaling (`deep_background` 0.2x, `background` 0.5x, `midground` 1.0x, `foreground` 1.4x, `floating_particles` 1.8x, `hud_overlay` 0.0x fixed/counter-tracking).
  - OpenXML curved cubic Bezier motion path (`build_bezier_motion_path`), elliptical orbital trajectories (`build_orbital_motion_path`), and multi-segment waypoint paths (`build_waypoint_motion_path`).
  - Massive scene hierarchy synthesis supporting 50, 100, 200+ addressable objects per slide (`generate_massive_constellation`).
  - Group dynamics and cascades: staggered reveals, radial blossom cascades, wave animations, and micro-pulsing/rotation.
- **Three Mandatory Acceptance Presentations:**
  - **Demo 1: Agribank Motion Identity System (8 slides):** Particle reassembly reveal, heritage split cards, concentric branch network, digital 3D frosted cards, ESG green agriculture wave, dynamic bar chart metrics, 4-pillar radial carousel, and circle wipe closing.
  - **Demo 2: Cinematic Shutter & Perspective Reveal (5 slides):** 8-strip vertical shutter reveal, 240% perspective macro zoom, horizontal slice cascade, frosted glass depth stack, and iris resolution.
  - **Demo 3: Deep Space Exploration — Beyond the Horizon (12 slides):** Flagship presentation with >100 objects on key slides, authentic NASA media integration, Jovian gravity assist, Saturn ring plane crossing, heliopause boundary, and 12-slide traveling spacecraft hero entity.
- **Expanded 44-Case Animation Benchmark Suite (`benchmarks/`):**
  - Added Group E (Cases 25..29): Orbital loops, radial blossoms, waypoint paths, micro-pulse/spin, and oscillating trajectories.
  - Added Group F (Cases 30..34): Massive hierarchies (50, 100, 200 nodes), 6-layer deep parallax, and counter-tracking HUD.
  - Added Group G (Cases 35..39): Decomposed particle scatter/gather, 10-strip shutter, radial sector iris, vector polygon morph, and frosted glass depth.
  - Added Group H (Cases 40..44): 15-second multi-act intra-slide timelines, storyboard-to-director compilation, 4-scene traveling entity, multi-target synchronized parallel tracks, and clean PowerPoint COM open verification.
- **Clean COM Open Verification:**
  - All presentations certified with Microsoft PowerPoint Desktop COM `DisplayAlerts = ppAlertsAll` (`2`) and `OpenAndRepair = False` with zero repair prompts.

## [v1.7.0] - 2026-10-29

Storyboard Intelligence, Scene Graph, Motion Graph & AI Motion Director release. Adds canonical skill pmg-storyboard and upgrades pmg-choreography with hierarchical scene composition and motion dependency graphs.

### Added
- **pmg-storyboard (skills/pmg-storyboard/SKILL.md, pmg.storyboard):**
  - High-level narrative architecture and scene planning (StoryboardEngine, StoryboardPlan, StoryboardScene).
  - Narrative pacing, visual hierarchy, thematic traveling hero continuity across scenes, camera framing, and asset requirement specification.
- **Hierarchical Scene Graph (pmg.scene_graph):**
  - Scene tree data structure (SceneGraph, SceneNode, Transform) with parent-child relative transforms and world coordinate resolution.
  - Semantic role definitions (ackground, card, subject, 	ypography, hud, overlay).
  - Intelligent overlap rules: allows intentional design layering (masks, glass cards, shadows) while strictly rejecting accidental readable text collisions (protected_readable).
  - Native PowerPoint compilation (compile_to_slide) preserving shape hierarchy and !! persistent identifiers.
- **Motion Graph & Dependency Resolution (pmg.motion_graph):**
  - Directed acyclic graph (MotionGraph, AnimationAction, Dependency) defining temporal animation relationships (fter_previous, with_previous).
  - Cycle detection and topological sorting for robust chronological compilation.
  - Native compilation to SlideAnimationTimeline with automatic target shape mapping.
- **AI Motion Director (pmg.motion_director):**
  - Coordinates storyboard intent with SceneGraph and MotionGraph synthesis.
  - Diverse cinematic composition templates (cinematic_hero, editorial_layout, 
adial_infographic, multi_card).
  - End-to-end compilation into certified editable PowerPoint decks.
- **Acceptance Verification:**
  - Automated 5-slide Interstellar Navigation Systems presentation verified via Microsoft PowerPoint Desktop COM clean open with 0 alerts and full animation timeline.

## [v1.6.0] - 2026-10-29

Intelligent Visual Asset Research & Image Decomposition release. Adds two canonical modular skills: `pmg-visual-assets` and `pmg-image-decomposer`.

### Added
- **`pmg-visual-assets` (`skills/pmg-visual-assets/SKILL.md`, `pmg.visual_assets`):**
  - Source-first media research pipeline discovering authentic photography, illustrations, icons, and videos from Wikimedia Commons and NASA archives.
  - Rigorous rights, license, provenance, and trademark classification; distinct separation of retrieval vs. generation fallback.
  - Asset quality scoring (resolution, contrast, transparency, suitability for decomposition).
- **`pmg-image-decomposer` (`skills/pmg-image-decomposer/SKILL.md`, `pmg.image_decomposer`, `pmg.layers`):**
  - Semantic decomposition isolating meaningful brand components (e.g., Agribank emblem into green background, red diagonal bands, white separators, stem, and granular wheat grains).
  - Explicit editability classification: `native_vector` (OpenXML `<a:custGeom>` custom polygons) vs `raster_layer` (transparent PNG).
  - Geometric image slicing: vertical strips, horizontal strips, grids, tiles, circular/radial sectors, puzzle, triangles, and custom polygons with zero dropped pixels.
  - Depth decomposition: spatial partitioning into foreground, midground, and background layers with explicit inferred/occlusion flags.
  - Automatic reassembly validation (`verify_reassembly` measuring pixel similarity >= 0.94 and contour alignment).
- **Native PowerPoint Object Compilation (`pmg.native_objects`):**
  - Compiles structured `LayerTree` into native PowerPoint shapes, freeform custom geometries, and cropped picture layers preserving `!!` persistent identifiers across slides.
- **Acceptance Demonstrations:**
  - Demo A: Agribank Logo Motion 7-slide assembly presentation (`build_logo_assembly_deck`), verified clean COM open and 96 native animations.
  - Demo B: Image Shutter Reveal presentation (`build_image_shutter_deck`), verified clean COM open.
- **Shared Segmentation Architecture:**
  - Unified `ForegroundSegmenter` and `ColorRegionSegmenter` in `pmg.segmentation`, reused by both still-image decomposition and `VideoMotionReconstructor`.
- **Empty Animation Node Safeguard:**
  - `SlideAnimationTimeline.compile_to_openxml` safely strips empty `<p:timing>` trees to eliminate PowerPoint modal stalls.

## [v1.5.2] - 2026-10-29

Reliability release; publication requires exact-commit reinstall and playback
retesting. The existing showcase work is retained; asset research remains v1.6.0.

### Added
- Read-only shared AuditReport/ReviewDefect inspection with complete OPC XML,
  ContentType and relationship checks, including duplicate/undefined parts.
- Legacy audit/repair scripts now forward to one shared validator/orchestrator;
  compatible fallback remains inside the maximum-three-attempt repair budget.
- Strict effect alias normalization, unknown-effect rejection and compilation
  checks against actual native PowerPoint target IDs.
- Metadata packaging before complete playback review; exact final byte hash and
  destination reopen verification; existing output retained on a failed gate.
- Bounded read-only worker probe with alerts enabled, repair disabled, modal and
  Protected View detection, source hashes and native Animation Pane counts.
  Remote evidence must bind to the exact uploaded PPTX; missing evidence blocks.
- Actual recording telemetry required by CLI review; expected timestamps can no
  longer masquerade as recorded events.
- Responsibility audit retained internally; synchronized ownership/routing/API
  guidance and gated roadmap through v2.0.1. Intentional overlap is supported.
- **OPC-Compliant Manifest Embedding & Companion Manifest Output:**
  - Added explicit OPC registration (`<Default Extension="json" ContentType="application/json"/>` in `[Content_Types].xml`) when embedding execution manifests in `.pptx` (`ppt/pmg_execution_manifest.json`), eliminating PowerPoint's "PowerPoint found a problem... Repair presentation" dialog.
  - Automatically exports companion external manifest `<pptx_path>.manifest.json` alongside the delivered presentation for external tool and audit consumption.
  - Enhanced `StaticValidator` to scan all parts in the ZIP package against `[Content_Types].xml` and flag unmapped parts as critical `openxml_structure` defects (`remediation_hint="declare_part_content_type"`).
- **Post-Validation Freeze & PowerPoint COM Clean Open Verification Gate:**
  - Strict delivery gate invariant: The `.pptx` deliverable is NEVER modified after marking validation `PASS`. Packaging and manifest embedding occur before final validation.
  - Added `verify_clean_powerpoint_open(pptx_path)`: Opens the final packaged presentation in Microsoft PowerPoint Desktop with `DisplayAlerts = ppAlertsAll` (`2`, not `1`) to verify zero repair dialogs or corruption warnings.
  - Awards official `status == "PASS"` only after the exact final deliverable has passed the complete static, playback, AI visual review, and COM open inspection suite, sealed externally with its SHA-256 fingerprint. Embedded provenance remains NOT_VERIFIED until interpreted with that external certificate.
- **Iconic Reference Cinematic Showcase Engine (`pmg.cinematic_showcase`):**
  - Reverse-engineered and recreated iconic slide patterns from `DataPowerPoint`:
    - **`SpotifyAppShellBuilder` (`--preset spotify`):** Spotify UI Shell & Radial Morph unlock from `MAC-LENIN.pptx` (ambient orb monolith -> lateral title convergence -> macOS 3-dot left sidebar + top bar + action pills -> in-frame viewport stream -> team grid).
    - **`EarthSpaceMotionBuilder` (`--preset earth_space`):** Earth & Space 21-slide cinematic sequence from `EARTH AND SPACE.pptx` (cosmic establishing shot -> Moon orbit pan -> targeting reticle / viewfinder corner brackets `[ ]` target lock -> expansion into 3 framed highlight cards -> "BUT SOMETHING IS MISSING" focal zoom -> 4-quadrant rotating segmented radar wheel).
    - **`UniversityBloomBuilder` (`--preset university_bloom`):** Concentric rings blooming into 8 radial bezier petals and 8-item two-column Mục Lục from `PROMPT.pdf` / `PROMPT-CHUYỂN-NỘI-DUNG.docx`.
  - Enforced **Universal Topic Adaptability**: All presets dynamically adapt to any user topic, title, subtitle, author, and color theme (light modern defaults or deep obsidian modes), strictly rejecting rigid single-row or single-domain layout constraints.
- **Proportional Timeline Video Review Alignment:**
  - Upgraded `pmg.visual_reviewer` to compute steady-state sampling in `powerpoint_export.mp4` via proportional timeline scaling (`exp_scale = exp_total_dur / expected_total_sec`) rather than naive linear splits, eliminating false capture-quality defects when slide transition durations vary.
  - Added fallback comparison against PowerPoint's native edit-mode slide PNG exports (`edit_mode_refs/edit_slide_XX.png`).
- **Preserved Delivery and Showcase Tests:**
  - Added `tests/test_delivery_and_showcase.py` verifying OPC ContentType registration, companion manifests, COM clean open, and showcase builders.


## [v1.5.1] - 2026-10-28
### Added
- **Upgrade A — Windows PowerPoint Validation Worker (`pmg.worker`):**
  - Implemented `pmg.worker` (`server.py`, `client.py`, `jobs.py`, `powerpoint_runner.py`, `video_capture.py`, `result_schema.py`, `security.py`) enabling AI agents running on Linux, macOS, cloud containers, or Windows to execute real Microsoft PowerPoint playback validation either locally or via a remote Windows worker.
  - Added environment variable support (`PMG_QA_WORKER_URL`, `PMG_QA_WORKER_TOKEN`, `PMG_QA_MODE=local|remote|auto|static`) and CLI commands (`pmg worker doctor`, `pmg worker serve`, `pmg worker status`, `pmg validate <file.pptx> --worker auto|local|remote|static`).
  - Added security controls: constant-time Bearer token authentication, max file size limits, ZIP bomb & path-traversal rejection, macro/OLE/executable payload rejection (`.pptm`, `vbaProject.bin`, `oleObject`), isolated per-job temporary workspaces, and secret redaction.
  - Unconfigured or unreachable worker in `remote` mode (or `auto` mode without local PowerPoint) returns `ValidationStatus.BLOCKED` (never `PASS`).
- **Upgrade B — 24-Case Animation Benchmark Suite (`benchmarks/`):**
  - Implemented 24 canonical animation benchmark cases across Group A (Cases 01–08: Fade In, Fade Out, Fly In/Out, Wipe, Rotation, Scale/Grow-Shrink, Linear Motion Path, Curved Bezier Motion Path), Group B (Cases 09–14: Simultaneous, Sequential, With Previous, After Previous, On Click, Multiple effects on one object), Group C (Cases 15–20: Morph position, Morph scale/rotation, Traveling object across slides, Camera zoom, Multi-plane parallax, Complex overlapping timeline), and Group D (Cases 21–24: Hero image visibility with outline, Title/text visibility, Multiple animated images, Failure detection & automatic repair).
  - Added CLI commands: `pmg benchmark run --all`, `pmg benchmark run --case <id>`, and `pmg benchmark report`, outputting `benchmarks/reports/benchmark_report.json` and `benchmarks/reports/benchmark_report.md`.
- **Upgrade C — Engine Distribution, Runtime Dependency Verification & Embedded Provenance:**
  - Added `verify_runtime_dependencies()` in `pmg.version` and embedded `ppt/pmg_execution_manifest.json` inside every delivered `.pptx` archive recording `engine_version`, `git_commit`, `canonical_entrypoint`, `validation_method`, `worker_identity`, `repair_loops_executed`, and `final_status`.
- **Upgrade D — Visibility-Aware Animation Validation & Disappearing Object Self-Healing:**
  - Reproduced and fixed the real-world regression from `Vu_tru_20_slide_PMG_inspired.pptx` where hero images (`!!Earth_Grand`) and animated titles (`!!HeroTitle`) vanished during Slide Show while sibling decorative outlines (`Oval 13`) remained visible.
  - Exported static edit-mode slide reference PNGs (`edit_mode_refs/edit_slide_XX.png`) and implemented inner-core ROI visibility verification in `VisualAnimationReviewer` (insetting hero image ROIs by 18% so outer decorative rings like `Oval 13` never mask missing inner content).
  - Enhanced `StaticValidator` and `_repair_disappearing_object_timing` in `pmg.self_healing` to detect and repair missing `<p:set>` (`style.visibility = "visible"`), missing `grpId` / `<p:bldLst>` bindings, untriggered `mainSeq` `<p:par>` siblings, and duplicate conflicting entrance animations while **preserving all legitimate Fade In, Wipe, Fly In, and Motion Path animations**.
  - Added regression test suite `tests/regression/test_disappearing_hero_image.py` and the 20-slide flagship builder `pmg.space_builder.build_v1_5_1_space_exploration_20_slide_deck` (**"SPACE EXPLORATION — BEYOND THE HORIZON"**).

## [v1.5.0] - 2026-10-28
### Added
- **Video-to-Editable-PowerPoint Motion Reconstruction Engine (`pmg.video_reconstructor`):**
  - Implemented `VideoMotionReconstructor`, `VideoReconstructionReport`, `VideoSceneSpec`, `TrackedVideoObject`, and `generate_controlled_reference_video`.
  - 12-stage CV & motion graphics pipeline: video decoding (`cv2.VideoCapture`), HSV histogram scene boundary detection, keyframe extraction, foreground contour & connected-component segmentation, multi-frame centroid & bounding-box trajectory tracking, scale/rotation/opacity estimation, timing & trigger extraction, cubic Bezier motion path curve fitting (`M 0 0 C ... E`), easing estimation, 100% native editable PowerPoint shape synthesis, and native `<p:timing>` Animation Pane compilation.
  - Strictly prohibits frame-flattening or video-background shortcuts; every reconstructed slide contains independent editable PowerPoint objects and native Animation Pane sequences verified via COM.
  - Explicit engineering-honesty reporting (`directly_observed_motion`, `estimated_motion`, `inferred_visual_properties`, `unsupported_effects`, and `approximate_reconstructions`).

## [v1.4.0] - 2026-10-28
### Added
- **PPTX Motion Reverse Engineering & Animation Style Transfer (`pmg.reverse_engineer`):**
  - Implemented `PPTXMotionReverseEngineer`, `PresentationMotionSpec`, `SlideBlueprintSpec`, `ExtractedObjectSpec`, and `ExtractedAnimationSpec`.
  - Extracts slide geometry, object hierarchy, z-order, shape/text/group properties, semantic roles (`badge`, `title`, `hero_device`, `ui_card`, `accent_node`, `status_bar`), transitions, Morph anchors, and native `<p:timing>` animation trees (effect type, direction, duration, delay, trigger, parallel group ID, Bezier motion paths, rotation, and scale factors).
  - Implemented `reconstruct_and_transfer` and `compare_reference_and_reconstructed` to recreate editable slides with transferred native Animation Pane effects (`15/15` exact matches) and content/palette style transfer.
  - Added `build_v1_4_reference_smartphone_slide` reference showcase (Smartphone flying in, 360° rotation, scale-up, curved Bezier signal orb trajectory, and staggered UI component reveals).
  - Wired `pmg reconstruct <source.pptx> -o <output.pptx> [--spec-out ...] [--compare-out ...] [--transfer-title ...]` into `pmg.cli`.

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
