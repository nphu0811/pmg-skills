# PMG development and release gates

This continues the existing master plan and the responsibility audit. It does not
replace completed v1.1–v1.5.1 features. The latest released tag at the start of
this work was **v1.5.1**. The implementation target is **v1.5.2**; release tags
require the evidence below. No later version is released or certified by this document.

The preexisting compatibility fixes and Spotify, Earth/Space and university
showcase builders are retained. Their earlier uncommitted v1.6.0 labels did not
constitute a release. Reliability changes belong to v1.5.2; the asset research
release remains v1.6.0.

## Fixed release sequence

| Version | Implementation and ownership | Acceptance and release gate |
| --- | --- | --- |
| v1.5.2 — active | OPC/relationship/XML integrity; strict effects and aliases; read-only structured audits; one repair budget; final byte-state playback and clean reopen; preserve existing APIs/presets | Negative manifest fixture, existing regressions and 24 benchmarks; 20-slide native space deck; measured native export and actual slideshow recording; review and clean final reopen; push, install exact published commit, regenerate and reverify; only then tag |
| v1.6.0 — gated | Add only pmg-visual-assets and pmg-image-decomposer. Shared asset/layer records with provenance, licenses, editability, parent IDs, masks, geometry and z-order. Reuse video_reconstructor segmentation through a common interface; preserve its decoding/tracking | Real retrieval or honestly unavailable provider; generation is separate fallback. Semantic/granular logo layers and reassembly; photograph strips/sectors; native logo/shutter demonstrations; prior regression and actual playback; publish/install/retest/tag |
| v1.7.0 — gated | Add only pmg-storyboard. One canonical Scene Graph and Motion Graph. Upgrade choreography into Motion Director. Consolidate Morph compilation and camera transforms behind compatible facades; reuse pmg.animations and cinematic_builder presets | Storyboard accepted by director; decomposed layers adapt directly to Scene Graph; hierarchy/world transforms/masks validated; graph dependencies compile to actual shape IDs; choreography invokes Morph compiler and 3D invokes shared camera; diverse scene planning; publish/install/playback/retest/tag |
| v1.8.0 — gated | Advanced compositing, configurable motion patterns, camera tracking, model/cutaway composition and reference reconstruction. Extend existing reverse_engineer, video_reconstructor and benchmarks; no second compiler or benchmark suite | Preserve all 24 benchmarks and add the 20 advanced cases from the master prompt. Logo, image shutter and complex space demonstrations; meaningful hierarchies, resource-aware performance, original-reference availability recorded; end-to-end playback, final reopen and installed-commit retest before release |
| v2.0.1 — completed | Creative Motion Intelligence; extend choreography, storyboard, shared graphs/camera and existing compiler. Provider-agnostic model planning and critique with deterministic offline fallback; distinguish provenance. Composable motion grammar, alternative concepts, selection, creative revision budgets and long one-shot direction | At least three substantially different concepts from one prompt; schema-validated custom choreography; real 20–30-second one-slide space scene; editable logo one-shot and concept-comparison demonstrations; extend existing suite with all 20 creative benchmarks; technical repairs still share the maximum-three-attempt budget; publish/install/regenerate/playback/reopen before v2.0.1 tag |
| v2.1.0 — completed | Universal Runtime & Cross-Platform Reliability: deterministic capability detection (`pmg.capabilities`), offline-first 6-tier dependency resolution (`pmg.dependencies`), consolidated native OOXML compiler without PowerPoint dependency (`pmg.compiler`), Three-Level Validation (Level 1 Structural, Level 2 Visual, Level 3 Native Playback in `pmg.validation`), unified public API (`PMGRuntime`), and CLI `doctor`/`capabilities`/`build`/`validate` commands | Deterministic capability detection across web sandbox / Linux / macOS / Windows; native PPTX compilation without PowerPoint Desktop; offline dependency priority verified; Three-Level Validation automated with honest reporting; zero-repair COM verification on Windows; all unit and integration tests passing |
| v2.2.0 — completed | Native Animation Compiler: declarative unified Animation IR (`pmg.animation_ir`), explicit capability status catalog (`SUPPORTED_AND_TESTED`, `SUPPORTED_UNVERIFIED_PLAYBACK`, `UNSUPPORTED`), deterministic `TimelineCompiler` with Kahn's algorithm cycle detection, deterministic `MorphReliabilityValidator` with cross-slide identity and duplicate `!!` collision checks, native `<p:timing>` synthesis without PowerPoint Desktop | Full test suite passing; deterministic timeline ordering and dependency scheduling; cycle detection; invalid target rejection; morph identity collision rejection; unverified playback effects tracked in compiler report |
| v2.3.0 — completed | Distributed PowerPoint QA: dedicated Windows PowerPoint automation worker (`pmg.worker`), hardened REST server with Bearer token authentication, 100MB payload guardrail, idempotent SHA-256 artifact job caching, lifecycle tracking (`QUEUED`, `PROCESSING`, `COMPLETED`, `FAILED`, `CANCELLED`), and orchestrator remote/local QA delegation | Server health check; Bearer token rejection (401); payload limit rejection (413); job cancellation; idempotent cache hits without duplicate COM execution; remote worker error handling; real PowerPoint COM slideshow validation |
| v2.4.0 — completed | Self-Healing Compatibility & Immutable Certification: structured `ErrorCode` diagnostics (`pmg.diagnostics`), safe automatic repairs for missing ContentTypes, duplicate shape IDs, and dangling animation targets, transactional repair manager with automatic rollback on failure, bounded 3-loop retry limit, and cryptographic `ImmutableValidationCertificate` bound to artifact SHA-256 | Structured defect classification; safe automatic repairs verified; transactional rollback on unresolvable errors; certificate invalidation on byte tampering; clean COM open verification with `DisplayAlerts=ppAlertsAll`; 133/133 tests passing |

Do not start the next implementation version while the active release gate is
incomplete. Do not ask for confirmation between versions when prerequisites and
existing authorization suffice. A missing prerequisite is BLOCKED/NOT_VERIFIED,
never an invented PASS. Record the exact blocker and preserve progress.

## Canonical data flow and contracts

User request → Storyboard → asset research → decomposition → Scene Graph →
Motion Director → Motion Graph → native compiler → QA audit → review/repair →
orchestrator delivery gate.

v1.6 layer records must preserve stable identity, parent/child relationships,
geometry and units, mask references, z-order, layer identity, pivot/transform
origin, source attribution and editability. v1.7 Scene Graph and Motion Graph
will be the only authoritative scene/motion schemas; existing extracted PPTX and
tracked-video records adapt to them rather than becoming competing graphs.
Motion references must resolve to canonical object IDs, then to the actual
slide's PowerPoint shape IDs in pmg.animations. Reject unknown targets, cycles,
invalid masks and unsupported effects explicitly. Intentional overlap is allowed;
protect readable regions and detect accidental collisions.

The Motion Director decides what, why and when; Morph/camera/animation modules
compile how. pmg-3d-cinematic owns models and model-focused composition, and uses
the shared camera. Native GLB support must be tested on supported PowerPoint;
documentation examples do not establish engine support.

## Architecture acceptance allocation

| Additional requirement | First owning release |
| --- | --- |
| 4: QA does not mutate PPTX; 5: review handles repairs/delivery; 6: one three-loop orchestration | v1.5.2 |
| 11: existing public functionality and 12: fresh skill executes canonical engine | Every release, beginning v1.5.2 |
| 7: decomposition/Scene Graph compatibility; 8: video shares decomposition | Shared layer interface in v1.6.0; graph adapter in v1.7.0 |
| 1: Morph owns compilation; 2: choreography invokes shared Morph; 3: 3D shares camera | v1.7.0 |
| 9: storyboard accepted by director; 10: graph targets resolve to PowerPoint IDs | v1.7.0 |
| Advanced cross-skill benchmarks and end-to-end complex playback | v1.8.0 |
| Creative concepts, model adapter, critique and continuous one-shot direction | v2.0.1 |

## v2.0.1 retained creative specification

Plan from topic, narrative objective, audience, emotional tone, hierarchy,
duration, assets, object relationships, references, complexity and technical
constraints. Choose stationary and moving subjects, retrieval/decomposition,
simultaneity/sequencing, compositing, camera and transitions purposefully.

Compose reveal/conceal, translate/rotate/scale, orbit/follow, split/merge,
assemble/disassemble, scatter/converge, slice/mask, expand/contract,
cascade/stagger, transform/focus/track, reverse/synchronize/repeat primitives.
Do not implement a second timeline compiler or create separate graph/director
skills. Existing presets remain configurable building blocks.

Generate and compare distinct feasible concepts. Assess clarity, originality,
hierarchy, composition, coherence, rhythm, asset suitability, native editability,
reliability and audience comprehension. Model-backed structured generation and
critique are optional through a provider-agnostic adapter; tests need no private
keys. Label deterministic fallback honestly and validate all output against the
same graph schemas. A numeric score does not prove beauty or artistic superiority.

Support meaningful long one-slide scenes: independent and simultaneous paths,
rotation, scaling, groups, camera-follow simulation, depth parallax, slicing,
assembly and typography. Compile child trajectories when PowerPoint cannot
inherit parent animation directly. Avoid arbitrary object/effect/file-size/time
caps; measure resource and playback limits and do not inflate complexity.

Creative pre-compilation revisions use a bounded planning budget. Technical
post-compilation defects use the existing single maximum-three-attempt repair
budget. Distinguish creative weakness from broken XML, missing objects, invalid
targets, missing media, repair prompts or playback failures. Record whether
critique was model-generated, deterministic or manually reviewed.

Extend benchmarks with: alternatives, strategy diversity, semantic/purposeful
planning, custom choreography, a 20-second one-shot, coordinated objects,
camera/object simultaneity, hierarchy, logo reassembly, shutter/radial/carousel
motion, diverse multi-slide composition, graph compilation, reference-guided
design, critique/revision, Animation Pane editability, compatibility and actual
playback. Separate planning correctness, compiler correctness, execution,
similarity and subjective judgment.

Required demonstrations: Agribank logo cinematic one-shot from the actual logo
when available; one physical 20–30-second space slide (Earth, spacecraft,
starfield, nebula/atmosphere, HUD and typography); three creative concepts from
one prompt, documented selection, and the selected native editable PPTX.

## Release evidence record

For each version record commit SHA, GitHub branch, tag, implemented features,
changed files, unit/integration/regression results, actual compatibility, measured
video metadata, remaining limitations and missing original reference files.
Keep videos, frames and diagnostic reports in internal_qa. Ordinary presentation
users receive only the final validated PPTX. No source-code change or static test
can certify an already-generated deck's playback without that evidence.
