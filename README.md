# PMG Skills — PowerPoint Motion Graphics (v2.4.0)

PMG creates 100% native, individually editable PowerPoint presentations through a universal, cross-platform Python runtime.
It combines 60 FPS Morph transitions, native Animation Pane timing trees, frosted glassmorphism layers, virtual camera parallax, and continuous traveling entities.

Current implementation: **v2.4.0 — Universal Runtime, Native Animation Compiler, Distributed PowerPoint QA & Self-Healing Compatibility**.

---

## What's New in v2.4.0

### 1. Universal Runtime & Offline-First Strategy (v2.1.0)
- **PowerPoint-Independent Native Compilation:** Native PowerPoint slide decks, Morph transitions, and Animation Pane effects compile without Microsoft PowerPoint.
- **Three Execution Environments:** Mode A (Web/Sandbox), Mode B (Local Desktop), Mode C (Remote Worker).
- **Capability Detection:** `pmg doctor` and `pmg capabilities` determine platform limits deterministically.
- **Three-Level Validation Architecture:**
  * **Level 1 (Structural):** Mandatory ZIP, XML, OPC ContentTypes, relationships, shape IDs, `<p:timing>` trees, and Morph metadata.
  * **Level 2 (Visual):** Slide previews rendered via available backend (PowerPoint COM or LibreOffice). Reports `SKIPPED` if absent.
  * **Level 3 (Playback):** Native Microsoft PowerPoint Slide Show playback. Reports `NOT_TESTED` if PowerPoint is absent. *Never fabricates Level 3 PASS without genuine PowerPoint Desktop execution.*

### 2. Native Animation Compiler & Unified IR (v2.2.0)
- **Declarative Animation Intermediate Representation (`pmg.animation_ir`):** Decouples creative intent from OpenXML syntax (`SceneIR`, `ObjectIR`, `AnimationTrackIR`, `TransformIR`, `KeyframeIR`).
- **Explicit Capability Catalog:** Distinguishes `SUPPORTED_AND_TESTED`, `SUPPORTED_UNVERIFIED_PLAYBACK`, and `UNSUPPORTED` effects without silent substitutions.
- **Deterministic Timeline Compiler:** Kahn's algorithm cycle detection, dependency resolution, and timing graph validation.
- **Morph Reliability Validator:** Enforces `!!` traveling entity persistent identifiers, rejects duplicate name collisions, and warns of slide gaps.

### 3. Distributed PowerPoint QA (v2.3.0)
- **Remote Validation Worker (`pmg.worker`):** Secure Windows PowerPoint COM automation service for environments without local PowerPoint.
- **Hardened HTTP REST Server:** Bearer token authentication, 100MB payload guardrails, and graceful cancellation.
- **Idempotent Job Manager:** SHA-256 artifact hashing and caching to prevent redundant PowerPoint COM runs.
- **Lifecycle Tracking:** `QUEUED` -> `PROCESSING` -> `COMPLETED` / `FAILED` / `CANCELLED`.

### 4. Self-Healing Compatibility & Immutable Certification (v2.4.0)
- **Structured ErrorCode Diagnostics (`pmg.diagnostics`):** Standardized taxonomy classifying defects into `SAFE_AUTOMATIC`, `REQUIRES_REVIEW`, or `UNSUPPORTED`.
- **Transactional Repair Pipeline (`pmg.self_healing`):** Operates on isolated `.tmp.pptx` working copies with cycle detection and automatic atomic rollback on failure.
- **Safe Automatic Repairs:** Missing OPC ContentTypes in `[Content_Types].xml`, duplicate shape `cNvPr id` renumbering with target preservation, and dangling animation target `<p:spTgt>` cleanup.
- **Tamper-Evident Immutable Validation Certificate:** Cryptographically bound to the final artifact SHA-256 hash. Any byte modification invalidates the certificate immediately.
- **Mandatory Clean COM Open Verification:** Evaluated with `DisplayAlerts = ppAlertsAll` (`2`) to guarantee zero "Repair presentation" prompts.

---

## Quickstart

### Python API

```python
from pmg import PMGRuntime

# 1. Initialize runtime with auto-detected capabilities
runtime = PMGRuntime.auto()

# 2. Build presentation with safe transactional self-healing
result = runtime.build(
    project="presentation_spec.json",
    output="presentation.pptx",
    validate=True,
    repair="safe",
)

print(f"Output: {result.output_path}")
print(f"Status: {result.validation_report.overall_status}")
print(f"Artifact SHA-256: {result.artifact_sha256}")
```

### Command-Line Interface (CLI)

```bash
# Environment diagnostics and capabilities
pmg doctor
pmg capabilities

# Build presentation from JSON specification or preset
pmg build presentation_spec.json -o presentation.pptx --repair safe
pmg build --preset space --slides 10 -o Space.pptx

# Full Three-Level validation
pmg validate presentation.pptx

# Distributed QA validation (local COM or remote worker)
pmg qa presentation.pptx
pmg qa presentation.pptx --remote --timeout 60

# Safe transactional self-healing repair
pmg repair presentation.pptx --safe -o presentation_repaired.pptx
```

---

## Installation & Packaging

### Core Engine (Zero Binary Dependencies)
For basic slide compilation, Morph transitions, and Level 1 structural validation, only pure-python dependencies are required:
```bash
pip install "python-pptx>=0.6.21" "lxml>=4.9.0" "Pillow>=9.0.0"
pip install -e .
```

### Full Capabilities (Computer Vision & Windows COM)
For video analysis, image decomposition, and local Windows PowerPoint COM playback:
```bash
pip install -e ".[full]"
```

See [Offline Execution Guide](docs/offline-execution.md) and [Cross-Platform Compatibility Matrix](docs/cross-platform-matrix.md).

---

## Skills Architecture

| Skill | Primary responsibility |
| --- | --- |
| [pmg-engine](skills/pmg-engine/SKILL.md) | Universal runtime entrypoint, execution modes, 7-step agent workflow |
| [pmg-core](skills/pmg-core/SKILL.md) | 16:9 widescreen canvas, bright modern light color systems, typography scale |
| [pmg-morph](skills/pmg-morph/SKILL.md) | 60 FPS Morph compilation, `<p159:morph>`, persistent `!!` traveling entities |
| [pmg-choreography](skills/pmg-choreography/SKILL.md) | AI Motion Director, narrative pacing, animation sequencing and timing |
| [pmg-camera-parallax](skills/pmg-camera-parallax/SKILL.md) | Multi-plane parallax (BG 0.7x, Mid 1.0x, FG 1.4x), macro push-in zooms |
| [pmg-3d-cinematic](skills/pmg-3d-cinematic/SKILL.md) | Model placement/orientation, cutaways and annotations |
| [pmg-glassmorphism](skills/pmg-glassmorphism/SKILL.md) | Authentic frosted acrylic glass cards (`a:alpha`), hairline borders |
| [pmg-qa-audit](skills/pmg-qa-audit/SKILL.md) | Read-only Three-Level inspection (Structural, Visual, Playback) |
| [pmg-review-validation](skills/pmg-review-validation/SKILL.md) | Self-healing policy, 3-loop repair budget, delivery certification |
| [pmg-storyboard](skills/pmg-storyboard/SKILL.md) | Storyboard intelligence and narrative presentation planner |
| [pmg-visual-assets](skills/pmg-visual-assets/SKILL.md) | Authentic media research and retrieval (Wikimedia/NASA) |
| [pmg-image-decomposer](skills/pmg-image-decomposer/SKILL.md) | Decomposition of flat raster images into structured `LayerTree` hierarchies |

---

## Documentation

- [Architecture & Ownership](docs/architecture.md)
- [API & CLI Reference](docs/api-and-cli.md)
- [Cross-Platform Compatibility Matrix](docs/cross-platform-matrix.md)
- [Offline Execution Guide](docs/offline-execution.md)
- [Development Roadmap](docs/development-roadmap.md)
