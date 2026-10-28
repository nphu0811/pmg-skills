# PMG Skills — PowerPoint Motion Graphics (v2.1.0)

PMG creates 100% native, individually editable PowerPoint presentations through a universal, cross-platform Python runtime.
It combines 60 FPS Morph transitions, native Animation Pane timing trees, frosted glassmorphism layers, virtual camera parallax, and continuous traveling entities.

Current implementation: **v2.1.0 Universal Runtime & Cross-Platform Reliability**.

---

## What's New in v2.1.0: Universal Runtime

- **PowerPoint-Independent Native Compilation:** Native PowerPoint slide decks, Morph transitions, and Animation Pane effects can now be compiled in any environment running Python — **Microsoft PowerPoint is NOT required to be installed**.
- **Three Execution Environments:**
  * **Mode A: Web / Sandbox** (Claude Web, ChatGPT sandbox, restricted Linux CI without Office or internet). Native PPTX compilation + Level 1 structural validation.
  * **Mode B: Local Desktop** (Windows, macOS, Linux with Python and optional Microsoft PowerPoint COM/AppleScript).
  * **Mode C: Remote Worker** (configured & authenticated Windows validation worker for remote Slide Show playback testing).
- **Universal Capability Detection (`pmg doctor` / `pmg capabilities`):** Deterministic environment inspection that adapts automatically without fragile crashes.
- **Offline-First Dependency Strategy:** Prioritizes installed runtime -> cached/bundled wheels (`PMG_OFFLINE_CACHE`, `~/.pmg/cache`, `./wheels`) -> standalone binary -> network install -> graceful fallback.
- **Three-Level Validation Architecture:**
  * **Level 1 (Structural):** Mandatory ZIP, XML, OPC ContentTypes, relationships, unique shape IDs, `<p:timing>` trees, and Morph metadata.
  * **Level 2 (Visual):** Slide previews rendered via available backend (PowerPoint COM or LibreOffice). Reports `SKIPPED — No slide rendering backend available` if missing.
  * **Level 3 (Playback):** Native Microsoft PowerPoint Slide Show playback. Reports `NOT_TESTED — PowerPoint Desktop unavailable` if PowerPoint is absent. *Never fabricates Level 3 PASS without genuine PowerPoint Desktop execution.*

---

## Quickstart

### Python API

```python
from pmg import PMGRuntime

# Auto-detects environment and optimal validation backends
runtime = PMGRuntime.auto()

# Build and validate from JSON presentation spec, dictionary, or preset
result = runtime.build(
    project="presentation_spec.json",
    output="presentation.pptx",
    validate=True,
)

print(result.output_path)
print(result.validation_report.overall_status)
```

### Command-Line Interface (CLI)

```bash
# Check environment readiness and installed capabilities
pmg doctor
pmg capabilities

# Build presentation from JSON specification or preset
pmg build presentation_spec.json -o presentation.pptx
pmg build --preset space --slides 10 -o Space.pptx

# Validate presentation across Level 1, Level 2, and Level 3
pmg validate presentation.pptx
pmg validate presentation.pptx -o Validated.pptx --max-repairs 3
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
