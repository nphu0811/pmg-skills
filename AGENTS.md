# Guidance for AI Agents Working in PMG-Skills

Welcome, AI Agent! This repository contains the official **PowerPoint Motion Graphics (PMG) AI Skills**.
When working in this repository or generating presentations for users, follow these architectural directives, mental models, and invariants.

---

## 1. Core Mental Models

### A. The Camera Mindset
Never treat PowerPoint slides as disconnected flat pages. **Treat the presentation as a continuous 3D world traversed by a moving virtual camera.**

### B. The Continuous Traveling Entity ("Vật Thể Bay Xuyên Slide")
When crafting narrative presentations, orchestrate a persistent traveling element that glides smoothly across slides via Morph transitions. **The traveling entity must dynamically fit the user's specific topic—never force a single rigid object (like a crane):**
- **Space & Exploration:** Satellite, Rocket, Asteroid, Lunar Probe.
- **Technology & Music:** Vinyl Disc, Soundwave Node, Glowing Cursor, Hologram Cube.
- **Nature, Ecology, Geography:** Water Droplet, Lotus Leaf, Golden Bee, River Boat.
- **Finance & Crypto:** Gold Token, Growth Arrow Node, Hexagonal Shield.
- **Corporate & Strategy:** Milestone Pin, Directional Prism, Origami Fold.
The object guides the viewer's gaze from concept to concept, entering from the margin, interacting with cards, and exiting gracefully.

### C. The Native Transition Palette
Do not rely exclusively on Morph. Select transitions that evoke the right narrative psychology:
- **Morph:** Continuous spatial transformation, camera zooms, traveling hero flights, and 3D rotations.
- **Circle Wipe (`circle`):** Climax iris spotlight, focal conclusions.
- **Fade (`fade`):** Reflective mood shifts, moving from problem to solution.
- **Split Curtain (`split`):** Grand keynote title unveilings.

---

## 2. Skill Directory Layout

Agents should route user requests to the appropriate modular skill:

| Skill | Directory | When to Activate |
| :--- | :--- | :--- |
| **`pmg-engine`** | `skills/pmg-engine/` | **Canonical Engine-First Entrypoint:** Mandatory whenever PMG Skills is requested. Executes `pmg` engine, real PowerPoint playback recording, AI video review, 3-loop repair, and final `.pptx` delivery gate. |
| **`pmg-core`** | `skills/pmg-core/` | Base layout, 16:9 canvas geometry (`13.333" x 7.5"`), **bright modern light palettes as primary default**, readability and intentional layering. |
| **`pmg-morph`** | `skills/pmg-morph/` | PowerPoint Morph transitions, OpenXML `<p159:morph>`, `!!` identifier invariants, **thematic cross-slide traveling entities**. |
| **`pmg-choreography`** | `skills/pmg-choreography/` | High-level motion direction, selection, narrative sequencing and timing; invokes shared Morph, camera and native animation compilers. |
| **`pmg-glassmorphism`** | `skills/pmg-glassmorphism/` | Frosted acrylic glass cards (90%-94% alpha), subtle glowing hairline borders (`0.75pt - 1.25pt`), multi-tier depth. |
| **`pmg-camera-parallax`** | `skills/pmg-camera-parallax/` | Multi-plane parallax speeds (BG 0.7x, Mid 1.0x, FG 1.4x), macro push-in zooms ($100\% \to 400\%$), tracking pans. |
| **`pmg-3d-cinematic`** | `skills/pmg-3d-cinematic/` | Model preparation, placement/orientation, cutaways and annotations; reuses the camera system. Verify actual host model support. |
| **`pmg-qa-audit`** | `skills/pmg-qa-audit/` | Read-only structured inspection of package, layout, anchors, timing, media, compatibility and playback evidence; no repairs or certification. |
| **`pmg-visual-assets`** | `skills/pmg-visual-assets/` | Media research, authentic asset discovery (Wikimedia/NASA), licensing, provenance, and preparation; does not compile animations. |
| **`pmg-image-decomposer`** | `skills/pmg-image-decomposer/` | Semantic, geometric, and depth decomposition of flat images into structured `LayerTree` hierarchies with vector/raster editability and reassembly verification. |
| **`pmg-review-validation`** | `skills/pmg-review-validation/` | Consumes shared audit reports, chooses repairs and certification policy; orchestrator owns one maximum-three-attempt loop and final authority. |

---

## 3. Strict Architectural Invariants

### A. Bright Modern Default vs. Purposeful Dark Mode
- **Default to Bright/Light Themes:** Use `LIGHT_ACADEMIC_CYAN`, `LIGHT_MODERN_STUDIO`, or `LIGHT_ECO_EMERALD` with pure white / ice white backgrounds and high-contrast deep slate text (`#0F172A`). Avoid muddy gray or unintended gloomy dark slides.
- **Selective Dark Mode:** Reserve Obsidian/Cyber dark themes strictly for aerospace, late-night tech keynotes, or explicit user requests.

### B. The `!!` Naming Invariant
- Any element morphing across slides MUST share identical names prefixed with `!!` in PowerPoint's shape tree (e.g. `!!Traveling_Crane`, `!!Hero_Card`, `!!Step_Tracker`).

### C. Continuous Coordinate Staging
- For traveling objects entering or exiting the canvas:
  - Left off-screen: $X \le -2.0 \text{ inches}$
  - Right off-screen: $X \ge 13.5 \text{ inches}$

### D. Zero Static Image Clutter in `examples/`
- Deliver demonstration artifacts as **high-definition 60 FPS MP4 videos** in `examples/`, capturing actual motion and transitions in real time.

### E. Universal Runtime, Three-Level Validation & Zero-Repair Delivery Gate
- **Universal Runtime Execution:** Execute the canonical `pmg` runtime (`pmg.runtime.PMGRuntime`, `pmg.orchestrator.PMGOrchestrator`, or `python -m pmg`). Adapts automatically across Mode A (Web/Sandbox), Mode B (Local Desktop), and Mode C (Remote Worker).
- **PowerPoint-Independent Compilation:** Creating native PowerPoint presentations and animations does NOT require Microsoft PowerPoint to be installed. Slides, editable shapes, text frames, Morph transitions, and intra-slide Animation Pane timing trees compile via pure Python OpenXML synthesis.
- **Structured Three-Level Validation:**
  * **Level 1 (Structural):** Mandatory everywhere. Validates ZIP integrity, XML well-formedness, OPC ContentTypes, relationships, unique shape IDs, animation target existence, and `<p:timing>` trees.
  * **Level 2 (Visual):** Executed when a slide renderer (PowerPoint COM or LibreOffice) is available. If unavailable, reports `SKIPPED - No slide rendering backend available`. Static rendering does NOT prove animation correctness.
  * **Level 3 (Native Playback):** Executed when Microsoft PowerPoint automation is genuinely available (local COM or remote Windows worker). If unavailable, reports `NOT_TESTED - PowerPoint Desktop unavailable`. *Never report Level 3 PASS without actual Microsoft PowerPoint Desktop execution.*
- **OPC ContentType Declaration:** When embedding manifest JSON in PPTX (`ppt/pmg_execution_manifest.json`), always declare a valid ContentType in `[Content_Types].xml` (`<Default Extension="json" ContentType="application/json"/>`) and write a companion manifest (`<pptx_path>.manifest.json`) outside the package. Unmapped parts that trigger the PowerPoint Repair dialog are strictly prohibited.
- **Zero Modifications Post-Validation:** NEVER modify a `.pptx` file after marking validation `PASS` without re-running full verification. The deliverable is frozen and verified in its final byte state.
- **Mandatory Clean COM Open Verification:** The final packaged deliverable must be opened in Microsoft PowerPoint Desktop COM with `DisplayAlerts = ppAlertsAll` (`2`; `1` suppresses alerts), with `OpenAndRepair=False` to ensure zero "Repair presentation" prompts or corruption warnings.
- **One Repair Budget:** Only PMGOrchestrator orchestrates up to three technical repair attempts using pmg.self_healing operations and shared audit reports. QA never repairs. Repackage, re-render and re-record after every repair.
- **Fallback Within Budget:** Use compatible native effect substitution inside attempt three; do not start extra repair loops or reset the budget. Unresolved defects block delivery.
- **Final Delivery Policy:** Deliver **ONLY** the final validated `.pptx` file to the user. Keep `internal_qa/<run_id>/` videos, keyframes, and JSON reports internal unless explicitly requested by the user.

---

## 4. Reusable Script & Package Engines

- `pmg/` (`pmg.runtime`, `pmg.capabilities`, `pmg.compiler`, `pmg.validation`, `pmg.dependencies`, `pmg.orchestrator`, `pmg.engine`, `pmg.validator`, `pmg.video_pipeline`, `pmg.visual_reviewer`, `pmg.self_healing`, `pmg.cli`): Canonical Python package and CLI (`pmg doctor`, `pmg capabilities`, `pmg build`, `pmg validate`, `pmg render`, `pmg record`, `pmg review`, `pmg inspect`, `pmg reconstruct`).
- `scripts/motion_graph_kit.py`: Core presentation builder with light themes, frosted cards, multi-transition suite, traveling entities, and video export.
- `scripts/motion_validator.py`: Backward compatible adapter to shared read-only QA and canonical orchestrator repair/certification; no separate validator or loop.
- `scripts/slide_inspector.py`: Read-only adapter to pmg.validator.StaticValidator structured reports.
- `scripts/pptx_morph_injector.py`: Standalone CLI transition injector.


## 5. Distinct ownership and continued roadmap

Follow [docs/architecture.md](docs/architecture.md) for primary responsibility and
dependencies. Morph owns technical cross-slide compilation; choreography decides
what/why/when and invokes Morph. Camera owns transforms/paths/depth/continuity;
3D owns models/cutaways/orientation/annotations and reuses camera. QA detects
defects read-only; review chooses repairs; orchestrator remains final authority.
All use pmg.models audit/defect/review records. Do not create competing compilers,
segmentation systems, validators, benchmark suites or repair loops.

Allow intentional overlap, masks, translucent overlays and text over imagery.
Protect readable regions and report accidental collisions; no global zero-overlap
rule. Keep object identity, hierarchy, geometry/units, mask references, z-order,
layer IDs, pivots and animation targets consistent in shared interfaces.

Runtime installation is required after copying skills. Verify `python -m pmg
--version` and install the same published commit as the skill when missing.
Explicitly load dependencies on hosts without automatic discovery. Never assume
SKILL.md installation includes the Python package.

Continue [docs/development-roadmap.md](docs/development-roadmap.md) without
restarting existing work: v1.5.2 → v1.6.0 → v1.7.0 → v1.8.0 → v2.0.1. Only begin
the next implementation after the prior release passes publish/install/retest
and real playback gates. v2.0.1 upgrades existing choreography/storyboard into
Creative Motion Intelligence; it adds no competing director/graph/compiler skills.
Do not claim future roadmap interfaces as working functionality.

The responsibility audit baseline was v1.5.1; reliability work is v1.5.2. Publication
requires the recorded release gates, including a published-commit reinstall. Packaging
precedes full validation. Embedded provenance is NOT_VERIFIED and points to the
external companion containing final status/hash. Never modify the certified PPTX.
