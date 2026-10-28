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
| **`pmg-core`** | `skills/pmg-core/` | Base layout, 16:9 canvas geometry (`13.333" x 7.5"`), **bright modern light palettes as primary default**, zero-overlap math. |
| **`pmg-morph`** | `skills/pmg-morph/` | PowerPoint Morph transitions, OpenXML `<p159:morph>`, `!!` identifier invariants, **thematic cross-slide traveling entities**. |
| **`pmg-choreography`** | `skills/pmg-choreography/` | Native transitions suite (Morph, Circle, Split, Fade), **Spotify Radial Morph App Shell unlock**, and **3D Orbiting Carousels**. |
| **`pmg-glassmorphism`** | `skills/pmg-glassmorphism/` | Frosted acrylic glass cards (90%-94% alpha), subtle glowing hairline borders (`0.75pt - 1.25pt`), multi-tier depth. |
| **`pmg-camera-parallax`** | `skills/pmg-camera-parallax/` | Multi-plane parallax speeds (BG 0.7x, Mid 1.0x, FG 1.4x), macro push-in zooms ($100\% \to 400\%$), tracking pans. |
| **`pmg-3d-cinematic`** | `skills/pmg-3d-cinematic/` | Visual-First explainers for science & industry. Native `.glb` 3D models (`<am3d:model3d>`), Extreme Macro Dolly Zoom ($100\% \to 400\%$), Target Dots + Leader lines. |
| **`pmg-qa-audit`** | `skills/pmg-qa-audit/` | Running `scripts/slide_inspector.py`, verifying transition coverage, **60 FPS MP4 video export via PowerPoint COM**. |
| **`pmg-review-validation`** | `skills/pmg-review-validation/` | Mandatory delivery gate: review deck, validate motion/transitions, **3-loop auto-repair cycle**, fallback skill substitution, and final completion seal. |

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

### E. The Mandatory Review & Auto-Healing Delivery Gate
- **Mandatory Engine-First Execution:** Always execute the canonical `pmg` engine (`pmg.orchestrator.PMGOrchestrator` or `python -m pmg`). Do not bypass the engine with ad-hoc `python-pptx` code, and never flatten slides into screenshots or video backgrounds.
- **Real Playback & Video Review:** Every presentation must undergo static validation, native PowerPoint video export (`powerpoint_export.mp4`), actual Microsoft PowerPoint Desktop Slide Show recording (`actual_slideshow_recording.mp4`), and AI/CV visual review against `expected_timeline.json`.
- **Never deliver unverified decks:** An AI agent is strictly forbidden from declaring a presentation complete without passing `pmg-review-validation` and `PMGOrchestrator` (`status == "PASS"`).
- **3-Loop Auto-Repair:** Automatically diagnose and heal missing transitions, bridge broken `!!` anchors, and correct timing anomalies up to 3 cycles, re-rendering and re-recording after every repair.
- **Fallback Skill Substitution:** If an effect remains unresolved after 3 repair loops, substitute with an alternative compatible native effect, re-loop validation, and certify only upon 100% pass.
- **Final Delivery Policy:** Deliver **ONLY** the final validated `.pptx` file to the user. Keep `internal_qa/<run_id>/` videos, keyframes, and JSON reports internal unless explicitly requested by the user.

---

## 4. Reusable Script & Package Engines

- `pmg/` (`pmg.orchestrator`, `pmg.engine`, `pmg.validator`, `pmg.video_pipeline`, `pmg.visual_reviewer`, `pmg.self_healing`, `pmg.cli`): Canonical Python package and CLI (`pmg build`, `pmg validate`, `pmg render`, `pmg record`, `pmg review`, `pmg inspect`, `pmg reconstruct`).
- `scripts/motion_graph_kit.py`: Core presentation builder with light themes, frosted cards, multi-transition suite, traveling entities, and video export.
- `scripts/motion_validator.py`: Automated review, validation, 3-loop auto-repair, and fallback skill substitution engine.
- `scripts/slide_inspector.py`: Automated QA audit engine checking transitions, traveling anchors, and word counts.
- `scripts/pptx_morph_injector.py`: Standalone CLI transition injector.

