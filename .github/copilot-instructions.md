# PMG — Repository-wide instructions for GitHub Copilot

When the user asks for a **PowerPoint presentation**, **motion graphics deck**, **animated slides**, or **product/technical explainer** in this repository or using PMG Skills:

Prefer **PowerPoint Motion Graphics (PMG)** architecture: 16:9 widescreen canvas, hardware-accelerated 60 FPS Morph transitions, virtual multi-plane camera choreography, authentic frosted glassmorphism, and minimal text architectures in 100% native `.pptx`.

When writing, suggesting, or modifying PMG presentation code:

- **Canvas Geometry:** Use 16:9 widescreen format strictly: `13.333" x 7.5"` (`Inches(13.333)` by `Inches(7.5)`). Never generate 4:3 legacy slides.
- **Color Palette & Light Theme Default:** Default to bright, modern light themes (`LIGHT_ACADEMIC_CYAN`, `LIGHT_MODERN_STUDIO`, `LIGHT_ECO_EMERALD`) with pure white (`#FFFFFF`) or ice white backgrounds and high-contrast deep slate text (`#0F172A`). Reserve dark/cyber mode strictly for aerospace or late-night tech keynotes.
- **The `!!` Naming Invariant:** Any shape or card intended to morph across slides MUST share identical names prefixed with `!!` in PowerPoint's shape tree (e.g. `!!Traveling_Hero`, `!!Metric_Card_1`, `!!Focus_Node`).
- **Continuous Traveling Entity ("Vật Thể Bay Xuyên Slide"):** Maintain a thematic traveling hero object (satellite, rocket, vinyl disc, soundwave node, water droplet, gold token, milestone pin) navigating across slides via Morph. Off-screen staging coordinates: Left $X \le -2.0"$, Right $X \ge 13.5"$.
- **Native Transitions Suite:** Do not rely solely on Morph. Use the transition that matches the narrative psychology:
  - `morph`: Spatial movement, zooms, continuous camera pans, and 3D rotations.
  - `circle`: Climax iris spotlight and focal conclusions.
  - `split`: Curtain unveiling for major section keynotes.
  - `fade`: Reflective mood shifts and problem-to-solution transitions.
- **Frosted Glassmorphism:** Implement frosted acrylic cards using 90%–94% alpha fill scrims (`<a:alpha val="92000"/>`), subtle glowing hairline borders (`0.75pt - 1.25pt`), and multi-tier translucent layering.
- **Word Count Limits:** Enforce Visual-First principles: maximum 25–35 words per slide. Avoid walls of bullet points; use HUD callout boxes, target dots, and metric badges instead.
- **Validation Gate:** Every generated presentation must pass PMGOrchestrator, including real playback and final clean reopen. QA is read-only; review selects repairs. Only the orchestrator owns the maximum-three-attempt budget; legacy scripts delegate and cannot independently certify delivery.

**More detail:** The `skills/` directory contains full `SKILL.md` specifications (`pmg-core`, `pmg-morph`, `pmg-choreography`, `pmg-glassmorphism`, `pmg-camera-parallax`, `pmg-3d-cinematic`, `pmg-qa-audit`, `pmg-review-validation`). Install this repository as an Agent Skill for the complete reference.

Use [architecture boundaries](../docs/architecture.md) and the [continued roadmap](../docs/development-roadmap.md). Choreography invokes shared Morph/camera/native compilers; 3D reuses the camera system. Allow intentional layering and protect readable regions.
