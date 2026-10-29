```text
  ██████╗ ███╗   ███╗ ██████╗ 
  ██╔══██╗████╗ ████║██╔════╝ 
  ██████╔╝██╔████╔██║██║  ███╗
  ██╔═══╝ ██║╚██╔╝██║██║   ██║
  ██║     ██║ ╚═╝ ██║╚██████╔╝
  ╚═╝     ╚═╝     ╚═╝ ╚═════╝ 

        ███████╗██╗  ██╗██╗██╗     ██╗     ███████╗
        ██╔════╝██║ ██╔╝██║██║     ██║     ██╔════╝
        ███████╗█████╔╝ ██║██║     ██║     ███████╗
        ╚════██║██╔═██╗ ██║██║     ██║     ╚════██║
        ███████║██║  ██╗██║███████╗███████╗███████║
        ╚══════╝╚═╝  ╚═╝╚═╝╚══════╝╚══════╝╚══════╝

  ──●────●────●────●────●────●──
   AI Skills for Claude • Cursor • Copilot • Antigravity
```
Camera Moves → 60 FPS Morph → Glassmorphism → Multi-Plane Parallax → Object Motion Intelligence → Visual-First

# PMG AI Skills — PowerPoint Motion Graphics (v2.5.0)

Official modular AI skills for **PowerPoint Motion Graphics (PMG)**. They teach AI agents to design and engineer broadcast-grade `.pptx` presentations using hardware-accelerated 60 FPS Morph transitions, virtual multi-plane camera choreography, frosted glassmorphism, 3D cutaway models, ultra-minimal text architectures, and **Universal Object-Specific Motion Intelligence**. Built according to the [Agent Skills](https://agentskills.io) standard; works with the [skills CLI](https://github.com/vercel-labs/skills) across Claude Code, Cursor, Codex, Windsurf, Copilot, Google Antigravity, and 40+ AI agents.

> **PMG is 100% open and native**: Generated presentations run directly in standard Microsoft PowerPoint (Desktop, 365, Mac, Windows) without requiring proprietary plugins, browser renderers, or subscription tools. Native PPTX slides and animation timing trees compile purely in Python without requiring Microsoft PowerPoint to be installed.

![Claude Skill](https://img.shields.io/badge/Claude-Skill-purple)
![Cursor Compatible](https://img.shields.io/badge/Cursor-Compatible-blue)
![Copilot Ready](https://img.shields.io/badge/Copilot-Ready-green)
![Antigravity Ready](https://img.shields.io/badge/Antigravity-Ready-orange)
![PowerPoint Morph 60FPS](https://img.shields.io/badge/PowerPoint-Morph_60FPS-orange.svg)
![Version 2.5.0](https://img.shields.io/badge/Version-2.5.0-blue.svg)
![Python 3.8+](https://img.shields.io/badge/Python-3.8+-brightgreen.svg)
![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)

**For AI agents:** When the user asks for a **PowerPoint presentation**, **motion graphics deck**, **animated slides**, or **product/technical explainer**, activate **PMG Skills** for smooth 60 FPS Morph transitions, virtual 3D camera choreography, authentic frosted glassmorphism, and semantic object-specific motion intelligence.

---

## 🎬 Live Presentation Demo

Broadcast-grade motion graphics running directly in **100% native Microsoft PowerPoint** (no proprietary plugins, web renderers, or external video layers):

<p align="center">
  <img src="demo/pmg_demo_preview.gif" alt="PowerPoint Motion Graphics Live Demo" width="100%" />
</p>

<p align="center">
  ▶️ <b><a href="demo/pmg_presentation_demo.mp4">Watch Full Slide Show Recording (50s MP4)</a></b> — Showcases 3D Orbiting Carousels, OpenXML 60 FPS Morph transitions, thematic traveling nodes, and authentic frosted glassmorphism.
</p>

---

## Installing

### npx skills (recommended)

Works with Cursor, Claude Code, Codex, Windsurf, Copilot, Google Antigravity, and [40+ agents](https://github.com/vercel-labs/skills#supported-agents):

```bash
npx skills add https://github.com/nphu0811/pmg-skills
```

The CLI auto-detects the installed agent. To target one explicitly (e.g. Antigravity):

```bash
npx skills add https://github.com/nphu0811/pmg-skills --agent antigravity
```

### Claude Code

In Claude Code, use the skill/plugin marketplace: `/plugin marketplace add nphu0811/pmg-skills`. See [Agent Skills docs](https://docs.anthropic.com/en/docs/agents-and-tools/agent-skills/overview).

### Cursor

**Settings → Rules → Add Rule → Remote Rule (Github)** and use `nphu0811/pmg-skills`. Or install via `npx skills add` above.

### Clone / copy

Copy the `skills/` folder into your agent’s skill directory:

| Agent | Skill Directory | Docs |
|---|---|---|
| Claude Code | `~/.claude/skills/` | [docs](https://docs.anthropic.com/en/docs/agents-and-tools/claude-code/skills) |
| Cursor | `~/.cursor/skills/` | [docs](https://docs.cursor.com/context/rules) |
| Google Antigravity | `~/.gemini/antigravity/skills/` (global) or `.agents/skills/` (workspace) | [docs](https://codelabs.developers.google.com/getting-started-with-antigravity-skills) |
| OpenAI Codex | `~/.codex/skills/` | [docs](https://developers.openai.com/codex/skills/) |
| OpenCode | `~/.config/opencode/skills/` | [docs](https://opencode.ai/docs/skills/) |

---

## Skills Ecosystem

| Skill | Category | Description |
|---|---|---|
| **`pmg-engine`** | **Runtime Core** | **Universal Engine-First Entrypoint (`pmg` CLI & `PMGRuntime`)**: Cross-platform runtime across Web Sandboxes (Mode A), Local Desktops (Mode B), and Remote Workers (Mode C). Generates 100% native editable PPTX without requiring Microsoft PowerPoint. |
| **`pmg-core`** | **Canvas** | Base canvas: 16:9 widescreen geometry (`13.333" x 7.5"`), **bright modern light palettes as primary default**, typographic scales, anti-overlap layout grids. |
| **`pmg-morph`** | **Transitions** | Hardware-accelerated 60 FPS PowerPoint Morph engine: OpenXML `<p159:morph>`, `!!` identifier naming invariants, **thematic cross-slide traveling entities (satellites, drones, vinyls, droplets)**, aspect-ratio locks. |
| **`pmg-glassmorphism`** | **Materials** | Frosted acrylic glass scrims (90%-94% alpha), glowing hairline borders (`0.75pt - 1.25pt`), multi-tier depth. |
| **`pmg-camera-parallax`** | **Cinematics** | Virtual multi-plane camera choreography, establishing wide shots, macro push-in zooms ($100\% \to 400\%$), 3-plane speed ratios (BG 0.7x, Mid 1.0x, FG 1.4x). |
| **`pmg-3d-cinematic`** | **3D Models** | Visual-First explainer engine for science & industry: Native `.glb` 3D models (`<am3d:model3d>`), Extreme Macro Dolly Zoom ($100\% \to 400\%$), ultra-minimal text ($\le 25$ words/slide), leader lines. |
| **`pmg-choreography`** | **Choreography** | AI Motion Director coordinating narrative sequencing, transition selection (Morph, Circle Wipe, Split Curtain, Fade), and cross-object choreography. |
| **`pmg-qa-audit`** | **Quality Assurance** | Read-only Three-Level inspection: Level 1 Structural (ZIP, XML, targets), Level 2 Visual (previews), and Level 3 Native Playback (real PowerPoint COM slideshow). |
| **`pmg-review-validation`**| **Certification** | Consumes audit defects, executes transactional self-healing within a bounded 3-loop budget, and issues tamper-evident cryptographic validation certificates. |
| **`pmg-storyboard`** | **Planning** | Narrative presentation planner: scene sequencing, camera intentions, asset requirements, and visual pacing. |
| **`pmg-visual-assets`** | **Media** | Intelligent media research and retrieval from verified public repositories (Wikimedia Commons, NASA). |
| **`pmg-image-decomposer`**| **Computer Vision** | Decomposes flat raster images into structured `LayerTree` hierarchies with vector/raster editability and reassembly verification. |
| **`pmg-chart-motion`** | **Object Motion** | Quantitative data charts and numerical value animation: directional bar growth, line draw, doughnut slice reveal, synchronized value counters, and peak metric emphasis. |
| **`pmg-kinetic-typography`**| **Object Motion** | Typographic motion intelligence: authoritative headline unveiling, reading-order word staggers, keyword scale pops (1.14x), typewriter cadences, and kinetic headline morphs. |
| **`pmg-image-motion`** | **Object Motion** | Photographic composition motion: directional unmasking along leading lines, 3-plane layered parallax, clamped macro dolly zooms ($\le 125\%$), and before/after wipes. |
| **`pmg-diagram-motion`** | **Object Motion** | Process flowcharts and architecture diagrams: strict causal sequence (Node A $\to$ Connector $\to$ Node B $\to$ Connector $\to$ Node C), and data packet throughput. |
| **`pmg-map-motion`** | **Object Motion** | Geospatial maps: territory highlights, location marker pin drops, hub radar ripples, and progressive route vector drawings. |
| **`pmg-infographic-motion`**| **Object Motion** | Timelines and progress indicators: chronological timeline spines, milestone pin emergence, horizontal progress bar growth, and circular gauges. |

---

## Quick Reference (For AI Agents & Developers)

### 1. Canonical PMG CLI (`pmg`)

```bash
# Environment inspection and capability detection
pmg doctor
pmg capabilities

# Build presentation from JSON specification or preset with safe transactional healing
pmg build presentation_spec.json -o deck.pptx --repair safe
pmg build --preset space --slides 10 -o Space_Mission.pptx

# Three-Level validation (Structural, Visual, Playback)
pmg validate deck.pptx

# Distributed QA validation (local COM or remote worker)
pmg qa deck.pptx
pmg qa deck.pptx --remote --timeout 60

# Transactional self-healing repair of damaged PPTX
pmg repair deck.pptx --safe -o deck_repaired.pptx
```

### 2. Universal Python API (`pmg`)

```python
from pmg import PMGRuntime, MotionDirector, ChartMotionPlanner, KineticTypographyPlanner

# 1. Universal Runtime execution with safe self-healing
runtime = PMGRuntime.auto()
result = runtime.build(
    project="presentation_spec.json",
    output="Production_Deck.pptx",
    validate=True,
    repair="safe",
)
print(f"Artifact SHA-256: {result.artifact_sha256}")
print(f"Validation Status: {result.validation_report.overall_status}")

# 2. Universal Object-Specific Motion Planning
director = MotionDirector()
chart_planner = ChartMotionPlanner()

# Quantitative bar growth with value tracking
chart_intents = chart_planner.plan_motion({
    "chart_type": "bar",
    "axis_id": "chart_axis",
    "items": [
        {"id": "bar_q1", "value": "$1.2M", "label_id": "lbl_q1"},
        {"id": "bar_q2", "value": "$1.8M", "label_id": "lbl_q2"},
        {"id": "bar_q3", "value": "$2.9M", "label_id": "lbl_q3", "is_peak": True},
    ]
})

for intent in chart_intents:
    print(f"[{intent.target_id}] -> {intent.primitive.value} ({intent.semantic_purpose})")
```

---

## Structure

```
pmg-skills/
  README.md          # Official overview and agent skill guide (v2.5.0)
  AGENTS.md          # Core mental models and strict invariants for AI agents
  CLAUDE.md          # Guidance pointer for Claude Code agents
  GEMINI.md          # Guidance pointer for Antigravity & Gemini agents
  LICENSE            # MIT License
  pyproject.toml     # Universal package configuration
  pmg/               # Canonical Python engine package
    runtime.py       # Universal runtime (Mode A, Mode B, Mode C)
    capabilities.py  # Environment capability detection
    animation_ir.py  # Unified Animation Intermediate Representation
    compiler.py      # Native OpenXML PowerPoint compiler
    motion_director.py # Motion Director & cross-object choreography
    semantic_classifier.py # Semantic object classification
    motion_decision.py     # Explainable motion decision engine
    motion_quality.py      # Motion quality & reduced motion validation
    self_healing.py  # Transactional repair & immutable certificates
    motion_planners/ # Specialized object motion planners
      chart_motion.py
      kinetic_typography.py
      image_motion.py
      diagram_motion.py
      map_motion.py
      infographic_motion.py
    worker/          # Distributed PowerPoint validation worker
  scripts/           # Core presentation scripts & adapters
    motion_graph_kit.py
    motion_validator.py
  skills/            # Agent Skills directory
    llms.txt         # Master skill index
    pmg-engine/      SKILL.md
    pmg-core/        SKILL.md
    pmg-morph/       SKILL.md
    pmg-glassmorphism/ SKILL.md
    pmg-camera-parallax/ SKILL.md
    pmg-3d-cinematic/ SKILL.md
    pmg-choreography/ SKILL.md
    pmg-qa-audit/    SKILL.md
    pmg-review-validation/ SKILL.md
    pmg-storyboard/  SKILL.md
    pmg-visual-assets/ SKILL.md
    pmg-image-decomposer/ SKILL.md
    pmg-chart-motion/ SKILL.md
    pmg-kinetic-typography/ SKILL.md
    pmg-image-motion/ SKILL.md
    pmg-diagram-motion/ SKILL.md
    pmg-map-motion/  SKILL.md
    pmg-infographic-motion/ SKILL.md
  demo/              # Presentation recordings and animated previews
    pmg_demo_preview.gif
    pmg_presentation_demo.mp4
```

---

## GitHub Copilot

Copilot doesn’t load Cursor/Claude skill files. To get PMG guidance in a repo, copy or adapt the [`.github/copilot-instructions.md`](.github/copilot-instructions.md) into that repo. See [GitHub Copilot customization](https://docs.github.com/en/copilot/concepts/response-customization).

---

## Risk Level

**LOW** — Pure OpenXML and Python presentation engineering without binary native code or elevated OS privileges.

---

## License

MIT — See [LICENSE](LICENSE).
