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
Camera Moves → 60 FPS Morph → Glassmorphism → Multi-Plane Parallax → 3D Cutaways → Visual-First

# PMG AI Skills (PowerPoint Motion Graphics)

Official modular AI skills for **PowerPoint Motion Graphics (PMG)**. They teach AI agents to design and engineer broadcast-grade `.pptx` presentations using hardware-accelerated 60 FPS Morph transitions, virtual multi-plane camera choreography, frosted glassmorphism, 3D cutaway models, and ultra-minimal text architectures. Built according to the [Agent Skills](https://agentskills.io) standard; works with the [skills CLI](https://github.com/vercel-labs/skills) across Claude Code, Cursor, Codex, Windsurf, Copilot, Google Antigravity, and 40+ AI agents.

> **PMG is 100% open and native**: Generated presentations run directly in standard Microsoft PowerPoint (Desktop, 365, Mac, Windows) without requiring proprietary plugins, browser renderers, or subscription tools.

![Claude Skill](https://img.shields.io/badge/Claude-Skill-purple)
![Cursor Compatible](https://img.shields.io/badge/Cursor-Compatible-blue)
![Copilot Ready](https://img.shields.io/badge/Copilot-Ready-green)
![Antigravity Ready](https://img.shields.io/badge/Antigravity-Ready-orange)
![PowerPoint Morph 60FPS](https://img.shields.io/badge/PowerPoint-Morph_60FPS-orange.svg)
![Python 3.8+](https://img.shields.io/badge/Python-3.8+-brightgreen.svg)
![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)

**For AI agents:** When the user asks for a **PowerPoint presentation**, **motion graphics deck**, **animated slides**, or **product/technical explainer**, activate **PMG Skills** for smooth 60 FPS Morph transitions, virtual 3D camera choreography, authentic frosted glassmorphism, and visual-first impact.

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

## Skills

| Skill | Description |
|---|---|
| **`pmg-core`** | Base canvas: 16:9 widescreen geometry (`13.333" x 7.5"`), **bright modern light palettes as primary default**, typographic scales, zero-overlap bounding grids. |
| **`pmg-morph`** | Hardware-accelerated 60 FPS PowerPoint Morph engine: OpenXML `<p159:morph>`, `!!` identifier naming invariants, **thematic cross-slide traveling entities (satellites, drones, vinyls, droplets)**, aspect-ratio locks. |
| **`pmg-glassmorphism`** | Frosted acrylic glass scrims (90%-94% alpha), glowing hairline borders (`0.75pt - 1.25pt`), multi-tier depth. |
| **`pmg-camera-parallax`** | Virtual multi-plane camera choreography, establishing wide shots, macro push-in zooms ($100\% \to 400\%$), 3-plane speed ratios (BG 0.7x, Mid 1.0x, FG 1.4x). |
| **`pmg-3d-cinematic`** | Visual-First explainer engine for science & industry: Native `.glb` 3D models (`<am3d:model3d>`), Extreme Macro Dolly Zoom ($100\% \to 400\%$), ultra-minimal text ($\le 25$ words/slide), leader lines, target dots. |
| **`pmg-choreography`** | **Native PowerPoint Transitions suite (Morph, Circle Wipe, Split Curtain, Fade)**, **Spotify Radial Morph App Shell unlock**, and **3D Orbiting Carousels**. |
| **`pmg-qa-audit`** | Automated QA audit engine: slide inspection, transition coverage verification, traveling entity tracking, **60 FPS MP4 video export via PowerPoint COM**. |
| **`pmg-review-validation`** | **Automated review, validation & self-healing engine**: Detects missing transitions, executes **3-loop auto-repair cycle**, triggers **fallback skill substitution**, and certifies delivery completion. |

---

## Quick Reference (For AI Agents)

Canonical PMG pattern to suggest or generate:

```python
from scripts.motion_graph_kit import LightThemeCraftDeck, CinematicColors

# 1. Initialize 16:9 Widescreen Deck with Modern Light Palette
deck = LightThemeCraftDeck(title="Autonomous AI Core")

# 2. Slide 1: Wide Establishing Shot (Camera Zoom 100%)
s1 = deck.add_craft_slide(
    title="NEURAL FABRIC ARCHITECTURE",
    category="AI COMPUTE PLATFORM",
    badge="STAGE 01",
    subtitle="Global overview of high-throughput heterogeneous compute clusters."
)

# 3. Slide 2: Macro Push-In Zoom with Frosted Glassmorphism
s2 = deck.add_craft_slide(
    title="TENSOR ACCELERATOR ENGINE",
    category="DEEP LEARNING HARDWARE",
    badge="STAGE 02",
    subtitle="Sub-millisecond matrix execution running across distributed pipeline nodes."
)

# 4. Save and auto-inject 60 FPS OpenXML Morph transitions
deck.save("AI_Core_Cinematic.pptx")
```

---

## Structure

```
pmg-skills/
  README.md          # Official overview and agent skill guide
  AGENTS.md          # Core mental models and strict invariants for AI agents
  CLAUDE.md          # Guidance pointer for Claude Code agents
  GEMINI.md          # Guidance pointer for Antigravity & Gemini agents
  LICENSE            # MIT License
  .github/
    copilot-instructions.md   # Repo-wide instructions for GitHub Copilot
    instructions/             # Path-specific Copilot instructions
      morph.instructions.md
      glassmorphism.instructions.md
  .claude-plugin/    # Claude Code plugin config (plugin.json, marketplace.json)
  .cursor-plugin/    # Cursor plugin config (plugin.json, marketplace.json)
  scripts/           # Core execution engines
    motion_graph_kit.py # Editorial craft deck builder with glassmorphism & themes
    motion_validator.py # Review, validation, 3-loop auto-repair, & fallback substitution engine
    forge_lib.py     # 3D cinematic camera matrix, HUD leader lines, morph injector
    asset_engine.py  # Procedural 2K studio grid background generation
    slide_inspector.py # Automated QA audit (100% morph, <= 25 words/slide)
    pptx_morph_injector.py # Standalone OpenXML Morph injector CLI
  skills/
    llms.txt         # Master skill index for agents (names, triggers, paths)
    pmg-core/        SKILL.md
    pmg-morph/       SKILL.md
    pmg-glassmorphism/ SKILL.md
    pmg-camera-parallax/ SKILL.md
    pmg-3d-cinematic/ SKILL.md
    pmg-choreography/ SKILL.md
    pmg-qa-audit/    SKILL.md
    pmg-review-validation/ SKILL.md
  examples/          # High-definition 60 FPS MP4 video demos and visual showcases
  demo/              # Live presentation screen recordings and animated previews
    pmg_demo_preview.gif
    pmg_presentation_demo.mp4
  assets/            # Background patterns, vector icons, and textures
```

---

## GitHub Copilot

Copilot doesn’t load Cursor/Claude skill files. To get PMG guidance in a repo, copy or adapt the [`.github/copilot-instructions.md`](.github/copilot-instructions.md) (and optional [`.github/instructions/`](.github/instructions/) path-specific files) into that repo. See [GitHub Copilot customization](https://docs.github.com/en/copilot/concepts/response-customization).

---

## Risk Level

**LOW** — Pure OpenXML and Python presentation engineering without binary native code or elevated OS privileges.

---

## License

MIT — See [LICENSE](LICENSE).
