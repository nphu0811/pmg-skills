---
name: pmg-engine
description: "Universal PowerPoint Motion Graphics (PMG) v2.1.0 engine. Cross-platform runtime for AI web sandboxes, local AI agents, and desktop environments. Generates 100% native editable PPTX with 60 FPS Morph and intra-slide Animation Pane timing trees without requiring Microsoft PowerPoint to be installed."
---

# PMG Universal Runtime & Canonical Entrypoint (v2.1.0)

PMG v2.1.0 provides a universal, offline-first, capability-aware PowerPoint generation runtime.
The compiler produces 100% native, editable PowerPoint presentations across:
- **Mode A: Web / Sandbox** (Claude Web, ChatGPT sandbox, restricted Linux environments without Microsoft PowerPoint or internet).
- **Mode B: Local Desktop** (Windows, macOS, Linux with local Python and optional PowerPoint COM/AppleScript).
- **Mode C: Remote Worker** (configured & authenticated Windows validation worker).

**Critical Invariant: Creating native PowerPoint animations does NOT require Microsoft PowerPoint to be installed.**
Microsoft PowerPoint is an optional Level 3 validation and playback backend, never a compilation prerequisite.

---

## Universal Agent Execution Workflow

All AI agents (Claude Web, ChatGPT, Claude Code, Codex, local agents) follow this 7-step execution workflow:

1. **Inspect Environment Capabilities:**
   Run `pmg doctor` or query in Python:
   ```python
   from pmg import PMGRuntime
   runtime = PMGRuntime.auto()
   print(runtime.doctor())
   ```
2. **Determine PMG Accessibility:**
   Verify if Python and core dependencies (`python-pptx`, `lxml`, `Pillow`) are present.
   - If available, use the active runtime directly.
   - If missing in an offline environment, inspect cached wheels via `PMG_OFFLINE_CACHE` or local `./wheels`.
   - Never abandon PMG merely because `pip install` or network access fails if core dependencies are already present!
   - If the environment strictly prohibits code execution, honestly report that limitation rather than pretending to execute PMG.
3. **Select Execution Backend:**
   - No PowerPoint installed -> **Mode A (Web/Sandbox)**. Compiles native PPTX and executes Level 1 structural validation.
   - Windows with PowerPoint -> **Mode B (Local Desktop)**. Compiles native PPTX and executes Level 1, Level 2, and Level 3 playback validation.
   - Remote worker configured -> **Mode C (Remote Worker)**. Compiles natively and submits candidate for remote Windows PowerPoint playback validation.
4. **Build Presentation Using Shared Native Compiler:**
   Compile from project JSON specification, dictionary, or preset:
   ```python
   result = runtime.build(
       project="presentation_spec.json",  # or builder function or preset name
       output="presentation.pptx",
       validate=True,
   )
   ```
   Or via CLI:
   ```bash
   pmg build presentation_spec.json -o presentation.pptx
   pmg build --preset space --slides 10 -o space.pptx
   ```
5. **Run All Available Validation Levels:**
   - **Level 1 (Structural):** Always executed. Validates ZIP integrity, XML well-formedness, OPC ContentTypes, relationships, unique shape IDs, animation target existence, and `<p:timing>` trees.
   - **Level 2 (Visual):** Executed when slide renderer (PowerPoint or LibreOffice) is available. Checks boundary overflow and generates slide previews. If unavailable, reports `SKIPPED — No slide rendering backend available`.
   - **Level 3 (Playback):** Executed when Microsoft PowerPoint automation is genuinely available. Verifies clean opening (`ppAlertsAll=2`, no repair prompts) and Slide Show playback. If unavailable, reports:
     `NOT_TESTED — PowerPoint Desktop unavailable`.
     *Never report Level 3 PASS without actual Microsoft PowerPoint Desktop execution.*
6. **Deliver Deliverables to User:**
   Deliver the final validated `.pptx` file along with its machine-readable validation report (`<presentation>.pptx.validation_report.json`).
7. **Explain Genuine Limitations:**
   Report honest environment capabilities (e.g. "Compiled native 100% editable PPTX with Morph transitions and Animation Pane effects. Level 1 Structural: PASS. Level 3 Playback: NOT_TESTED — PowerPoint Desktop unavailable in this sandbox environment").

---

## Offline-First Dependency Strategy

When setting up or resolving dependencies:
1. **Existing Installed:** Uses installed `python-pptx`, `lxml`, `Pillow`. (Optional: `numpy`, `opencv-python`, `scipy` for CV analysis; `pywin32` for COM).
2. **Bundled / Cached:** Resolves wheels from `PMG_OFFLINE_CACHE`, `~/.pmg/cache`, or `./wheels`.
3. **Standalone Executable:** Uses standalone binary if provided.
4. **Network Installation:** Runs `pip install` only when network access is explicitly permitted.
5. **Alternative Backend:** Falls back gracefully to standard library tools for read-only inspection.
6. **Unsupported Report:** Explicitly reports any genuinely missing required capability.

---

## Command-Line Interface (CLI)

```bash
# Environment diagnostics and capability inspection
pmg doctor
pmg capabilities

# Build presentation from JSON specification or preset
pmg build presentation_spec.json -o presentation.pptx
pmg build --preset space --slides 20 -o Space.pptx

# Validate existing presentation across Level 1, Level 2, Level 3
pmg validate presentation.pptx
pmg validate presentation.pptx -o Validated.pptx --max-repairs 3

# Render, record, and inspect (when PowerPoint / FFmpeg is available)
pmg inspect presentation.pptx --timeline-out expected_timeline.json
pmg render presentation.pptx -o powerpoint_export.mp4 --fps 60
pmg record presentation.pptx -o slideshow_recording.mp4 --fps 30
```
