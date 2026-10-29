---
name: pmg-engine
description: "Universal PowerPoint Motion Graphics (PMG) v2.4.0 engine. Cross-platform runtime for AI web sandboxes, local AI agents, and desktop environments. Generates 100% native editable PPTX with 60 FPS Morph, intra-slide Animation Pane timing trees, distributed QA, transactional self-healing, and immutable certification without requiring Microsoft PowerPoint to be installed."
---

# PMG Universal Runtime & Motion Intelligence Engine (v2.4.0)

PMG v2.4.0 provides a production-quality, capability-aware PowerPoint generation, validation, and self-healing system.
The compiler produces 100% native, editable PowerPoint presentations across:
- **Mode A: Web / Sandbox** (Claude Web, ChatGPT sandbox, restricted Linux environments without Microsoft PowerPoint or internet).
- **Mode B: Local Desktop** (Windows, macOS, Linux with local Python and optional PowerPoint COM/AppleScript).
- **Mode C: Remote Worker** (configured & authenticated Windows validation worker with Bearer token authentication).

**Critical Invariant: Creating native PowerPoint animations does NOT require Microsoft PowerPoint to be installed.**
Microsoft PowerPoint is an optional Level 3 validation and playback backend, never a compilation prerequisite.

---

## Unified Architecture Overview

- **v2.1.0 Universal Runtime:** Deterministic capability detection (`pmg doctor`), offline-first 6-tier dependency resolution, and Three-Level Validation.
- **v2.2.0 Native Animation Compiler:** Decoupled Animation Intermediate Representation (`AnimationIR`), explicit Capability Catalog, deterministic `TimelineCompiler` with cycle detection, and `MorphReliabilityValidator` for persistent traveling entity tracking.
- **v2.3.0 Distributed PowerPoint QA:** Remote Windows PowerPoint validation worker, secure Bearer token authentication, artifact SHA-256 integrity verification, and idempotent job lifecycle (`CREATED -> QUEUED -> RUNNING -> COMPLETED / FAILED / TIMED_OUT / CANCELLED`).
- **v2.4.0 Self-Healing Compatibility:** Standardized `ErrorCode` diagnostics, transactional repair pipeline with automatic rollback (`TransactionalRepairManager`), bounded 3-retry limit, safe automatic repair of missing ContentTypes and duplicate IDs, and `ImmutableValidationCertificate` cryptographically bound to artifact SHA-256.

---

## Universal Agent Execution Workflow

All AI agents (Claude Web, ChatGPT, Claude Code, Codex, local agents) follow this canonical execution workflow:

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
   Compile from project JSON specification, dictionary, or `AnimationIR`:
   ```python
   result = runtime.build(
       project="presentation_spec.json",  # or AnimationIR instance or preset
       output="presentation.pptx",
       validate=True,
       repair="safe",
   )
   print(result.output_path)
   print(result.artifact_sha256)
   print(result.validation_report.overall_status)
   ```
   Or via CLI:
   ```bash
   pmg build presentation_spec.json -o presentation.pptx --repair safe
   pmg build --preset space --slides 10 -o space.pptx
   ```
5. **Run All Available Validation Levels:**
   - **Level 1 (Structural):** Always executed. Validates ZIP integrity, XML well-formedness, OPC ContentTypes, relationships, unique shape IDs, animation target existence, and `<p:timing>` trees.
   - **Level 2 (Visual):** Executed when slide renderer (PowerPoint or LibreOffice) is available. Checks boundary overflow and generates slide previews. If unavailable, reports `SKIPPED — No slide rendering backend available`.
   - **Level 3 (Playback):** Executed when Microsoft PowerPoint automation is genuinely available. Verifies clean opening (`ppAlertsAll=2`, no repair prompts) and Slide Show playback. If unavailable, reports:
     `NOT_TESTED — PowerPoint Desktop unavailable`.
     *Never report Level 3 PASS without actual Microsoft PowerPoint Desktop execution.*
6. **Apply Transactional Self-Healing if Defects Detected:**
   - When defects are detected, invoke `TransactionalRepairManager`.
   - Apply only `SAFE_AUTOMATIC` repairs: missing ContentTypes, duplicate shape IDs, dangling animation targets.
   - Automatically roll back to the original if repair does not achieve a valid state.
7. **Deliver Deliverables to User:**
   Deliver the final validated `.pptx` file along with its machine-readable validation report and immutable validation certificate.

---

## Command-Line Interface (CLI)

```bash
# Environment diagnostics and capability inspection
pmg doctor
pmg capabilities

# Build presentation from JSON specification, AnimationIR, or preset
pmg build presentation_spec.json -o presentation.pptx --repair safe
pmg build --preset space --slides 20 -o space.pptx

# Validate existing presentation across Level 1, Level 2, Level 3
pmg validate presentation.pptx [--json]

# Transactional safe self-healing
pmg repair presentation.pptx --safe -o presentation_healed.pptx

# Distributed or local PowerPoint QA
pmg qa presentation.pptx
pmg qa presentation.pptx --remote --worker-url https://worker.internal:8765

# Render, record, and inspect (when PowerPoint / FFmpeg is available)
pmg inspect presentation.pptx --timeline-out expected_timeline.json
pmg render presentation.pptx -o powerpoint_export.mp4 --fps 60
pmg record presentation.pptx -o slideshow_recording.mp4 --fps 30
```
