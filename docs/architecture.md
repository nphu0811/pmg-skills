# PMG Skills Architecture & Engine Responsibilities (v2.1.0)

PMG v2.1.0 introduces the **Universal Runtime Architecture**: a cross-platform foundation enabling native PowerPoint slide compilation and multi-level validation across AI web sandboxes, local developer environments, and remote validation workers.

---

## 1. Core Architecture Principles

1. **PowerPoint-Independent Compilation:**
   Creating native PowerPoint animations (Morph transitions, shape transformations, intra-slide Animation Pane timing trees) requires only pure Python and OpenXML synthesis (`python-pptx`, `lxml`, `Pillow`). Microsoft PowerPoint is **never** a compilation prerequisite.

2. **Three Execution Environments:**
   - **Mode A: Web / Sandbox:** Isolated environments (Claude Web, ChatGPT sandbox, Linux CI without Office). Compiles native PPTX and executes Level 1 structural validation.
   - **Mode B: Local Desktop:** Developer machines (Windows, macOS, Linux). Executes Level 1, Level 2 (if renderer present), and optional Level 3 on Windows with Microsoft PowerPoint Desktop.
   - **Mode C: Remote Worker:** Secure HTTP/HTTPS adapter to a Windows PowerPoint Validation Worker for Slide Show playback testing from non-Windows hosts.

3. **Structured Three-Level Validation:**
   - **Level 1 (Structural):** Mandatory everywhere. Validates ZIP integrity, XML well-formedness, OPC ContentTypes, relationships, unique shape IDs, animation targets, and `<p:timing>` trees.
   - **Level 2 (Visual):** Slide previews rendered via available backend (PowerPoint COM or LibreOffice). Reports `SKIPPED — No slide rendering backend available` if missing.
   - **Level 3 (Playback):** Native Microsoft PowerPoint Slide Show playback. Reports `NOT_TESTED — PowerPoint Desktop unavailable` if PowerPoint is absent. *Never fabricates Level 3 PASS without genuine PowerPoint Desktop execution.*

4. **One Repair Budget & Final Authority:**
   `pmg.orchestrator.PMGOrchestrator` owns the only retry budget (maximum 3 attempts) and the final delivery decision. QA detects defects read-only; self-healing applies targeted repairs.

---

## 2. Module Responsibilities

| Module | Responsibility | Invariants & Boundaries |
| :--- | :--- | :--- |
| `pmg.runtime` | High-level unified API (`PMGRuntime`, `PMGBuildResult`) | Coordinates capability detection, compilation, and validation stages. |
| `pmg.capabilities` | Environment detection & modeling (`detect_environment`) | Deterministic inspection with explicit available/unavailable/unknown states. |
| `pmg.compiler` | Consolidated native OOXML compiler (`NativePPTXCompiler`) | 100% native shapes, text, Morph, and Animation Pane trees without PowerPoint. |
| `pmg.validation` | Three-level validation engine (`validate_presentation`) | Level 1 structural, Level 2 visual, Level 3 native playback. |
| `pmg.dependencies` | Offline-first resolution strategy (`resolve_dependencies`) | 6-tier fallback sequence; supports `PMG_OFFLINE_CACHE` and local `./wheels`. |
| `pmg.orchestrator` | Pipeline orchestration & bounded 3-loop self-healing | Enforces delivery gate, packaging, clean reopen, and external companion manifest. |
| `pmg.animations` | Intra-slide `<p:timing>` Animation Pane synthesis | Compiles entrances, emphasis, motion paths, exits, triggers, and timing graphs. |
| `pmg.validator` | Deep OpenXML static inspection & ExpectedTimeline | Read-only structural defect detection and semantic role classification. |
| `pmg.worker` | Windows PowerPoint Validation Worker (Server/Client) | Headless Slide Show recording, clean-open preflight (`ppAlertsAll=2`), and security controls. |

---

## 3. Data Flow

```
Agent Request / Spec
   ↓
Capability Detection (pmg.capabilities)
   ↓
Dependency Resolution (pmg.dependencies)
   ↓
Shared Native Compiler (pmg.compiler)
   ↓
Generated Native PPTX Artifact
   ↓
Level 1 Structural Validation (pmg.validation)  → [FAIL blocks delivery]
   ↓
Level 2 Visual Validation (Optional Renderer)    → [If missing: SKIPPED]
   ↓
Level 3 Playback Validation (Optional Desktop)   → [If missing: NOT_TESTED]
   ↓
Delivery Gate (Certified PPTX + Validation Report)
```
