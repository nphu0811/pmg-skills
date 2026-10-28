---
name: pmg-engine
description: "Main canonical entrypoint for PMG Skills (Engine-First Architecture). Enforces generation via the official PMG Python engine, native editable PowerPoint objects, mandatory real PowerPoint slideshow playback recording (local or via Windows Validation Worker), AI video & visibility review, 3-loop self-healing, and validated PPTX delivery."
---

# PMG Canonical Engine & Delivery Pipeline (`pmg-engine`)

> **Mandatory Engine-First Directive:** When a user requests a presentation using PMG Skills, the AI agent **MUST** execute the official `pmg` engine (`pmg.orchestrator.PMGOrchestrator` or `python -m pmg`). Do **NOT** bypass the engine with ad-hoc `python-pptx` scripts, and **NEVER** deliver a `.pptx` until it has passed real Microsoft PowerPoint slideshow playback and AI video + object visibility validation.

---

## 1. Mandatory 9-Step Execution & Delivery Pipeline

Every presentation generated through PMG Skills must pass through `PMGOrchestrator`:

1. **Step 1 — Build (`pmg build`):** Construct the `.pptx` using canonical `pmg` engine builders with 100% native editable shapes, text boxes, and animations (never flattened slide screenshots or background videos).
2. **Step 2 — Static Validation (`pmg inspect` / `StaticValidator`):** Verify OpenXML package structure, unique shape IDs, `!!` Morph anchors, transitions, `<p:timing>` entrance `<p:set>` visibility and `grpId`/`<p:bldLst>` synchronization, and generate `expected_timeline.json`.
3. **Step 3 — Native PowerPoint Video & Reference Image Export (`pmg render`):** Export static edit-mode slide reference PNGs (`edit_mode_refs/edit_slide_XX.png`) and `powerpoint_export.mp4` via Microsoft PowerPoint COM (`1920x1080`, 60 FPS requested) and inspect actual metadata via `ffprobe`.
4. **Step 4 — Actual Slideshow Playback Recording (`pmg record`):** Launch Microsoft PowerPoint Desktop in full-screen Slide Show mode (locally or on a remote Windows Validation Worker via `PMG_QA_WORKER_URL`), drive transitions/clicks via `expected_timeline.json`, and record `actual_slideshow_recording.mp4`.
5. **Step 5 — AI & Computer Vision Visual & Visibility Review (`pmg review`):** Extract keyframes (`frame_samples/`), verify object trajectories, compare Edit-Mode reference images against live Slide Show steady states (ensuring hero images and titles never vanish behind sibling outlines), and output `video_analysis.json` + `review_report.json`.
6. **Step 6 — Automatic 3-Loop Self-Healing (`pmg.self_healing`):**
   - **Loop 1 (Targeted):** Heal missing transitions, inject missing `<p:set>` (`style.visibility = "visible"`), synchronize `grpId` / `<p:bldLst>`, remove duplicate conflicting entrances while preserving Fade In / Wipe / Fly In / Motion Path animations, and normalize durations.
   - **Loop 2 (Structural):** Interpolate and bridge broken `!!` traveling anchor trajectories and recompile canonical `<p:timing>` trees via `SlideAnimationTimeline`.
   - **Loop 3 (Compatible Fallback):** Substitute unresolved effects with compatible native PowerPoint transitions/entrances while preserving all editable objects.
7. **Step 7 — Repeat Rendering & Recording:** After every repair loop, rebuild, re-export video, re-record live slideshow, and re-run visual review.
8. **Step 8 — Final Acceptance & Embedded Provenance:** Require `status == "PASS"` in `execution_manifest.json` and embed `ppt/pmg_execution_manifest.json` inside the delivered `.pptx`.
9. **Step 9 — Validated Delivery Policy:** Deliver **ONLY** the final validated `.pptx` to the user. Keep `internal_qa/<run_id>/` videos and JSON reports internal unless explicitly requested.

---

## 2. Canonical CLI Usage

```bash
# Build & run full 9-step playback validation + 3-loop self-healing:
python -m pmg build -o Space_Exploration.pptx --preset space --slides 20

# Validate an existing PPTX (auto, local, remote Windows worker, or static):
python -m pmg validate Input_Deck.pptx -o Validated_Deck.pptx --worker auto
python -m pmg validate Input_Deck.pptx --worker remote --worker-url http://192.168.1.50:8765

# Windows Validation Worker management (for Linux / Cloud AI agents):
python -m pmg worker doctor
python -m pmg worker serve --host 0.0.0.0 --port 8765 --token "$PMG_QA_WORKER_TOKEN"
python -m pmg worker status --url http://192.168.1.50:8765

# 24-Case Animation Benchmark Suite:
python -m pmg benchmark run --all
python -m pmg benchmark run --case 21
python -m pmg benchmark report
```
