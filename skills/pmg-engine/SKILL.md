---
name: pmg-engine
description: "Main canonical entrypoint for PMG Skills (Engine-First Architecture). Enforces generation via the official PMG Python engine, native editable PowerPoint objects, mandatory real PowerPoint slideshow playback recording, AI video review, 3-loop self-healing, and validated PPTX delivery."
---

# PMG Canonical Engine & Delivery Pipeline (`pmg-engine`)

> **Mandatory Engine-First Directive:** When a user requests a presentation using PMG Skills, the AI agent **MUST** execute the official `pmg` engine (`pmg.orchestrator.PMGOrchestrator` or `python -m pmg`). Do **NOT** bypass the engine with ad-hoc `python-pptx` scripts, and **NEVER** deliver a `.pptx` until it has passed real Microsoft PowerPoint slideshow playback and AI video validation.

---

## 1. Mandatory 9-Step Execution & Delivery Pipeline

Every presentation generated through PMG Skills must pass through `PMGOrchestrator`:

1. **Step 1 — Build (`pmg build`):** Construct the `.pptx` using canonical `pmg` engine builders with 100% native editable shapes, text boxes, and animations (never flattened slide screenshots or background videos).
2. **Step 2 — Static Validation (`pmg inspect` / `StaticValidator`):** Verify OpenXML package structure, unique shape IDs, `!!` Morph anchors, transitions, media links, and generate `expected_timeline.json`.
3. **Step 3 — Native PowerPoint Video Export (`pmg render`):** Export `powerpoint_export.mp4` via Microsoft PowerPoint COM (`1920x1080`, 60 FPS requested) and inspect actual metadata via `ffprobe`.
4. **Step 4 — Actual Slideshow Playback Recording (`pmg record`):** Launch Microsoft PowerPoint Desktop in full-screen Slide Show mode, drive transitions/clicks via `expected_timeline.json`, and record `actual_slideshow_recording.mp4`.
5. **Step 5 — AI & Computer Vision Visual Review (`pmg review`):** Extract keyframes and transition samples (`frame_samples/`), verify object trajectories, check visual/temporal/capture quality, and output `video_analysis.json` + `review_report.json`.
6. **Step 6 — Automatic 3-Loop Self-Healing (`pmg.self_healing`):**
   - **Loop 1 (Targeted):** Heal missing transitions, timing references, and durations.
   - **Loop 2 (Structural):** Interpolate and bridge broken `!!` traveling anchor trajectories.
   - **Loop 3 (Compatible Fallback):** Substitute unresolved effects with compatible native PowerPoint transitions while preserving all editable objects.
7. **Step 7 — Repeat Rendering & Recording:** After every repair loop, rebuild, re-export video, re-record live slideshow, and re-run visual review.
8. **Step 8 — Final Acceptance:** Require `status == "PASS"` in `execution_manifest.json`.
9. **Step 9 — Validated Delivery Policy:** Deliver **ONLY** the final validated `.pptx` to the user. Keep `internal_qa/<run_id>/` videos and JSON reports internal unless explicitly requested.

---

## 2. Canonical CLI Usage

```bash
# Build & run full 9-step playback validation + 3-loop self-healing:
python -m pmg build -o Protecting_Our_Earth.pptx --preset environment --slides 10

# Validate an existing PPTX through the full real playback pipeline:
python -m pmg validate Input_Deck.pptx -o Validated_Deck.pptx

# Individual pipeline commands:
python -m pmg inspect Validated_Deck.pptx --timeline-out expected_timeline.json
python -m pmg render Validated_Deck.pptx -o powerpoint_export.mp4 --fps 60
python -m pmg record Validated_Deck.pptx -o actual_slideshow_recording.mp4 --fps 30
python -m pmg review Validated_Deck.pptx --export-video powerpoint_export.mp4 --record-video actual_slideshow_recording.mp4
python -m pmg reconstruct Reference_Deck.pptx -o Reconstructed_Deck.pptx
```

---

## 3. Canonical Python API Usage

```python
from pmg import PMGOrchestrator, build_environmental_protection_deck, ValidationStatus

orchestrator = PMGOrchestrator()
manifest = orchestrator.run_pipeline(
    output_pptx_path="Protecting_Our_Earth.pptx",
    builder_fn=lambda path: build_environmental_protection_deck(path, num_slides=10),
    hold_duration_ms=1200,
    max_repair_loops=3,
    require_video_playback=True,
)

assert manifest.status == ValidationStatus.PASS
print("Delivered validated PPTX:", manifest.delivered_pptx)
```

