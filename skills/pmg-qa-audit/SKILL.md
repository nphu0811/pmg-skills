---
name: pmg-qa-audit
description: "Automated QA audit engine and quality standard for PowerPoint Motion Graphics: slide inspection, transition coverage verification, traveling entity tracking, word-count limits (<= 25-35 words), and 60 FPS video export protocols."
---

# PMG QA & Audit Engine (`pmg-qa-audit`)

A presentation is only marked as **PASSED / ACCEPTED** when it fulfills 100% of the rigorous automated quality assurance criteria.

---

## 1. Automated Audit Criteria

Every PMG presentation must satisfy the 5-point verification matrix:

| Metric | Required Threshold | Verification Method |
| :--- | :--- | :--- |
| **Transition Coverage** | **100% of transitions orchestrated** | OpenXML parser verifies `<p159:morph>`, `<p:circle>`, `<p:fade>`, `<p:split>` |
| **Traveling Entity Tracking** | **Verified '!!' continuous anchors** | Checks coordinate continuity across consecutive slides |
| **Word Count (Visual-First)** | **$\le 25$ words/slide in 3D, $\le 35$ in Academic** | Text token counter across all shapes |
| **Aspect Ratio Integrity** | **Zero distortion** | Verifies $W_2/W_1 = H_2/H_1$ across `!!` shapes |
| **Zero Overlap** | **$0$ Collisions** | Bounding box spatial intersection check |

---

## 2. Using `scripts/slide_inspector.py`

Run the automated inspection tool against any generated `.pptx`:

```bash
python scripts/slide_inspector.py Deck_Output.pptx
```

### Sample Output:
```text
======================================================================
[*] RUNNING MOTION & TRANSITION QA AUDIT ON: Demo_Showcase.pptx
======================================================================
[+] Total slides detected: 5
  Slide 01: Transition [PASS (None / Cover)  ] | Words: 12 [PASS]
  Slide 02: Transition [PASS (Morph)         ] | Words: 18 [PASS]
  Slide 03: Transition [PASS (Morph)         ] | Words: 21 [PASS]
  Slide 04: Transition [PASS (Morph)         ] | Words: 19 [PASS]
  Slide 05: Transition [PASS (Circle Wipe)   ] | Words: 14 [PASS]
----------------------------------------------------------------------
[+] Transition Choreography Coverage: 5/5 slides (100.0%)
[+] Traveling Morph Anchors Detected: 2 ['!!Traveling_Hero', '!!prog_fill']
[+] Visual-First Word Count Compliance: 5/5 slides (100.0%)
[+] Average Words/Slide: 16.8 words
======================================================================
[RESULT: VERIFIED & ACCEPTED] Presentation meets PMG Motion Design standards!
======================================================================
```

---

## 3. High-Definition 60 FPS Video Export Protocol

Never deliver static slide screenshots when reviewing motion graphics. Export full 60 FPS broadcast-grade MP4 video directly from PowerPoint:

```python
from scripts.motion_graph_kit import export_deck_to_video

# Automatically exports native 60 FPS MP4 video with full transitions
video_file = export_deck_to_video(
    pptx_path="examples/showcase.pptx",
    video_path="examples/showcase_demo.mp4",
    slide_duration=3.5,
    fps=60,
    resolution=1080
)
```

PowerPoint's hardware rendering engine captures all 60 FPS Morph interpolations, 3D rotations, and traveling entity flights in crisp 1080p video.
