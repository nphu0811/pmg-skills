---
name: pmg-kinetic-typography
description: Specialized kinetic typography and text motion intelligence for PMG. Animates authoritative headline reveals, word-by-word reading cadences, keyword scale punches, typewriter rhythms, and kinetic cross-slide morphs.
---

# PMG Kinetic Typography (`pmg-kinetic-typography`)

Specialized motion intelligence for headlines, titles, body paragraphs, and keywords in PowerPoint presentations.

## Core Mental Model

> **Typography motion must guide reading cadence, not create visual noise.**
> - Titles command the slide with authoritative, steady unveiling.
> - Keywords punch forward with subtle scale pops (1.12x - 1.15x) to anchor memory.
> - Word-by-word cascades respect human reading speed (40ms - 80ms micro-delays).
> - Cross-slide headers morph smoothly preserving the persistent `!!` anchor.
> - Monospaced code blocks unveil with crisp typewriter rhythm.

---

## Supported Patterns

| Pattern | Shared Primitive | Native PPTX Effect | Semantic Application |
| :--- | :--- | :--- | :--- |
| **`title_unveiling`** | `REVEAL` | `wipe` / `fade` | Authoritative headline establishing narrative topic. |
| **`word_stagger`** | `STAGGER` | `appear` (staggered delay) | Rhythmic word-level reveal following reading order. |
| **`keyword_punch`** | `EMPHASIZE` | `scale` (1.14x) | High-contrast visual pop commanding focal attention. |
| **`typewriter`** | `REVEAL` | stepped `appear` | Terminal code / computational data sequence. |
| **`subtitle_glide`** | `REVEAL` | `fade` / `fly_in` (`up`) | Secondary context settling beneath the title. |
| **`kinetic_morph`** | `MORPH` | `<p159:morph>` | Cross-slide traveling headline transition (`!!`). |

---

## Python API Usage

```python
from pmg.motion_planners import KineticTypographyPlanner

planner = KineticTypographyPlanner()
intents = planner.plan_motion({
    "style": "keyword_punch",
    "target_id": "headline_phrase",
    "keyword_ids": ["kw_autonomous", "kw_realtime"],
})

for intent in intents:
    print(f"[{intent.target_id}] -> {intent.primitive.value} ({intent.semantic_purpose})")
```
