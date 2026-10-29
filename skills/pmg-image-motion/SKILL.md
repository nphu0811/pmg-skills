---
name: pmg-image-motion
description: Specialized image and visual composition motion intelligence for PMG. Animates directional photographic reveals, layered multi-plane parallax, macro dolly zooms, cinematic pans, and before/after comparisons.
---

# PMG Image Motion (`pmg-image-motion`)

Specialized motion intelligence for photographs, illustrations, and decomposed visual layers in PowerPoint presentations.

## Core Mental Model

> **Preserve photographic integrity, aspect ratios, and spatial depth.**
> - Directional reveals follow the photograph's natural leading lines.
> - Multi-plane parallax uses authentic speed ratios: Background (0.7x), Midground Subject (1.0x), Foreground Scrim (1.4x).
> - Macro dolly push-in zooms are strictly clamped ($\le 125\%$) to prevent raster pixelation.
> - Before/after comparisons synchronize wipe boundaries with visual divider lines.
> - Foreground cutouts separate from background atmosphere with subtle depth elevation.

---

## Supported Patterns

| Pattern | Shared Primitive | Native PPTX Effect | Semantic Application |
| :--- | :--- | :--- | :--- |
| **`directional_reveal`** | `REVEAL` | `wipe` | Unmasks image along composition leading lines. |
| **`layered_parallax`** | `MOVE` | `motion_path` | Differential 3-plane speed ratios (0.7x / 1.0x / 1.4x). |
| **`controlled_zoom`** | `SCALE` | `scale` (1.10x - 1.20x) | Subtle macro push-in revealing fine visual texture. |
| **`pan`** | `MOVE` | `motion_path` (linear) | Horizontal tracking pan across wide panoramic views. |
| **`before_after`** | `DRAW` | `wipe` (`right`) | Transformation state wiped over baseline reference. |
| **`depth_separation`** | `FOCUS` | `scale` / `fade` | Cutout lifts forward while background softens. |

---

## Python API Usage

```python
from pmg.motion_planners import ImageMotionPlanner

planner = ImageMotionPlanner()
intents = planner.plan_motion({
    "pattern": "layered_parallax",
    "background_id": "nebula_bg",
    "target_id": "spacecraft_subject",
    "foreground_id": "hud_reticle",
})

for intent in intents:
    print(f"[{intent.target_id}] -> {intent.primitive.value} ({intent.semantic_purpose})")
```
