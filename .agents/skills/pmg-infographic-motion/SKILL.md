---
name: pmg-infographic-motion
description: Specialized infographic, timeline, and progress indicator motion intelligence for PMG. Animates chronological timeline spines, milestone pin reveals, horizontal progress bars, circular progress gauges, and comparison matrices.
---

# PMG Infographic Motion (`pmg-infographic-motion`)

Specialized motion intelligence for timelines, progress gauges, feature comparisons, and structured infographics in PowerPoint presentations.

## Core Mental Model

> **Animation must mirror informational structure and chronological flow.**
> - Timeline spines draw first to establish the temporal horizon.
> - Chronological milestones anchor onto the spine in strict sequential order.
> - Progress bars advance smoothly from 0% toward completion target.
> - Percentage numeric labels track active progress fill tips synchronously.
> - Strategic comparison matrices unveil side-by-side before highlighting the winning option.

---

## Supported Patterns

| Pattern | Shared Primitive | Native PPTX Effect | Semantic Application |
| :--- | :--- | :--- | :--- |
| **`timeline_progression`**| `DRAW` + `FOCUS` | `wipe` + `zoom` | Spine traces $\to$ milestones anchor sequentially. |
| **`progress_bar`** | `GROW` | `wipe` (`right`) | Active fill expands to exact completion percentage. |
| **`percent_tracking`** | `TRACK` | `appear` | Percentage counter tracks advancing bar tip. |
| **`circular_gauge`** | `GROW` + `EMPHASIZE` | `wheel` + `scale` | Radial dial unfolds with central value punch. |
| **`feature_comparison`**| `STAGGER` | `fly_in` (`up`) | Options appear side-by-side for structured review. |
| **`winner_callout`** | `EMPHASIZE` | `scale` (1.15x) | Strategic badge highlights recommended alternative. |

---

## Python API Usage

```python
from pmg.motion_planners import InfographicMotionPlanner

planner = InfographicMotionPlanner()
intents = planner.plan_motion({
    "pattern": "timeline_progression",
    "spine_id": "time_spine",
    "milestones": [
        {"node_id": "m1_pin", "card_id": "m1_card", "year": "2024"},
        {"node_id": "m2_pin", "card_id": "m2_card", "year": "2025"},
        {"node_id": "m3_pin", "card_id": "m3_card", "year": "2026"},
    ]
})

for intent in intents:
    print(f"[{intent.target_id}] -> {intent.primitive.value} ({intent.semantic_purpose})")
```
