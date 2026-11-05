---
name: pmg-chart-motion
description: Specialized quantitative chart and numerical motion intelligence for PMG. Animates bar growth, line trajectories, doughnut slice reveals, synchronized metric counters, and peak metric emphasis in native PowerPoint.
---

# PMG Chart Motion (`pmg-chart-motion`)

Specialized motion intelligence for quantitative data charts, numerical value counters, and metric indicators in PowerPoint presentations.

## Core Mental Model

> **Never apply generic Fade In to quantitative data.**
> Animation must reflect the data's geometry, scale, and narrative takeaway:
> - Coordinate axes ground the visual field before numbers appear.
> - Data bars grow directionally from the baseline (upward for column, rightward for horizontal).
> - Time-series lines trace chronologically from left to right.
> - Doughnut sectors reveal proportionally in radial progression.
> - Numeric counters synchronize with physical bar tips.
> - The peak metric or conclusion is emphasized with an intentional scale punch.

---

## Supported Patterns

| Pattern | Shared Primitive | Native PPTX Effect | Semantic Application |
| :--- | :--- | :--- | :--- |
| **`axis_setup`** | `DRAW` | `wipe` | Grounds coordinate baseline before data disclosure. |
| **`bar_growth`** | `GROW` | `wipe` (`up` or `right`) | Bar grows from zero to quantitative target value. |
| **`line_draw`** | `DRAW` | `wipe` (`right`) | Line charts draw along chronological time axis. |
| **`doughnut_reveal`** | `GROW` | `wheel` / `zoom` | Slices unfold sequentially in proportional radial order. |
| **`value_counter_sync`** | `COUNT` | `appear` | Numeric callout synchronizes with bar growth. |
| **`peak_emphasis`** | `EMPHASIZE` | `scale` (1.06x - 1.12x) | Focal pop drawing eye to the primary takeaway metric. |

---

## Python API Usage

```python
from pmg.motion_planners import ChartMotionPlanner

planner = ChartMotionPlanner()
intents = planner.plan_motion({
    "chart_type": "bar",
    "axis_id": "chart_axes",
    "items": [
        {"id": "bar_q1", "value": "$1.2M", "label_id": "val_q1"},
        {"id": "bar_q2", "value": "$1.8M", "label_id": "val_q2"},
        {"id": "bar_q3", "value": "$2.9M", "label_id": "val_q3", "is_peak": True},
    ]
})

for intent in intents:
    print(f"[{intent.target_id}] -> {intent.primitive.value} ({intent.semantic_purpose})")
```
