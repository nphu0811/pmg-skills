---
name: pmg-map-motion
description: Specialized geospatial and map motion intelligence for PMG. Animates region highlights, progressive route drawings, location marker pin drops, hub radar pulses, and multi-location itinerary storytelling in native PowerPoint.
---

# PMG Map Motion (`pmg-map-motion`)

Specialized motion intelligence for geographic maps, territorial markets, logistics routes, and location markers in PowerPoint presentations.

## Core Mental Model

> **Preserve geographic truth, spatial proportions, and verified coordinates.**
> - Landmass baselines settle into place to orient the viewer.
> - Location markers drop into verified coordinates with authoritative settling.
> - Hub radar rings ripple outward to indicate active regional clusters.
> - Logistics and travel routes progressively trace along the travel trajectory.
> - Telemetry and demographic overlays dock adjacent to active territory highlights.

---

## Supported Patterns

| Pattern | Shared Primitive | Native PPTX Effect | Semantic Application |
| :--- | :--- | :--- | :--- |
| **`region_highlight`** | `HIGHLIGHT` | `pulse` / `fade` | Illuminates territory polygon with glowing accent fill. |
| **`route_progression`** | `DRAW` | `wipe` (directional) | Route path vector traces from origin to destination. |
| **`marker_drop`** | `MOVE` | `fly_in` (`down`) | Pin marker drops into geographic coordinate. |
| **`marker_pulse`** | `HIGHLIGHT` | `pulse` | Concentric radar ring ripples from active hub. |
| **`geo_zoom`** | `FOCUS` | `scale` (1.15x - 1.25x) | Macro push-in focusing on metropolitan sub-region. |

---

## Python API Usage

```python
from pmg.motion_planners import MapMotionPlanner

planner = MapMotionPlanner()
intents = planner.plan_motion({
    "pattern": "route_progression",
    "base_map_id": "vietnam_map",
    "locations": [
        {"id": "pin_hanoi", "label": "Hanoi Hub", "pulse_id": "pulse_hanoi"},
        {"id": "pin_danang", "label": "Da Nang Node"},
        {"id": "pin_hcmc", "label": "HCMC Hub", "pulse_id": "pulse_hcmc"},
    ],
    "routes": [
        {"id": "route_hn_dn", "direction": "down"},
        {"id": "route_dn_hcm", "direction": "down"},
    ]
})

for intent in intents:
    print(f"[{intent.target_id}] -> {intent.primitive.value} ({intent.semantic_purpose})")
```
