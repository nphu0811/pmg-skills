---
name: pmg-camera-parallax
description: "Virtual multi-plane camera choreography and parallax mechanics: wide establishing shots, macro push-in zooms (100% to 240%), 3-plane speed ratios (BG 0.7x, Mid 1.0x, FG 1.4x), and cinematic tracking pans."
---

# PMG Camera & Parallax Engine (`pmg-camera-parallax`)

A presentation should not feel like flat slides changing; it must feel like a **continuous 3D space viewed through a virtual cinematic camera**.

---

## 1. Multi-Plane Parallax Architecture

To simulate genuine physical camera depth, elements are segregated into 3 distinct spatial planes moving at differentiated velocities during Morph transitions:

```
+---------------------------------------------------------------------------------+
| PLANE 1: BACKGROUND GRID / HORIZON (Shift: 0.7x)                                |
| Moves slowly across slides to create infinite horizon depth.                    |
| Identifier: !!BG_Studio, !!BG_Grid                                              |
|                                                                                 |
| PLANE 2: 3D MODEL / MACHINE MIDGROUND (Shift: 1.0x, Zoom: 100% -> 240%)         |
| Primary focal anchor. Undergoes camera push-in and lateral panning.             |
| Identifier: !!Hero3D_Focus, !!Midground_Core                                    |
|                                                                                 |
| PLANE 3: FOREGROUND PARTICLES / MIST (Shift: 1.4x)                             |
| Floats close to the virtual lens, moving faster to exaggerate depth.            |
| Identifier: !!FG_Dust, !!FG_LightStreak                                         |
+---------------------------------------------------------------------------------+
```

### Parallax Mathematics
When the virtual camera pans by $\Delta X_{\text{camera}}$:
$$\Delta X_{\text{BG}} = 0.7 \times \Delta X_{\text{camera}}$$
$$\Delta X_{\text{MID}} = 1.0 \times \Delta X_{\text{camera}}$$
$$\Delta X_{\text{FG}} = 1.4 \times \Delta X_{\text{camera}}$$

---

## 2. Virtual Camera Move Types

### A. Wide Establishing Shot (Zoom = 100%)
- **Purpose:** Orients the audience by displaying the complete system or facility.
- **Coordinates:** `x = 0.0", y = 0.0", w = 13.333", h = 7.5"`.

### B. Macro Push-In Zoom (Zoom = 180% – 240%)
- **Purpose:** Zooms deep inside a specific subsystem (e.g., combustion chamber, turbine blades, microchip core).
- **Coordinate Formula:**
  $$W_{\text{zoomed}} = W_{\text{base}} \times Z$$
  $$H_{\text{zoomed}} = H_{\text{base}} \times Z$$
  $$X_{\text{zoomed}} = X_{\text{screen\_target}} - (X_{\text{subsystem\_rel}} \times Z)$$
  $$Y_{\text{zoomed}} = Y_{\text{screen\_target}} - (Y_{\text{subsystem\_rel}} \times Z)$$

### C. 3D Tracking Pan
- **Purpose:** Camera slides laterally along an energy conduit or pipeline, maintaining focus on the moving energy packet.

### D. Pull-Back Reveal
- **Purpose:** Smoothly decelerates from a macro detail back to the global overview, illuminating the entire synchronized system.

---

## 3. Python Camera Control Pattern

```python
from scripts.forge_lib import Cinematic3DDeck

deck = Cinematic3DDeck(title="AI Neural Core")

# Slide 1: Wide Overview (Zoom 1.0x)
s1 = deck.add_cinematic_slide(
    scene_idx=1,
    camera_zoom=1.0,
    camera_focus_center=(6.66, 3.75),
    header_title="NEURAL SUPERCOMPUTER",
    callout_title="CLUSTER OVERVIEW",
    callout_desc="High-density parallel processing architecture."
)

# Slide 2: Macro Push-In into Tensor Core (Zoom 2.2x)
s2 = deck.add_cinematic_slide(
    scene_idx=2,
    camera_zoom=2.2,
    camera_focus_center=(4.2, 3.8), # Targets tensor processor coordinates
    header_title="TENSOR PROCESSING UNIT",
    callout_title="FP8 MATRIX CORE",
    callout_desc="100 petaflops sparse mathematical execution."
)
```
