---
name: pmg-morph
description: "Master PowerPoint 60 FPS Morph transition engine: OpenXML <p159:morph> injection, '!!' persistent identifier naming invariants, dynamic thematic traveling entities (satellite, vinyl, leaf, arrow, core), shape deformation rules, and aspect-ratio locks."
---

# PMG Morph Engine (`pmg-morph`)

PowerPoint Morph is the hardware-accelerated motion interpolation engine of PMG. When configured with strict OpenXML attributes, it produces broadcast-grade 60 FPS continuous animations matching Adobe After Effects and Apple Keynote.

---

## 1. The `!!` Naming Invariant

To morph **Shape A** on Slide 1 into **Shape B** on Slide 2, both shapes must share an identical custom name prefixed with `!!` in PowerPoint's shape tree:

```xml
<!-- Slide 1 -->
<p:cNvPr id="101" name="!!Hero_Focus"/>

<!-- Slide 2 -->
<p:cNvPr id="205" name="!!Hero_Focus"/>
```

When PowerPoint detects matching `!!` prefixes:
1. It bypasses default cross-fade.
2. It linearly interpolates coordinates: `(x1, y1, w1, h1, rot1) -> (x2, y2, w2, h2, rot2)`.
3. It smoothly blends fills, outlines, and opacity channels across 60 FPS.

---

## 2. Dynamic Thematic Traveling Entity ("Vật Thể Bay Xuyên Slide")

A persistent traveling entity must **match the specific topic** of the presentation:

- **Space & Aerospace:** 3D Rotating Planet, Satellite Probe, Rocket, Astronaut
- **Music & Media (Spotify Style):** Glowing Vinyl Disc, Soundwave Pulse, 3D Floating Earbud
- **Technology & AI:** Cyber Core, Neural Data Orb, Glowing Tech Badge
- **Environment & Ecology:** Floating Green Leaf, Pure Water Droplet, Sprout
- **Business & Finance:** 3D Growth Arrow, Golden Token, Metric Compass Pin
- **Healthcare & Biotech:** DNA Double Helix, Molecular Sphere, Heartbeat Pulse
- **Culture & Literature:** Floating Lantern, Calligraphy Quill, Imperial Seal

```
Slide 1 (Entry / Launch):
[ !!Traveling_Hero ] at (x: -1.5", y: 4.2", rot: +20°) ---> Enters from left margin
           ↓  (Morph 1.6s)
Slide 2 (Ascent & Survey):
                 [ !!Traveling_Hero ] at (x: 5.8", y: 1.8", rot: -10°) ---> Glides above Hero Title
           ↓  (Morph 1.6s)
Slide 3 (Spotlight & Docking):
                                   [ !!Traveling_Hero ] at (x: 10.2", y: 4.0", rot: +15°) ---> Hovers near Key Metrics
           ↓  (Morph 1.6s)
Slide 4 (Graceful Exit):
                                                         [ !!Traveling_Hero ] at (x: 14.5", y: 1.0", rot: -25°) ---> Departs into right margin
```

### Staging Rules for Traveling Entities
1. **Off-Screen Safe Coordinates:**
   - Left Off-Screen: $X \le -2.0 \text{ inches}$ ($X \le -1828800 \text{ EMUs}$)
   - Right Off-Screen: $X \ge 13.5 \text{ inches}$ ($X \ge 12344400 \text{ EMUs}$)
   - Top Off-Screen: $Y \le -2.0 \text{ inches}$
   - Bottom Off-Screen: $Y \ge 7.8 \text{ inches}$
2. **Dynamic Rotation (`rot`):**
   - In PowerPoint OpenXML, rotation is expressed in 60,000ths of a degree (e.g., $15^\circ = 900,000$).
   - As the entity ascends, tilt upwards ($-10^\circ$ to $-20^\circ$); as it swoops down, tilt forwards ($+15^\circ$ to $+25^\circ$).
3. **Aspect-Ratio Lock:**
   $$\frac{W_2}{W_1} = \frac{H_2}{H_1}$$
   Never scale width or height disproportionately during flight.

---

## 3. OpenXML Morph Specification

The morph transition element must be injected into the slide's XML root (`ppt/slides/slide{N}.xml`):

```xml
<p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
              xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"
              spd="slow"
              advClick="1">
  <p159:morph option="byObject"/>
</p:transition>
```

### Morph Option Modes
- `option="byObject"`: The standard 60 FPS spatial interpolation mode for shapes, cards, 3D models, and traveling entities.
- `option="byWord"`: Fluid kinetic typography morphing where shared words glide to their new positions.
- `option="byChar"`: Anagram-style character deformation.

### Transition Timing Standards
- **Standard Spatial Flow:** `duration = 1400ms – 1600ms`.
- **Traveling Mascot Flight:** `duration = 1600ms – 1800ms`.
- **Macro Camera Zoom:** `duration = 1800ms – 2000ms`.

---

## 4. Automated CLI & Python API

Inject Morph transitions using:
```bash
python scripts/pptx_morph_injector.py Presentation.pptx --type morph --dur 1600 --mode byObject
```

Or in Python code:
```python
from scripts.motion_graph_kit import MotionDeck

deck = MotionDeck()
s1 = deck.add_slide(transition="none")
deck.add_thematic_traveling_entity(s1, "assets/satellite.png", left=-1.5, top=4.0, rot_deg=20)

s2 = deck.add_slide(transition="morph", duration_ms=1600)
deck.add_thematic_traveling_entity(s2, "assets/satellite.png", left=6.0, top=1.8, rot_deg=-10)
```
