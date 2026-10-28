---
name: pmg-3d-cinematic
description: "Prepare and compose 3D assets, model orientation, technical cutaways, exploded views and annotations. Reuses the shared camera system; requires actual PowerPoint model compatibility evidence."
---

# PMG 3D Cinematic Engine (`pmg-3d-cinematic`)

## Responsibility and dependencies

3D asset preparation, model placement/orientation, cutaways, exploded views, annotations, model-focused composition and model compatibility. Use pmg-camera-parallax for all camera transforms; choreography owns high-level sequencing.

Dependencies: pmg-engine; pmg-camera-parallax; pmg-morph; core/asset helpers. All generated candidates must pass
PMGOrchestrator before delivery; direct save/COM snippets below are lower-level
composition examples, not independent production entrypoints.

Transforms PowerPoint into a cinematic 3D documentary experience. Built on **Native PowerPoint 3D Models (`.glb`)**, **Morph Camera Rotations**, and **Macro Dolly Zooms**, adaptable to any domain without complex prompts.

---

## 1. Native 3D Model Architecture (`.glb` / `<am3d:model3d>`)

PowerPoint features a native 3D rendering pipeline supporting industry-standard `.glb` and `.gltf` 3D assets:

```xml
<am3d:model3d r:embed="rId8">
  <am3d:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="7985021" cy="7623432"/></a:xfrm>
  </am3d:spPr>
  <am3d:camera>
    <am3d:pos x="0" y="0" z="56149260"/>
    <am3d:lookAt x="0" y="0" z="0"/>
    <am3d:perspective fov="2700000"/>
  </am3d:camera>
  <am3d:trans>
    <am3d:scale><am3d:sx n="1000000" d="1000000"/><am3d:sy n="1000000" d="1000000"/><am3d:sz n="1000000" d="1000000"/></am3d:scale>
    <am3d:rot rx="1500000" ry="-3000000" rz="0"/>
  </am3d:trans>
</am3d:model3d>
```

### 3D Morph Camera Interpolation
When the same 3D model appears on consecutive slides with different camera angles or rotations:
$$\mathbf{R}_1(\theta_{x1}, \theta_{y1}, \theta_{z1}) \xrightarrow{\text{Morph 60 FPS}} \mathbf{R}_2(\theta_{x2}, \theta_{y2}, \theta_{z2})$$
PowerPoint's hardware rasterizer smoothly rotates and scales the 3D model in real time without video renders.

### Python COM Insertion
```python
# Programmatically insert native 3D GLB model via PowerPoint COM
model = slide.Shapes.Add3DModel("assets/earth_model.glb", Left=100, Top=50, Width=500, Height=500)
model.Model3DFormat.RotationX = 15
model.Model3DFormat.RotationY = -45
model.Model3DFormat.RotationZ = 10
```

---

## 2. Shared camera direction: Macro 3D Dolly Zoom ($100\% \to 400\%$)

Demonstrated in top-tier science decks (e.g. *Earth & Space*):

```
+-------------------------------------------------------------------------------+
| STAGE 1: WIDE ESTABLISHING ORBIT (Camera Zoom 100%)                           |
|   - 3D Model / Globe centered at normal scale (w: 8.0", h: 7.6").             |
|   - Deep spatial perspective, titles floating in space.                       |
|                                                                               |
| STAGE 2: THE APPROACH & ROTATION (Camera Zoom 180%)                           |
|   - Model rotates by 45° around Y-axis, scaling to (w: 14.8", h: 14.0").      |
|   - Camera draws closer, revealing regional topography.                       |
|                                                                               |
| STAGE 3: TOUCHDOWN MACRO PUSH-IN (Camera Zoom 400%)                           |
|   - Massive camera plunge! Model scales to (w: 41.0", h: 19.2").              |
|   - The horizon fills the entire viewport; audience feels surface arrival!   |
|                                                                               |
| STAGE 4: HORIZON TRACKING PAN (Camera Zoom 220%)                              |
|   - Model slides laterally along the rim (x: -6.3" -> 32.8").                 |
|   - Stages open space for deep-dive technical callout cards.                  |
+-------------------------------------------------------------------------------+
```

### Universal Domain Adaptation
This 4-beat camera move applies to **any subject**:
- **Astronomy & Earth Science:** Deep space $\to$ Planetary orbit $\to$ Atmospheric touchdown.
- **Consumer Hardware / Tech:** Device exterior $\to$ Enclosure opening $\to$ Microchip die macro zoom.
- **Automotive & Engineering:** Aerodynamic chassis $\to$ Hood cutaway $\to$ Turbocharger cylinder.
- **Healthcare & Biotech:** Human anatomy $\to$ Organ tissue $\to$ Cellular DNA double helix.
- **Ecology & Geography:** World map $\to$ National coastline $\to$ Mangrove forest canopy.

---

## 3. The 3D Orbiting Carousel (Team & Feature Showcase)

Arranges cards, team members, or feature modules along a virtual 3D elliptical orbit:

```
                  [ Card 3: Deep Background (Scale 60%, Opacity 40%) ]
                 /                                                     \
   [ Card 2: Left Flank ]                                     [ Card 4: Right Flank ]
      (Scale 85%, Opacity 75%)                                   (Scale 85%, Opacity 75%)
                 \                                                     /
                  [ Card 1: FOREGROUND HERO (Scale 120%, Opacity 100%) ]
```

- **Morph Dynamics:** Advancing the slide rotates all cards around the orbit.
- The hero card steps aside, and the next card rotates into the foreground, smoothly enlarging to $120\%$ scale.

---

## 4. HUD Callouts & Leader Lines

Anchor data cards directly to the 3D model component using vector Target Dots:

```
[ Target Dot ]  (Dual circle: outer glow 0.25", inner dot 0.08")
      |
      |  (Hairline 1.2pt leader line, angled 45° or 90°)
      v
+-------------------------------------------------------------+
| !!HUD_Card (Frosted Glass Card, 3.8" x 1.3")                |
|   TITLE IN BOLD: ATMOSPHERIC BOUNDARY LAYER                 |
|   Desc: High-velocity jet streams drive regional climates.  |
+-------------------------------------------------------------+
```

- **Visual-First Rule:** Keep text under 25 words per slide. Let the 3D model carry 85% of the cognitive load.
Native GLB/COM and OpenXML snippets describe host-level capabilities, not a
fully verified model importer in this release. Test the actual target PowerPoint
version, register media/relationships, and use the canonical validation gate.
Do not claim support solely from the example XML. Carousel coordination belongs
to choreography; this skill supplies model-focused composition.
