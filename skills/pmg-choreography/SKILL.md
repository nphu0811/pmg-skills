---
name: pmg-choreography
description: "Choreography for cinematic openings, Spotify-style UI shell unlocks, 3D orbiting carousels, native PowerPoint transitions (Morph, Circle Wipe, Split Curtain, Fade), and thematic traveling motion."
---

# PMG Choreography & Transition Suite (`pmg-choreography`)

Motion design in PowerPoint transcends static slide-flipping. By combining **Hardware-Accelerated Morph**, **UI App Shell Unlocks**, and **Focal Transitions**, slides feel like an interactive software application or high-end motion graphics film.

---

## 1. Iconic Choreography Patterns

### Pattern A: The Spotify UI Shell & Radial Morph Unlock (from `MAC-LENIN.pptx`)
Used in product launches, tech keynotes, music, and SaaS dashboards:

```
+-------------------------------------------------------------------------------+
| SLIDE 1: THE BRAND MONOLITH                                                   |
|   - Giant radial ambient orb in background (Oval: w: 19.6", h: 12.0").        |
|   - Hero Brand Emblem / Icon centered (w: 1.67", h: 1.66").                   |
|   - Text (Title, Subtitle, Badges) staged OFF-SCREEN (x >= 14.5" & x <= -2.5").|
|                                                                               |
| SLIDE 2: CONVERGENCE & TITLE LOCKUP (Morph 1.6s)                              |
|   - Ambient orb contracts & docks into corner (w: 6.2").                      |
|   - Title & Author credits swoop in laterally from off-screen margin.         |
|   - Floating pill badges glide into place.                                    |
|                                                                               |
| SLIDE 3: THE APP SHELL / DASHBOARD UNLOCK (Morph 1.6s)                        |
|   - Morph unlocks the full interactive Application Frame:                     |
|     * Left Sidebar Navigation Bar (w: 3.3", h: 7.5").                         |
|     * Top Navigation Header (Search, Controls, User Profile Pill & Avatar).   |
|     * Main Viewport Stage ready for presentation content.                     |
|                                                                               |
| SLIDE 4+: IN-FRAME VIEWPORT STREAM (Morph 1.4s)                               |
|   - The App Shell remains pinned in place.                                    |
|   - Inner content cards smoothly stream and scroll inside the viewport.       |
+-------------------------------------------------------------------------------+
```

### Pattern B: The 3D Orbiting Carousel (from `EARTH AND SPACE.pptx`)
Used for team rosters, feature comparisons, and multi-product portfolios:
- Elements are positioned along a 3D elliptical orbit in space ($Z$-depth):
  - **Foreground Hero:** $120\%$ scale, $100\%$ luminescence.
  - **Flank Items:** $85\%$ scale, $75\%$ luminescence.
  - **Deep Background:** $60\%$ scale, $40\%$ opacity.
- Advancing slides with Morph revolves all elements around the orbit, bringing the next item seamlessly into the center spotlight.

---

## 2. The Native Transition Palette

Select transitions strictly based on narrative psychology:

| Transition | OpenXML Snippet | When to Use | Emotional Impact |
| :--- | :--- | :--- | :--- |
| **Morph** | `<p159:morph option="byObject"/>` | 3D rotations, camera dolly zoom, UI shell unlocks, traveling hero flight | Seamless, fluid, 60 FPS continuity |
| **Circle Wipe** | `<p:circle/>` | Climax spotlight, iris focus, grand ending punch | Focused, memorable resolution |
| **Split Curtain** | `<p:split orient="vert"/>` | Title reveal, keynote stage opening | Grand theatrical unveiling |
| **Fade** | `<p:fade/>` | Tone shift, moving from problem to solution | Reflective, soft dissolve |
| **Push** | `<p:push dir="u"/>` | Vertical timeline advancement, feed scrolling | Dynamic stream |
---

## 3. Dynamic Thematic Traveling Entity ("Vật Thể Bay Xuyên Slide")

A persistent traveling entity must **match the specific topic** of the presentation:

| Topic / Domain | Thematic Traveling Entity |
| :--- | :--- |
| **Space & Astronomy** | 3D Rotating Planet, Satellite Probe, Rocket, Astronaut |
| **Music & Entertainment** | Glowing Vinyl Disc, Soundwave Pulse, 3D Floating Earbud |
| **Tech & Software / AI** | Cyber Core, Neural Data Orb, Glowing Cursor / Pointer |
| **Ecology & Environment** | Floating Green Leaf, Pure Water Droplet, Sprout, Bird |
| **Business & Finance** | 3D Growth Arrow, Golden Token, Metric Compass Pin |
| **Medical & Science** | DNA Helix, Molecular Sphere, Heartbeat Pulse |
| **Culture & History** | Floating Lantern, Calligraphy Quill, Imperial Seal |

### Coordinates & Motion Arc
- **Entry:** Starts off-screen ($X \le -2.0"$ or $X \ge 13.5"$).
- **Midpoint Traversal:** Glides across center stage ($X: 5.5" - 7.5"$, $Y: 1.5" - 2.5"$) with dynamic tilt (`rot`).
- **Docking / Exit:** Hovers adjacent to callout cards or departs gracefully into the margin.

---

## 4. Python API Implementation

```python
from scripts.motion_graph_kit import MotionDeck, Themes

deck = MotionDeck(theme=Themes.LIGHT_ACADEMIC_CYAN)

# 1. Slide 1: Off-screen Title & Staging
s1 = deck.add_slide(transition="none")
deck.add_thematic_traveling_entity(s1, "assets/satellite.png", left=-2.0, top=3.5, rot_deg=15)

# 2. Slide 2: Morph Convergence
s2 = deck.add_slide(transition="morph", duration_ms=1600)
deck.add_thematic_traveling_entity(s2, "assets/satellite.png", left=6.0, top=1.8, rot_deg=-10)

# 3. Slide 3: Morph Docking with Telemetry
s3 = deck.add_slide(transition="morph", duration_ms=1600)
deck.add_thematic_traveling_entity(s3, "assets/satellite.png", left=10.2, top=4.0, rot_deg=12)

# 4. Slide 4: Circle Spotlight Outro
s4 = deck.add_slide(transition="circle", duration_ms=1600)
deck.save("Output.pptx")
```
