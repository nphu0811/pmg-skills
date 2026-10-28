---
name: pmg-choreography
description: "AI Motion Director and choreography suite. Directs narrative motion, coordinates Scene Graph and Motion Graph dependencies, plans cinematic timing, and compiles purposeful animations via native PowerPoint engines."
---

# PMG Choreography & Transition Suite (`pmg-choreography`)

## Responsibility and dependencies

High-level direction: what moves, why and when. Own motion selection, narrative coordination, scene transitions and timing. Reuse cinematic_builder presets and invoke the shared Morph/camera/pmg.animations implementations; never implement duplicate XML compilers.

Dependencies: pmg-engine; pmg-morph; pmg-camera-parallax; pmg.animations; cinematic_builder. All generated candidates must pass
PMGOrchestrator before delivery; direct save/COM snippets below are lower-level
composition examples, not independent production entrypoints.

Motion design in PowerPoint transcends static slide-flipping. By combining **Hardware-Accelerated Morph**, **UI App Shell Unlocks**, **Viewfinder Reticle Targeting**, **Radial Blooming Petals**, and **Focal Transitions**, presentations feel like an interactive software application or high-end motion graphics film.

---

## 1. Iconic Choreography Patterns (from `DataPowerPoint`)

### Pattern A: The Spotify UI Shell & Radial Morph Unlock (from `MAC-LENIN.pptx`)
Used in product launches, tech keynotes, music, SaaS dashboards, and modern portfolios:

```
+-------------------------------------------------------------------------------+
| SLIDE 1: THE BRAND / TOPIC MONOLITH                                           |
|   - Giant radial ambient orb in background (Oval: w: 17.9", h: 10.9").        |
|   - Hero Brand Emblem / Icon centered (w: 1.67", h: 1.66").                   |
|   - Text (Title, Subtitle, Badges) staged OFF-SCREEN (x >= 14.5" & x <= -2.5").|
|                                                                               |
| SLIDE 2: CONVERGENCE & TITLE LOCKUP (Morph 2.0s)                              |
|   - Ambient orb contracts & docks into top-left corner (w: 6.2").             |
|   - Secondary accent orb docks into bottom-right corner.                      |
|   - Title & Subtitle credits swoop in laterally from off-screen margin.       |
|   - Floating pill badges glide into place from opposite edges.                |
|                                                                               |
| SLIDE 3: THE APP SHELL / DASHBOARD UNLOCK (Morph 2.0s)                        |
|   - Morph unlocks the full interactive Application Frame:                     |
|     * Left Sidebar Navigation Bar (w: 3.29", h: 7.50") with macOS 3-dots.    |
|     * Navigation Rail items with active category highlight.                   |
|     * Top Navigation Header (History controls, Upgrade pill, User avatar).    |
|     * Main Viewport Stage ready for presentation content.                     |
|                                                                               |
| SLIDE 4+: IN-FRAME VIEWPORT STREAM (Morph 1.5s - 2.0s)                        |
|   - The App Shell remains permanently pinned in place.                        |
|   - Content cards smoothly stream, transform, and scroll inside viewport.    |
|   - Team member cards or project stats expand across the canvas.             |
+-------------------------------------------------------------------------------+
```

### Pattern B: The Earth & Space Reticle Viewfinder & 4-Quadrant Sequence (from `EARTH AND SPACE.pptx` Slides 1-21)
Used for space exploration, technology architectures, defense/aerospace, and industrial deep dives:
- **Slide 1 (Establishing Shot):** Giant celestial globe enters from the bottom edge (`!!Space_Globe`), with cosmic particle dust nodes.
- **Slide 2 (Planetary Pan):** Globe shifts to the top-left flank as the camera tracks right; moon/satellite enters orbit; targeting reticle corner brackets (`!!Reticle_TL`, `!!Reticle_TR`, `!!Reticle_BL`, `!!Reticle_BR`) wait off-screen (`x=17.0", y=12.0"`).
- **Slide 3 (Viewfinder Target Lock):** Reticle brackets converge and clamp onto the subject (`center_x=9.5", center_y=3.5"`), forming a dynamic targeting square `[ ]`. Title updates to "TARGET COORDINATES LOCKED".
- **Slide 4 (Reticle Expansion to 3 Framed Cards):** The reticle brackets expand outwards, morphing into 3 distinct framed callout cards (`Landmark`, `Innovation`, `Community`).
- **Slide 5 (Centered Focal Transition - "BUT SOMETHING IS MISSING"):** Macro zoom into a central mystery card; high-contrast dramatic question triggers narrative pivot.
- **Slide 6 (4-Quadrant Rotating Segmented Radar Wheel):** A segmented circular radar wheel partitioned into 4 quadrants (`Q1: 0°`, `Q2: 90°`, `Q3: 180°`, `Q4: 270°`). As slides advance, the wheel rotates $90^\circ$ per slide, shifting the active highlight from quadrant to quadrant.

### Pattern C: University Concentric Rings Blooming into 8-Petal Mục Lục (from `PROMPT.pdf`)
Used for academic defenses, thesis presentations, environmental research, and corporate annual reports:
- **Slide 1 (Cover / Concentric Anchor):** Three concentric hairline rings (`r=1.2", 2.2", 3.2"`) frame the cover topic card with institutional headers and student/advisor credits.
- **Slide 2 (Radial Bloom into 8 Petals):** The concentric rings bloom outwards into 8 distinct radial petals (`!!Petal_0` through `!!Petal_7`) distributed at $45^\circ$ angles around a glowing center node.
- **Slide 3 (Two-Column Mục Lục / Table of Contents):** The 8 petals morph into an elegant two-column 8-row agenda list (Items 1-4 on the left column, Items 5-8 on the right column).
- **Slide 4+ (Active Section Deep Dive):** The active section card enlarges while non-active items dock into a compact progress tracker rail.

---

## 2. Universal Topic Adaptability Directive

> [!IMPORTANT]
> **Never force a single rigid layout or topic.**
> PMG choreography patterns are **universal structural primitives**, not single-topic templates.
> - **Do NOT follow a single row of cards** if the content calls for a grid, concentric circles, 4 quadrants, or an app shell.
> - **Theme Flexibility:** Every pattern can be rendered in **Bright Modern Light** (`LIGHT_ACADEMIC_CYAN`, `LIGHT_MODERN_STUDIO`, `LIGHT_ECO_EMERALD`) or **Deep Obsidian** (`SPOTIFY_CYBER`, `COSMOS_DEEP`).
> - Adapt titles, icons, metrics, and card counts dynamically to the user's domain (Healthcare, Agriculture, FinTech, Academic Research, or Art).

---

## 3. The Native Transition Palette

Select transitions strictly based on narrative psychology:

| Transition | Compiler operation | When to Use | Emotional Impact |
| :--- | :--- | :--- | :--- |
| **Morph** | `apply_transition(..., "morph")` | 3D rotations, camera dolly zoom, UI shell unlocks, reticle targeting, radial blooms | Seamless, fluid, 60 FPS continuity |
| **Circle Wipe** | `apply_transition(..., "circle")` | Climax spotlight, iris focus, grand ending punch | Focused, memorable resolution |
| **Split Curtain** | `apply_transition(..., "split")` | Title reveal, keynote stage opening | Grand theatrical unveiling |
| **Fade** | `apply_transition(..., "fade")` | Tone shift, moving from problem to solution | Reflective, soft dissolve |
| **Push** | `apply_transition(..., "push")` | Vertical timeline advancement, feed scrolling | Dynamic stream |

---

## 4. Thematic Traveling Entity ("Vật Thể Bay Xuyên Slide")

A persistent traveling entity must **match the specific topic** of the presentation:

| Topic / Domain | Thematic Traveling Entity |
| :--- | :--- |
| **Space & Astronomy** | 3D Rotating Planet, Satellite Probe, Rocket, Lunar Rover |
| **Music & Entertainment** | Glowing Vinyl Disc, Soundwave Pulse, 3D Floating Earbud |
| **Tech & Software / AI** | Cyber Core, Neural Data Orb, Glowing Cursor / Pointer |
| **Ecology & Environment** | Floating Green Leaf, Pure Water Droplet, Sprout, Golden Bee |
| **Business & Finance** | 3D Growth Arrow, Golden Token, Metric Compass Pin |
| **Medical & Science** | DNA Helix, Molecular Sphere, Heartbeat Pulse |
| **Education & Thesis** | Concentric Ring Node, Academic Quill, Ribbon Pin |

### Coordinates & Motion Arc
- **Entry:** Starts off-screen ($X \le -2.0"$ or $X \ge 13.5"$).
- **Midpoint Traversal:** Glides across center stage ($X: 5.5" - 7.5"$, $Y: 1.5" - 2.5"$) with dynamic tilt (`rot`).
- **Docking / Exit:** Hovers adjacent to callout cards or departs gracefully into the margin.

---

## 5. Python API Usage

```python
from pmg.cinematic_showcase import (
    SpotifyAppShellBuilder,
    EarthSpaceMotionBuilder,
    UniversityBloomBuilder,
)

# 1. Spotify App Shell UI
builder = SpotifyAppShellBuilder(
    topic_title="AI AUDIO WORKBENCH",
    topic_subtitle="REAL-TIME NEURAL SYNTHESIS",
    author_name="Audio Intelligence Lab",
)
builder.build_deck("output_spotify.pptx", num_slides=4)

# 2. Earth & Space Viewfinder / Targeting Reticle Sequence
earth_builder = EarthSpaceMotionBuilder(
    topic_title="DEEP SPACE OBSERVATORY",
    topic_subtitle="EXOPLANET ATMOSPHERIC ANALYSIS",
)
earth_builder.build_deck("output_space.pptx", num_slides=6)

# 3. University Concentric Bloom & 8-Petal Mục Lục
bloom_builder = UniversityBloomBuilder(
    presentation_title="NGHIÊN CỨU PHÂN LOẠI RÁC THẢI TẠI NGUỒN",
    subtitle="ĐỀ TÀI TỐT NGHIỆP THẠC SĨ",
    student_name="Nguyễn Văn A",
)
bloom_builder.build_deck("output_university.pptx")
```
## Autonomous Motion Director 2.0 & Creative Intelligence (v2.0.1)

In v2.0.1, `pmg-choreography` integrates with `pmg.creative_director.CreativeDirector`:
- **Composable Motion Grammar**: Coordinates multi-phase intra-slide choreography, independent simultaneous paths, orbital assembly, scale pulsing, perspective tilt, and synchronized typography.
- **Continuous One-Shot Direction (20–30s Scenes)**: Long continuous single-slide scenes coordinating starfields, celestial rotations, traveling hero trajectories, and camera-follow simulation (`build_v2_0_oneshot_space_scene`).
- **Multi-Candidate Creative Comparison**: Generates 3 distinct concepts from one prompt, scored across 6 quantitative axes (`clarity`, `originality`, `visual_hierarchy`, `cinematic_coherence`, `rhythm_and_pacing`, `native_editability`).
- **Bounded Pre-Compilation Critique**: Evaluates narrative rhythm, transition variety, and camera progression; applies refinements prior to OpenXML compilation.
- **Native Editability & OPC Invariant**: All animations compile to native `<p:timing>` Animation Pane trees with zero repair warnings under PowerPoint Desktop COM.
