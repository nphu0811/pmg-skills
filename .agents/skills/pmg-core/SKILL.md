---
name: pmg-core
description: "Plan native PPTX layout, typography, geometry and readable layered composition. Allows intentional overlap; delegates production and animation compilation to the canonical PMG engine."
---

# PMG Core (`pmg-core`)

## Responsibility and dependencies

Layout, typography, readable regions and intentional layered composition. Uses existing layout helpers; does not compile motion XML.

Dependencies: pmg-engine; style helpers; shared scene geometry as released. All generated candidates must pass
PMGOrchestrator before delivery; direct save/COM snippets below are lower-level
composition examples, not independent production entrypoints.

The architectural foundation of **PowerPoint Motion Graphics (PMG)**. It establishes strict spatial geometry, aspect-ratio constraints, high-contrast luminescence tokens, and typographic hierarchy.

---

## 1. Canvas Dimensions & Geometry

All PMG presentations strictly use **16:9 Widescreen (1920x1080 equivalent in PowerPoint OpenXML)**:
- **Slide Width:** `13.333 inches` (`12192000 EMUs`)
- **Slide Height:** `7.5 inches` (`6858000 EMUs`)

### Spatial Zones
```
+-------------------------------------------------------------------------------+
| TOP BADGE / PROGRESS MARGIN (top: 0.4" - 0.9", height: 0.4")                  |
|                                                                               |
| [ MEDIA / 3D FOCUS ZONE ]               [ HUD / NARRATIVE ZONE ]              |
| Left: 0.9" to 6.2"                      Left: 6.6" to 12.4"                   |
| Width: 5.3"                             Width: 5.8"                           |
| Top: 1.2" to 6.6"                       Top: 1.2" to 6.6"                     |
|                                                                               |
| BOTTOM TRACKER MARGIN (top: 6.9" - 7.2", height: 0.3")                        |
+-------------------------------------------------------------------------------+
```

---

## 2. Color Systems & Luminescence Tokens

### Bright Modern / Studio Light (Primary Default)
By default, PMG uses **Bright, Crisp, Modern Light Themes** (pure white `#FFFFFF`, ice white `#F8FAFC`, tinted snow) with razor-sharp dark slate typography (`#0F172A`) and vibrant energy accents. This delivers maximum clarity, academic authority, and corporate sophistication without murky dark screens.

| Token | Hex | Usage |
| :--- | :--- | :--- |
| `BG_PURE_WHITE` | `#FFFFFF` | Primary crisp white background floor |
| `BG_STUDIO_ICE` | `#F8FAFC` | Sleek studio light backdrop |
| `SURFACE_FROSTED` | `#FFFFFF` (85%-92% alpha) | Translucent frosted acrylic card scrim |
| `BORDER_HAIRLINE` | `#0284C7` or `#E2E8F0` | Card border definition (0.75pt - 1.25pt) |
| `TEXT_PRIMARY` | `#0F172A` / `#1E293B` | High-contrast deep slate for headings and key copy |
| `TEXT_SECONDARY` | `#475569` | Slate for descriptive body text |
| `TEXT_MUTED` | `#94A3B8` | Category badges and telemetry labels |

### Curated Palette Catalog
1. **Light Academic Cyan (Default):**
   - Background `#FFFFFF` | Card `#FFFFFF` (92% alpha) | Accent `#0284C7` (Sky Cyan) | Text `#0F172A`
2. **Light Modern Studio:**
   - Background `#F8FAFC` | Accent `#4F46E5` (Indigo) & `#10B981` (Emerald) | Text `#1E293B`
3. **Light Eco Emerald (Sustainability / Environment / Waste Sorting):**
   - Background `#F6FBF7` | Accent `#059669` (Forest Emerald) | Text `#064E3B`
4. **Light Editorial Warm (Literature / Culture / Humanities):**
   - Background `#FAF7F2` (Ivory Book) | Accent `#C2410C` (Terracotta) & `#D97706` (Gold) | Text `#292524`
5. **Dark Obsidian / Cyber Space (Secondary Option):**
   - Background `#060608` | Accent `#1DB954` / `#00F5D4` | Text `#FFFFFF` (For aerospace, late-night keynotes).

---

## 3. Typographic Scale & Rhythm

Use clean, geometric modern sans-serif typefaces (`Montserrat`, `Segoe UI`, `Inter`, `Be Vietnam Pro`):

- **Hero Title:** `32pt – 42pt`, Bold / Black, leading: `1.1`
- **Section / Slide Title:** `22pt – 28pt`, Bold
- **Card Subtitle / Category:** `10pt – 12pt`, SemiBold, All Caps, Letter-spacing `+1.5pt`
- **Body Text:** `11pt – 13pt`, Regular, max 2-3 lines per callout card
- **Technical Badges / Telemetry:** `9pt – 11pt`, SemiBold

---

## 4. Readability and intentional overlap

Protect readable regions with calculated boundaries. Intentionally overlapping
images, masks, cards, glass surfaces and foreground layers are supported. For
separate text blocks, use spacing such as:
$$\text{Top}_{N+1} \ge \text{Top}_N + \text{Height}_N + \text{Gap}$$
Where $\text{Gap} \ge 0.15 \text{ inches}$.

- Text boxes inside cards must be bounded so text cannot overflow outside the card border.
- Cards must never collide with screen borders (minimum margin: `0.75 inches`).
