---
name: pmg-glassmorphism
description: "Authentic frosted glassmorphism architecture for PowerPoint: light-mode acrylic surfaces, OpenXML alpha scrims (a:alpha), glowing hairline borders (0.75pt - 1.25pt), and multi-tier translucent depth."
---

# PMG Glassmorphism (`pmg-glassmorphism`)

Glassmorphism in PMG creates high-end, tactile telemetry surfaces. It allows underlying motion, 3D elements, or photography to subtly reveal depth while maintaining 100% text readability.

---

## 1. Light Mode Frosted Acrylic (Primary Standard)

In light themes, frosted cards provide structural elevation without heavy dark shadows:

```xml
<a:spPr>
  <!-- Frosted Pure White Scrim Base -->
  <a:solidFill>
    <a:srgbClr val="FFFFFF">
      <a:alpha val="92000"/> <!-- 92% Opacity Frosted Acrylic -->
    </a:srgbClr>
  </a:solidFill>
  
  <!-- Subtle Glowing Cyan or Hairline Slate Border (1.0 pt) -->
  <a:ln w="12700" cmpd="s">
    <a:solidFill>
      <a:srgbClr val="0284C7">
        <a:alpha val="55000"/> <!-- 55% Opacity Border Glow -->
      </a:srgbClr>
    </a:solidFill>
  </a:ln>
</a:spPr>
```

### Contrast Invariants for Light Glass
- **Card Background:** `#FFFFFF` at `90% - 94%` alpha.
- **Card Border:** `#0284C7` (Sky Cyan), `#4F46E5` (Indigo), or `#CBD5E1` (Neutral Hairline).
- **Text Inside Card:** Strictly `#0F172A` (Headings) and `#475569` (Body). **Never use white or light gray text on light glass cards.**

---

## 2. Dark Mode Obsidian Glass (Secondary Option)

For aerospace and late-night tech keynotes:

```xml
<a:solidFill>
  <a:srgbClr val="152238">
    <a:alpha val="82000"/> <!-- 82% Opacity Obsidian Scrim -->
  </a:srgbClr>
</a:solidFill>
<a:ln w="12700">
  <a:solidFill>
    <a:srgbClr val="00F5D4">
      <a:alpha val="70000"/>
    </a:srgbClr>
  </a:solidFill>
</a:ln>
```

---

## 3. Multi-Tier Glass Card Architecture

```
+-------------------------------------------------------------+
| PRIMARY FROSTED CARD (width: 5.8", height: 4.8", alpha: 92%)|
|                                                             |
|   [ STATUS PILL: alpha 95%, bg: #0284C7, text: white ]      |
|   KHẢO SÁT THỰC TẾ // BƯỚC 01                               |
|                                                             |
|   Title: NHẬN THỨC VÀ HÀNH VI SINH VIÊN                     |
|   Subtitle: Dữ liệu phân loại rác thải tại HUTECH 2026      |
|                                                             |
|   +-------------------------------------------------------+ |
|   | NESTED METRIC CARD (white alpha: 80%, height: 1.2")   | |
|   | Tỷ lệ nhận thức: 97.7%  | Muốn giải quyết: 100%       | |
|   +-------------------------------------------------------+ |
+-------------------------------------------------------------+
```

### Python Implementation
```python
from scripts.motion_graph_kit import MotionDeck

deck = MotionDeck()
deck.add_glass_card(
    slide,
    left=6.6,
    top=1.3,
    width=5.8,
    height=4.8,
    name="!!hero_glass",
    bg_hex="FFFFFF",
    alpha="92000",
    border_hex="0284C7",
    border_alpha="55000",
    border_width_pt=1.25,
    corner_radius=6000
)
```
