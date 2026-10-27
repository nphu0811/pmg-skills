---
applyTo: ["scripts/**/*.py", "skills/pmg-glassmorphism/**", "skills/pmg-core/**"]
---

# Glassmorphism & Visual Aesthetics — Path-specific instructions

When creating frosted glass cards, UI components, or color schemes:

- **Alpha Scrim Values:** Use 90%–94% alpha on acrylic card fills (`<a:alpha val="92000"/>`). Never use 100% solid fills for cards on textured or photo backgrounds.
- **Hairline Glowing Borders:** Add a subtle hairline border (`0.75pt - 1.25pt`) with translucent pure white or accent glow (`#38BDF8` / `#0284C7`) at 40%–60% alpha to give cards tactile physical depth.
- **Modern Light Theme Standard:** Default to light themes with pure white backgrounds (`#FFFFFF`), high-contrast dark text (`#0F172A`), and vibrant jewel-tone accents (`#0284C7`, `#059669`, `#7C3AED`).
- **Zero-Overlap Containers:** Calculate strict bounding boxes for all cards and labels (`Margin >= 0.35"`). No two text boxes or shape cards may collide or overlap.
