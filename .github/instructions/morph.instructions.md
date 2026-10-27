---
applyTo: ["scripts/**/*.py", "skills/pmg-morph/**", "skills/pmg-choreography/**"]
---

# Morph Transitions & Choreography — Path-specific instructions

When writing or modifying slide transitions, traveling entities, or choreography:

- **OpenXML Morph Injection:** Always inject `<p159:morph option="byObject"/>` under `<p:transition>` in `ppt/slides/slide{N}.xml`. Ensure XML namespace `xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"` is present in the slide root.
- **The `!!` Shape Naming:** Every morphing shape pair must be named `!!<Identifier>` identically on both source and destination slides.
- **Aspect Ratio Locking:** Never distort or stretch traveling entities between slides. Preserve geometry across scaling.
- **Choreographed Entrances/Exits:** Stage entering elements off-screen ($X \le -2.0"$ or $X \ge 13.5"$) on the preceding slide so PowerPoint's Morph engine smoothly slides them into view.
- **Duration Tuning:** Default Morph duration is 1.25s – 1.75s for cinematic feel; quick HUD state changes should run at 0.75s.
