---
name: pmg-qa-audit
description: "Inspect PPTX integrity, OpenXML geometry, Morph anchors, native timing trees, assets, compatibility and playback evidence across Level 1 Structural, Level 2 Visual, and Level 3 Playback validation. Produces shared structured reports; does not mutate presentation."
---

# Read-only QA Audit (v2.1.0)

Own inspection and defect detection. Use `pmg.validation.validate_presentation`,
`pmg.validator.StaticValidator`, and shared `ComprehensiveValidationReport`, `Level1Result`,
`Level2Result`, `Level3Result`, and `ReviewDefect`. Do not create a parallel validator,
repair objects or start an unauthorized repair loop.

## Three-Level QA Audit Framework

1. **Level 1 — Structural Validation (Package & OOXML):**
   - ZIP package integrity, XML parsing, OPC ContentTypes (`[Content_Types].xml`).
   - Relationships (`_rels/.rels`, `ppt/_rels/*`), unique shape and animation IDs.
   - Animation targets (`<p:spTgt spid="...">` matches existing shape).
   - Intra-slide `<p:timing>` trees (`mainSeq`, `bldLst`, `<p:set>` visibility).
   - Morph continuity (`<p159:morph>`, matching `!!` object tags).
   - Missing or orphaned media parts.
2. **Level 2 — Visual Validation (Slide Rendering & Geometry):**
   - Slide rendering via available backend (PowerPoint COM or LibreOffice).
   - Bounding box overflow (elements outside 16:9 13.333" x 7.5" canvas).
   - Obvious clipping, text truncation, and missing fonts.
   - If renderer is missing, reports: `SKIPPED — No slide rendering backend available`.
   - *Static slide rendering does NOT prove animation correctness.*
3. **Level 3 — Native PowerPoint Playback Validation:**
   - Real Microsoft PowerPoint Desktop execution (local COM or remote Windows worker).
   - Clean-open preflight (`ppAlertsAll=2`, `OpenAndRepair=False`).
   - Measured export and Slide Show recording frame rates (requested 60 FPS vs actual).
   - If PowerPoint Desktop is unavailable, reports: `NOT_TESTED — PowerPoint Desktop unavailable`.
   - *Never report Level 3 PASS without actual Microsoft PowerPoint Desktop execution.*

```bash
# Validate existing presentation across available levels
pmg validate Deck.pptx
pmg inspect Deck.pptx --timeline-out expected_timeline.json
```

```python
from pmg.validation import validate_presentation

report = validate_presentation("Deck.pptx")
print("Level 1:", report.level1.status)
print("Level 2:", report.level2.status)
print("Level 3:", report.level3.status)
print("Summary:", report.summary)
```
