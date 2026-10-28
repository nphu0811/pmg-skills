---
name: pmg-qa-audit
description: "Inspect PPTX integrity, geometry, Morph anchors, native timing, assets, compatibility and playback evidence without modifying the presentation. Produces shared structured defects; does not repair or certify delivery."
---

# Read-only QA audit

Own inspection and defect detection. Use `pmg.validator.StaticValidator`,
`VisualAnimationReviewer` and shared `pmg.models.AuditReport`, `ReviewDefect` and
`ReviewReport`. Do not create a parallel validator, repair objects or start a
repair loop. Return findings to pmg-review-validation/PMGOrchestrator.

Inspect package parts, ContentTypes, XML, relationships, duplicate IDs, object
geometry, media, Morph identities, animation targets/timing and compatibility.
Use actual native export/slideshow evidence for temporal/visibility inspection.
As Scene Graph/mask/asset interfaces are released, inspect their integrity through
the same report model, not separate incompatible reports.

Intentional layering, masks, foreground occlusion, glass overlays and text over
imagery are permitted. Report accidental collisions or obscured readable content;
do not restore a global zero-overlap rule. Text-density guidance is contextual,
not proof of correct motion. Inspect actual measured FPS rather than assuming 60.

```text
python -m pmg inspect Deck.pptx --timeline-out expected_timeline.json
python scripts/slide_inspector.py Deck.pptx
```

```python
from pmg import StaticValidator
audit = StaticValidator().validate("Deck.pptx")
print(audit["is_valid"], audit["status"], audit["defects"])
```

`is_valid=True` in a static audit is NOT_VERIFIED for real playback. The legacy
slide_inspector and motion_validator inspection APIs delegate here and cannot
certify delivery. Package compatibility defects are critical; a repair prompt,
missing native effects or absent evidence must reach the orchestrator's gate.
