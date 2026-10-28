---
name: pmg-review-validation
description: "Consume PMG audit defects, choose repair strategies and request fresh validation through PMGOrchestrator. Owns repair policy and certification decisions using one maximum-three-attempt technical repair budget."
---

# Review, repair and certification policy

Consume the shared read-only audit/playback reports from pmg-qa-audit. Classify
critical/major defects as blocking; decide supported remediation and request fresh
checks. The canonical authority is `PMGOrchestrator.run_pipeline`, which owns the
only retry budget and final PASS/FAIL/BLOCKED/NOT_VERIFIED delivery decision.

`pmg.self_healing.execute_repair_loop` implements one repair attempt. It never
orchestrates another loop. `scripts.motion_validator.review_and_heal_deck` is a
backward compatible orchestrator adapter, not another validator or certificate.

1. Targeted repair: transitions, known manifest MIME registration and native
   visibility/timing/target defects.
2. Structural repair: bridge continuous !! anchors and rebuild supported native
   timing through the existing `SlideAnimationTimeline` compiler.
3. Compatible fallback within the same budget: substitute supported native
   effects while retaining editable content. Do not start extra attempts after 3.

After each change, repackage, audit, export, record and review again. A final
failure blocks delivery. Missing playback infrastructure is BLOCKED; static-only
success is NOT_VERIFIED. Do not use static acceptance or a high aggregate score to
hide a critical defect.

Metadata packaging precedes playback; final static/animation/compatibility,
clean-open and hash checks apply to those exact bytes. PowerPoint alerts must be
enabled (`ppAlertsAll=2`), OpenAndRepair disabled. The final external certificate
contains the SHA-256; embedded provenance points to it and remains NOT_VERIFIED.
Never modify a certified PPTX. Keep engineering evidence internal and deliver only
the validated final PPTX to ordinary presentation users.

```text
python -m pmg validate Input.pptx -o Final.pptx --worker auto --max-repairs 3
```

Creative planning/revision in future Motion Director releases is separate from
technical repair. It must not reset or compete with this repair budget.
