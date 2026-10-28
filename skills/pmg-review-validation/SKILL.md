---
name: pmg-review-validation
description: "Consume PMG audit defects across Level 1 Structural, Level 2 Visual, and Level 3 Playback validation, choose repair strategies, and execute self-healing within a bounded 3-loop repair budget."
---

# Review, Repair, and Certification Policy (v2.1.0)

Consume shared audit defects from `pmg-qa-audit` across all three validation levels.
Classify critical/major defects as blocking. The canonical authority is `PMGOrchestrator.run_pipeline`,
which owns the only retry budget (maximum 3 attempts) and final delivery decision.

1. **Targeted Repair:** Transitions, missing ContentTypes, and native visibility/timing defects.
2. **Structural Repair:** Bridge continuous `!!` Morph anchors and rebuild supported native timing trees through `SlideAnimationTimeline`.
3. **Compatible Fallback:** Substitute supported native effects while retaining 100% editable shapes. Never start extra loops after attempt 3.

## Universal Certification Invariants

- Structural success in Mode A (Sandbox without PowerPoint) yields `PASS (Structural)`, with Level 3 accurately marked `NOT_TESTED — PowerPoint Desktop unavailable`.
- Real playback certification in Mode B (Windows with PowerPoint) or Mode C (Remote Worker) verifies actual slideshow execution and zero repair warnings (`ppAlertsAll=2`, `OpenAndRepair=False`).
- Never report Level 3 PASS without actual PowerPoint Desktop execution.
- Deliver ONLY the final validated `.pptx` file to ordinary users. Keep internal QA videos and reports in `internal_qa/`.

```bash
# Validate and heal deck within 3 loops
pmg validate Input.pptx -o Final.pptx --max-repairs 3
```
