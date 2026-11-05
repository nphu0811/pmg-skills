---
name: pmg-review-validation
description: "Consume PMG audit defects across Level 1 Structural, Level 2 Visual, and Level 3 Playback validation, classify repair safety, and execute transactional self-healing with automatic rollback within a bounded 3-loop repair budget."
---

# Review, Repair, and Certification Policy (v2.4.0)

Consume shared audit defects from `pmg-qa-audit` across all three validation levels.
Classify critical/major defects as blocking. The canonical authority is `PMGOrchestrator` and
`pmg.self_healing.TransactionalRepairManager`, which owns the bounded retry budget (maximum 3 attempts)
and transactional rollback authority.

## Transactional Self-Healing Pipeline

1. **Preserve Original Artifact:** Never overwrite original presentation without explicit authorization.
2. **Safety Classification:**
   - `SAFE_AUTOMATIC`: Missing ContentTypes (`repair_missing_content_types`), duplicate shape IDs (`repair_duplicate_shape_ids`), and dangling animation targets (`repair_broken_animation_targets`).
   - `REQUIRES_REVIEW`: Morph collision disambiguation, transition duration normalization, timing tree rebuild.
   - `UNSUPPORTED`: Irreparable ZIP corruption or semantic loss; aborts without modifying presentation.
3. **Atomic Rollback:** If re-validation fails following a repair attempt, the working copy is discarded and the presentation is restored to its unmodified state (`rolled_back=True`).
4. **Loop Prevention:** Tracks defect signatures across iterations; aborts immediately if identical defects repeat without progress.

## Universal Certification Invariants

- Structural success in Mode A (Sandbox without PowerPoint) yields `PASS (Structural)`, with Level 3 accurately marked `NOT_TESTED — PowerPoint Desktop unavailable`.
- Real playback certification in Mode B (Windows with PowerPoint) or Mode C (Remote Worker) verifies actual slideshow execution and zero repair warnings (`ppAlertsAll=2`, `OpenAndRepair=False`).
- Never report Level 3 PASS without actual PowerPoint Desktop execution.
- Deliver ONLY the final validated `.pptx` file to ordinary users. Keep internal QA videos and reports in `internal_qa/`.
- Every certified deliverable is bound to its exact SHA-256 hash. Any byte alteration invalidates the certificate.

```bash
# Validate and heal deck within 3 loops
pmg validate input.pptx -o final.pptx --max-repairs 3

# Transactional safe repair
pmg repair input.pptx --safe -o final.pptx
```
