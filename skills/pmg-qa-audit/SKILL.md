---
name: pmg-qa-audit
description: "Inspect PPTX integrity, OpenXML geometry, Morph anchors, native timing trees, assets, compatibility, playback evidence, and issue tamper-evident validation certificates across Level 1 Structural, Level 2 Visual, and Level 3 Playback validation. Produces shared structured reports; does not mutate presentation."
---

# Read-only QA Audit & Immutable Certification (v2.4.0)

Own inspection and defect detection. Use `pmg.validation.validate_presentation`,
`pmg.validator.StaticValidator`, `pmg.diagnostics.DiagnosticClassifier`, and shared `ComprehensiveValidationReport`.
Do not create a parallel validator, repair objects or start an unauthorized repair loop.

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

## Standardized Diagnostics & Immutable Certification

- **Error Classification:** Classify detected defects into standard `ErrorCode` categories:
  `PKG_INVALID_ZIP`, `XML_PARSE_FAILURE`, `REL_BROKEN_TARGET`, `OPC_MISSING_CONTENT_TYPE`, `SHP_DUPLICATE_ID`, `ANIM_MISSING_TARGET`, `TIMING_INVALID_REFERENCE`, `MORPH_IDENTITY_COLLISION`, `PPT_OPEN_FAILURE`.
- **Immutable Certification:** Certified presentations receive an `ImmutableValidationCertificate` cryptographically bound to the artifact's exact SHA-256 hash. Any byte-level modification invalidates the certificate.

```bash
# Validate existing presentation across available levels
pmg validate deck.pptx [--json]
pmg inspect deck.pptx --timeline-out expected_timeline.json
pmg qa deck.pptx [--remote]
```

```python
from pmg.validation import validate_presentation
from pmg.self_healing import issue_validation_certificate, verify_certificate

report = validate_presentation("deck.pptx")
cert = issue_validation_certificate("deck.pptx", validation_level="Level 1 Structural", overall_status=report.overall_status.value)
is_valid, msg = verify_certificate(cert, "deck.pptx")
print(f"Certified: {cert.certificate_id}, Valid={is_valid}")
```
