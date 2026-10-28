---
name: pmg-engine
description: "Create or validate animated PPTX through the canonical PMG engine. Routes professional presentation, logo-to-PPTX and reference-animation requests to installed capabilities, real PowerPoint playback and the final delivery gate."
---

# PMG canonical entrypoint

Use `python -m pmg` or `pmg.orchestrator.PMGOrchestrator` for PPTX production and
delivery. Builders create native editable candidates; only the orchestrator can
certify them. Do not flatten scenes into screenshots or background videos.

## Runtime prerequisite for freshly installed skills

Skills contain instructions; the skill installer does not install the Python
engine. Run `python -m pmg --version` in the actual execution environment. If
unavailable, install the engine from the same verified repository commit:

```text
python -m pip install "git+https://github.com/nphu0811/pmg-skills.git@<verified-commit-sha>"
python -m pmg --version
python -m pmg worker doctor
```

From the repository use `python -m pip install -e .`. Verify provenance against
the selected published commit. Do not assume a copied skill includes `pmg`, or
that other skills load automatically on every agent. Explicitly load the relevant
installed instructions; use existing module capabilities even without a dedicated
skill. See the checkout's docs/api-and-cli.md and docs/development-roadmap.md.

## Routing and boundaries

- Animated professional presentation: use core/style guidance and choreography;
  reuse cinematic_builder presets as configurable building blocks.
- Morph continuity: load pmg-morph for technical identity/compilation; choreography
  decides motion choice and timing, and invokes the lower-level engine.
- Camera/depth motion: load pmg-camera-parallax. Model-focused 3D composition loads
  pmg-3d-cinematic and reuses that camera system.
- Recreate animation from PPTX/video: reuse pmg.reverse_engineer or
  pmg.video_reconstructor, then submit the reconstructed candidate to the gate.
- Logo-to-PPTX, asset research or decomposition: use actual available asset tools;
  do not claim future capabilities. pmg-visual-assets/pmg-image-decomposer are
  scheduled for v1.6.0 and pmg-storyboard/shared graphs for v1.7.0. Load them only
  after they exist in the installed release. Do not fabricate providers or layers.
- QA detects defects read-only; review chooses repairs. The orchestrator owns one
  maximum-three-attempt budget and the final certificate. Worker is infrastructure.

## Mandatory production gate

Build → package OPC-compliant metadata → read-only static/animation audit →
PowerPoint clean-open preflight → native export → actual slideshow recording →
visual/timeline/object-visibility review → bounded repair if needed → repackage
and repeat all checks → final clean reopen/hash verification → delivery.

The embedded JSON part must declare application/json and records provenance with
NOT_VERIFIED; the external companion records the final status and exact SHA-256.
Never change a PPTX after certification without full verification again. Use
`ppAlertsAll = 2` (1 suppresses alerts), and never request OpenAndRepair. Missing
PowerPoint/worker or skipped checks mean BLOCKED/NOT_VERIFIED, never PASS. Request
60 FPS export, measure actual export/capture rates, and report those honestly.

```text
python -m pmg build -o Space.pptx --preset space --slides 20 --worker auto
python -m pmg validate Input.pptx -o Final.pptx --worker auto
python -m pmg inspect Input.pptx --timeline-out expected_timeline.json
```

Ordinary presentation users receive only the validated final PPTX. Keep videos,
keyframes and JSON evidence in internal_qa unless explicitly requested. Static
mode and direct builder saves are engineering candidates, not certified delivery.
