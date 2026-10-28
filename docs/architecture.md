# Shared engine responsibilities

The orchestrator is the single production and final delivery authority.
Skill loading is agent guidance, not a runtime dependency injector. A host must
load the relevant installed skill instructions explicitly and install the Python
runtime before executing the documented CLI or API.

| Skill | Primary responsibility | Dependencies and boundary |
| --- | --- | --- |
| pmg-engine | Route production through PMGOrchestrator | Shared builders/compilers, audit/review, worker; only entrypoint to delivery |
| pmg-core | Canvas, geometry, typography, readability | Shared layout helpers; allow intentional layering; no animation XML |
| pmg-morph | Cross-slide Morph XML, persistent !! identities, matching and parameters | Existing transition helpers; no narrative planning |
| pmg-choreography | Motion choice, coordination, narrative sequence and timing | Calls Morph/camera/pmg.animations and reusable cinematic presets; no duplicate compiler |
| pmg-glassmorphism | Acrylic surfaces, alpha, borders and readable compositing | Core geometry; permits independent moving layers |
| pmg-camera-parallax | Camera/depth transforms, paths, pan, zoom and continuity | Shared geometry and motion compilers; 3D must reuse camera math |
| pmg-3d-cinematic | Model assets/orientation/placement, cutaways, annotations and compatibility | Shared camera, Morph, core and asset interfaces; no separate camera |
| pmg-qa-audit | Read-only detection of structural, visual and temporal defects | StaticValidator, VisualAnimationReviewer and playback evidence; never repairs/certifies |
| pmg-review-validation | Consume defects, choose supported repairs and request fresh checks | self_healing single-step operations; orchestrator alone owns the maximum-three-attempt loop and certificate |

The v1.5.2 audit/report model is pmg.models.AuditReport, ReviewDefect and
ReviewReport, with existing dictionary interfaces preserved. Static validity is
NOT_VERIFIED for playback. Compatibility inspection belongs to the same validator;
repairs belong to self_healing. Legacy scripts are adapters, not parallel engines.

pmg.worker supplies read-only PowerPoint execution evidence and has no repair
loop or delivery authority. It binds observations to the exact uploaded SHA-256.
The orchestrator rejects unsupported/missing evidence from older workers. Native
effect counts must survive opening in PowerPoint; a changed in-memory document,
repair/protected-view interruption, skipped check or timeout cannot certify it.

Manifest packaging precedes the complete validation pass. The embedded manifest
records provenance and NOT_VERIFIED, pointing to external certification; the
external companion records final status and the exact file hash. A final hash
cannot be embedded in the same bytes it hashes. A failed gate preserves a prior
output. Nothing writes the PPTX after final clean reopen and certification.

Use `ppAlertsAll = 2`, with `OpenAndRepair=False`. The older value 1 suppresses
alerts. See Microsoft's [PpAlertLevel](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppalertlevel)
and [Open2007](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open2007)
references. The worker probe observes modal interruptions in a bounded subprocess.

Morph/camera consolidation and the shared graph schemas are staged in v1.7.0;
this document does not claim those refactors are already implemented. Existing
presets, animation compiler, PPTX/video reconstruction and benchmarks are reused.
The [development roadmap](development-roadmap.md) records the future three skill
boundaries, schema contracts and all release gates through v2.0.1.
