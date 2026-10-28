# PMG Skills — PowerPoint Motion Graphics

PMG creates editable native PowerPoint presentations through one Python engine.
It combines Morph, native Animation Pane effects, layered composition, virtual
camera movement and reusable cinematic patterns with actual PowerPoint playback
review and bounded technical repair.

Current implementation: **v1.5.2**. Publication status is recorded in
[GitHub releases](https://github.com/nphu0811/pmg-skills/releases).
The [continued roadmap](docs/development-roadmap.md) preserves the required
v1.5.2 → v1.6.0 → v1.7.0 → v1.8.0 → v2.0.1 sequence and release gates.
Future assets, graphs and Creative Motion Intelligence are not yet advertised
as implemented features. Existing showcase work is retained.

## Install

```text
npx skills add https://github.com/nphu0811/pmg-skills
```

Skill installation provides instructions. Install the canonical Python runtime
separately in the interpreter that will run it:

```text
python -m pip install "git+https://github.com/nphu0811/pmg-skills.git@<verified-commit-sha>"
python -m pmg --version
python -m pmg worker doctor
```

From a checkout use `python -m pip install -e .`. Pin the same published commit
for skills and runtime when validating a release. Do not rely on the current
working directory or unsupported automatic discovery to supply dependencies.

The skill files follow the Agent Skills format. Explicitly load the relevant
installed guidance when the host does not discover dependencies. Copilot users
can use [.github/copilot-instructions.md](.github/copilot-instructions.md).
The CLI/API remain usable independently of any particular agent host.

## Production

```text
python -m pmg build -o Space.pptx --preset space --slides 20 --worker auto
python -m pmg validate Input.pptx -o Final.pptx --worker auto
python -m pmg inspect Input.pptx --timeline-out expected_timeline.json
```

```python
from pmg import PMGOrchestrator, build_environmental_protection_deck, ValidationStatus

manifest = PMGOrchestrator().run_pipeline(
    output_pptx_path="Earth.pptx",
    builder_fn=lambda path: build_environmental_protection_deck(path, num_slides=4),
)
if manifest.status == ValidationStatus.PASS:
    print(manifest.delivered_pptx)
```

The engine packages metadata before static/animation checks, native export,
actual slideshow recording and visual/timeline/visibility review. Repairs use one
maximum-three-attempt budget, with compatible fallback inside attempt three.
Every repair requires fresh validation and playback. Final clean PowerPoint
reopen and SHA-256 checks certify the exact deliverable bytes; no PPTX writes
follow certification.

QA is read-only and reports defects. Review chooses repairs. The orchestrator is
the only final delivery authority. An unavailable PowerPoint/Windows worker is
BLOCKED; static success is NOT_VERIFIED. Export requests 60 FPS and capture
requests 30 FPS by default; actual rates are measured, not assumed.

Ordinary presentation users receive only the final validated PPTX. Videos,
frames and JSON reports stay in internal_qa unless explicitly requested.
The external companion manifest contains the final status and hash; embedded
provenance remains NOT_VERIFIED and points to that external certificate.

## Skills and ownership

| Skill | Primary responsibility |
| --- | --- |
| [pmg-engine](skills/pmg-engine/SKILL.md) | Canonical production routing and delivery entrypoint |
| [pmg-core](skills/pmg-core/SKILL.md) | Geometry, typography, readable layered composition |
| [pmg-morph](skills/pmg-morph/SKILL.md) | Technical Morph compilation, matching and persistent identities |
| [pmg-choreography](skills/pmg-choreography/SKILL.md) | Motion direction, selection, coordination, sequence and timing |
| [pmg-camera-parallax](skills/pmg-camera-parallax/SKILL.md) | Shared camera transforms, depth, pan/zoom and continuity |
| [pmg-3d-cinematic](skills/pmg-3d-cinematic/SKILL.md) | Model preparation/composition, cutaways and annotations; uses shared camera |
| [pmg-glassmorphism](skills/pmg-glassmorphism/SKILL.md) | Acrylic surfaces, alpha, borders and contrast |
| [pmg-qa-audit](skills/pmg-qa-audit/SKILL.md) | Read-only structured defect detection |
| [pmg-review-validation](skills/pmg-review-validation/SKILL.md) | Repair policy and certification through the orchestrator |

See [ownership and boundaries](docs/architecture.md) and the
[skills index](skills/llms.txt). The three future skills retain distinct roles:
visual-assets researches/prepares media; image-decomposer produces layers and
masks; storyboard plans narrative and scenes. They do not duplicate compilers.

## Existing engine capabilities

- `pmg.animations`: native entrances/exits, rotation, scaling, opacity, pulse and
  linear/Bezier paths, triggers, repeat and parallel/sequential timing. Unknown
  effects are rejected; documented aliases are normalized.
- `pmg.cinematic_builder`: reusable configurable presets, infographics,
  parallax/zoom sequences, carousel and interface composition.
- `pmg.cinematic_showcase`: Spotify app shell, Earth/Space reticle and university
  bloom builders; preserved from the existing working tree.
- `pmg.reverse_engineer`: PPTX object/geometry/transition/animation extraction,
  reconstruction and style transfer.
- `pmg.video_reconstructor`: decoding, scene detection, tracking, motion
  estimation and editable approximations. Video does not prove source editability.
- `pmg.worker`: authenticated Windows playback infrastructure for local/remote
  environments; no independent repair loop or certificate.
- `benchmarks/`: the existing 24-case suite, extended in place by future releases.
- Existing `MotionDeck`, `LightThemeCraftDeck`, `ForgeDeck`, `Cinematic3DDeck`,
  themes and low-level helpers remain available. Direct saves are candidates.

Native GLB capabilities depend on target PowerPoint support; skill XML/COM
examples alone do not establish a tested model importer. Intentional compositing
is allowed; readable content must remain clear.

## Engineering references

- [API and CLI](docs/api-and-cli.md)
- [Repository guidance](AGENTS.md)
- [Changelog](CHANGELOG.md)
- [Benchmark suite](benchmarks/README.md)
- [Real API example](examples/validate_native_deck.py)

Run `python -m pytest -q` for regression tests. Some tests execute actual
Windows PowerPoint and require an interactive desktop and FFmpeg. Static checks
must never be represented as real playback verification.

Historical demonstration: [native slideshow recording](demo/pmg_presentation_demo.mp4).
It does not certify the current development revision.

## License

MIT — see [LICENSE](LICENSE).
