# Canonical API and CLI

Install the engine in the interpreter used by the agent. Copying SKILL.md files
does not install Python or dependencies. For a published commit:

```text
python -m pip install "git+https://github.com/nphu0811/pmg-skills.git@<verified-commit-sha>"
python -m pmg --version
python -m pmg worker doctor
```

From a repository checkout use `python -m pip install -e .`. Do not depend on the
agent's current directory accidentally exposing an uninstalled `pmg` module.

```python
from pmg import PMGOrchestrator, build_environmental_protection_deck, ValidationStatus

manifest = PMGOrchestrator().run_pipeline(
    output_pptx_path="Earth.pptx",
    builder_fn=lambda path: build_environmental_protection_deck(path, num_slides=4),
)
if manifest.status == ValidationStatus.PASS:
    print(manifest.delivered_pptx)
```

`run_pipeline` retains builder_fn/existing_pptx_path, hold_duration_ms,
max_repair_loops, require_video_playback, worker_mode/url/token and run_id.
max_repair_loops accepts 0–3, including compatible fallback within attempt 3.
allow_substitution controls that fallback. Static mode may repair technical
defects but returns NOT_VERIFIED and never sets delivered_pptx.

`StaticValidator.validate(path)` returns a read-only structured audit with
is_valid, status, defects and existing slide/anchor/animation summaries.
`verify_com_open=True` is retained as a read-only compatibility option; skipped
PowerPoint checks cannot report clean. `ReviewDefect` is shared by inspection and
repair. `execute_repair_loop` applies one attempt, never starts an independent loop.

`SlideAnimationTimeline.add_effect` compiles through the existing native engine.
Aliases: fade_in→fade, zoom_in→zoom, spin→rotate, grow_shrink→scale,
linear_motion_path/bezier_motion_path→motion_path. Unknown effects raise
ValueError, including directly created or subsequently mutated AnimationStep
records. Targets must exist on the compiled slide, including grouped shapes.

```text
python -m pmg build -o Space.pptx --preset space --slides 20 --worker local
python -m pmg validate Input.pptx -o Final.pptx --worker auto
python -m pmg inspect Input.pptx --timeline-out expected_timeline.json
python -m pmg render Input.pptx -o internal_qa/powerpoint_export.mp4 --fps 60
python -m pmg record Input.pptx -o internal_qa/actual_slideshow_recording.mp4 --fps 30
python -m pmg review Input.pptx --export-video internal_qa/powerpoint_export.mp4 --record-video internal_qa/actual_slideshow_recording.mp4 --telemetry internal_qa/recording_telemetry.json
python -m pmg reconstruct Reference.pptx -o Reconstructed.pptx
python -m pmg reconstruct Reference.mp4 -o Reconstructed.pptx
python -m pmg benchmark run --all --worker static
```

`record` writes actual recording_telemetry.json beside the video. `review` requires
that evidence; it never invents slide events from the expected timeline. Review
does not publish a deliverable: `validate` runs the complete orchestrator gate.

Reconstruction and direct builder APIs remain available for engineering and
composition. Their saved outputs are candidates until passed to run_pipeline.
Existing SpotifyAppShellBuilder, EarthSpaceMotionBuilder, UniversityBloomBuilder,
MotionPresetRegistry, PPTXMotionReverseEngineer and VideoMotionReconstructor remain
available. Future graph/asset/director APIs are specified in the roadmap, not
advertised as implemented commands.
