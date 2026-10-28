# PMG Canonical API and CLI (v2.1.0)

PMG v2.1.0 provides a unified public API and CLI for universal PowerPoint generation, multi-level validation, and offline execution.

---

## 1. Unified Python API (`PMGRuntime`)

```python
from pmg import PMGRuntime

# Auto-detects runtime capabilities (Mode A: Sandbox, Mode B: Local Desktop, Mode C: Remote Worker)
runtime = PMGRuntime.auto()

# Inspect capabilities programmatically
print(runtime.capabilities.platform)
print(runtime.capabilities.powerpoint_available)
print(runtime.capabilities.can_compile_native_pptx)

# Build a native presentation from a JSON specification or dict
result = runtime.build(
    project="presentation_spec.json",  # or spec dict or builder callable
    output="quantum_architecture.pptx",
    validate=True,  # Runs all available validation levels
)

print(f"Delivered: {result.output_path}")
print(f"Overall Status: {result.validation_report.overall_status}")
print(f"Can Deliver: {result.can_deliver}")
print(f"Summary: {result.summary}")
```

### Presentation Specification Format

```json
{
  "theme": "LIGHT_ACADEMIC_CYAN",
  "slides": [
    {
      "title": "Quantum Entanglement",
      "subtitle": "Cross-Slide 60 FPS Morph Continuity",
      "transition": "morph",
      "duration_ms": 1400,
      "elements": [
        {
          "type": "card",
          "name": "!!HeroCard",
          "left": 1.0,
          "top": 2.5,
          "width": 5.5,
          "height": 3.0,
          "title": "Cryogenic Control",
          "body": "Phase coherence maintained across multi-node topologies."
        }
      ],
      "animations": [
        {
          "target": "!!HeroCard",
          "effect": "fade",
          "trigger": "with_previous",
          "duration_ms": 800,
          "delay_ms": 200
        }
      ]
    }
  ]
}
```

---

## 2. Command-Line Interface (CLI)

### Environment Diagnostics (`pmg doctor` & `pmg capabilities`)
```bash
# Print diagnostic report for active environment
pmg doctor

# Print machine-readable capabilities JSON
pmg capabilities
pmg capabilities --table
```

### Building Presentations (`pmg build`)
```bash
# Compile from JSON specification
pmg build presentation_spec.json -o presentation.pptx

# Compile from built-in preset
pmg build --preset space --slides 10 -o space_presentation.pptx
pmg build --preset spotify --slides 6 -o spotify_app.pptx

# Compile in static mode (skipping live PowerPoint Slide Show playback)
pmg build presentation_spec.json -o presentation.pptx --skip-playback
```

### Three-Level Validation (`pmg validate`)
```bash
# Validate existing presentation across Level 1, Level 2, and Level 3
pmg validate presentation.pptx

# Print structured JSON report
pmg validate presentation.pptx --json

# Validate and apply 3-loop self-healing to an output deck
pmg validate input.pptx -o healed.pptx --max-repairs 3
```

### Video Rendering & Slide Show Recording (Requires PowerPoint Desktop)
```bash
# Export native 60 FPS MP4 video
pmg render presentation.pptx -o powerpoint_export.mp4 --fps 60

# Record actual Slide Show playback with telemetry
pmg record presentation.pptx -o actual_slideshow.mp4 --fps 30

# Run AI/CV visual review comparing export vs recording
pmg review presentation.pptx --export-video powerpoint_export.mp4 --record-video actual_slideshow.mp4 --telemetry recording_telemetry.json
```

---

## 3. Supported Native Animation Effects

| Category | Native Effect Types | Triggers |
| :--- | :--- | :--- |
| **Entrances** | `appear`, `fade` (`fade_in`), `fly_in`, `wipe`, `zoom` (`zoom_in`) | `with_previous`, `after_previous`, `on_click` |
| **Emphasis** | `scale` (`grow_shrink`), `rotate` (`spin`), `transparency`, `pulse` | `with_previous`, `after_previous`, `on_click` |
| **Motion Paths** | `motion_path` (linear, cubic Bezier `M ... C ... E`) | `with_previous`, `after_previous`, `on_click` |
| **Exits** | `disappear`, `fade_out`, `fly_out`, `wipe_out` | `with_previous`, `after_previous`, `on_click` |
| **Transitions** | `morph`, `fade`, `circle`, `split`, `wipe`, `push`, `cover` | Slide transition (`p14:dur`) |
