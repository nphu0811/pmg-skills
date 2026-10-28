# PMG Skills Cross-Platform Compatibility Matrix (v2.1.0)

This matrix details the capabilities, supported operations, and validation levels across different host execution environments.

---

## Compatibility Matrix

| Feature / Capability | Web Sandbox (Claude Web / ChatGPT) | Linux CI (Headless Docker) | macOS (Local Desktop) | Windows (without PowerPoint) | Windows (with Microsoft PowerPoint Desktop) | Remote Windows Worker (Mode C) |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Execution Mode** | `web_sandbox` (Mode A) | `web_sandbox` (Mode A) | `local_desktop` (Mode B) | `local_desktop` (Mode B) | `local_desktop` (Mode B) | `remote_worker` (Mode C) |
| **Native PPTX Compilation** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** |
| **Morph Transitions (`<p159:morph>`)** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** |
| **Traveling Entities (`!!` Anchors)** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** |
| **Animation Pane Timing Trees (`<p:timing>`)**| **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** |
| **Frosted Glassmorphism (`a:alpha`)** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** |
| **Level 1 Structural Validation** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** | **SUPPORTED** |
| **Level 2 Visual Validation** | `SKIPPED` (No renderer) | `SUPPORTED` (via LibreOffice) | `SKIPPED` / LibreOffice | `SKIPPED` / LibreOffice | **SUPPORTED** (PowerPoint COM) | **SUPPORTED** (Remote Worker) |
| **Level 3 Playback Validation** | `NOT_TESTED` (No PowerPoint) | `NOT_TESTED` (No PowerPoint) | `NOT_TESTED` / AppleScript | `NOT_TESTED` (No PowerPoint) | **SUPPORTED** (Live COM SlideShow) | **SUPPORTED** (Remote COM SlideShow) |
| **Deliverable Certification** | `PASS (Structural)` | `PASS (Structural)` | `PASS (Structural)` | `PASS (Structural)` | `PASS (Playback Verified)` | `PASS (Playback Verified)` |

---

## Detailed Environment Profiles & Examples

### 1. Claude Web Sandbox (Mode A)
- **Environment Profile:** Isolated Linux container, no internet access for `pip install`, no Microsoft PowerPoint installed.
- **Workflow:**
  ```python
  from pmg import PMGRuntime

  runtime = PMGRuntime.auto()
  result = runtime.build(
      project="presentation_spec.json",
      output="quantum_computing.pptx",
      validate=True,  # Executes Level 1 Structural Validation
  )
  print(result.summary)
  # Output: Built native presentation: quantum_computing.pptx. PASS (Structural): Native OpenXML valid. Level 3 Playback: NOT_TESTED - PowerPoint Desktop unavailable.
  ```
- **Limitations:** Slideshow recording and video export cannot run; Level 3 is reported as `NOT_TESTED`. The generated `.pptx` file is 100% native and will animate with full 60 FPS Morph when opened in PowerPoint by the user.

### 2. ChatGPT Code-Execution Environment (Mode A)
- **Environment Profile:** Ephemeral Jupyter/Python sandbox, pre-installed Python packages, no COM, no PowerPoint Desktop.
- **Workflow:**
  ```python
  from pmg import PMGRuntime

  runtime = PMGRuntime.auto()
  result = runtime.build(
      project={
          "theme": "LIGHT_ACADEMIC_CYAN",
          "slides": [
              {
                  "title": "Machine Learning Fundamentals",
                  "elements": [{"type": "card", "name": "!!ModelCard", "title": "Neural Networks", "left": 1.0, "top": 2.5, "width": 5.0, "height": 3.0}],
                  "animations": [{"target": "!!ModelCard", "effect": "fade", "duration_ms": 800}]
              }
          ]
      },
      output="ml_fundamentals.pptx"
  )
  ```
- **Limitations:** Static structural checks pass; Level 3 playback is reported as `NOT_TESTED`.

### 3. Linux CI (Headless Docker / GitHub Actions) (Mode A)
- **Environment Profile:** Linux x86_64 or arm64 runner, optional headless LibreOffice (`soffice`), no PowerPoint Desktop.
- **Workflow:**
  ```bash
  # Run structural validation
  pmg validate candidate.pptx
  ```
- **Behavior:** Level 1 passes. If LibreOffice is installed, Level 2 produces slide preview PDFs/images. Level 3 is reported as `NOT_TESTED`.

### 4. Claude Code / Codex on Windows (without PowerPoint) (Mode B)
- **Environment Profile:** Local Windows workstation without Microsoft Office installed.
- **Workflow:**
  ```bash
  pmg doctor
  pmg build project.json -o Output.pptx
  ```
- **Behavior:** Compiles all native shapes, text, transitions, and animations. Level 1 passes. Level 3 reports `NOT_TESTED — PowerPoint Desktop unavailable`.

### 5. Windows Workstation with Microsoft PowerPoint Desktop (Mode B)
- **Environment Profile:** Windows 10/11, Microsoft 365 or PowerPoint 2019/2021 Desktop, pywin32 installed.
- **Workflow:**
  ```bash
  pmg doctor
  pmg build project.json -o Output.pptx
  pmg validate Output.pptx
  ```
- **Behavior:** Full 3-Level validation:
  * Level 1: OpenXML structural pass.
  * Level 2: Slide previews exported via PowerPoint COM.
  * Level 3: Clean COM opening preflight (`ppAlertsAll=2`, no repair dialogs), Slide Show playback recording, and visual CV timeline review. Status = `PASS`.

### 6. Remote Worker Architecture (Mode C)
- **Environment Profile:** AI agent in sandbox or macOS connecting to a dedicated Windows PowerPoint Validation Worker via HTTP/HTTPS.
- **Workflow:**
  ```bash
  export PMG_QA_WORKER_URL="https://worker.internal:8765"
  export PMG_QA_WORKER_TOKEN="secret_token"
  pmg validate Candidate.pptx --worker remote
  ```
- **Behavior:** Candidate PPTX is securely transmitted to worker; worker executes real PowerPoint Slide Show playback, collects telemetry, and returns structured `WorkerJobResult`.

