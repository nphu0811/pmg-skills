# PMG Skills Offline Execution & Dependency Guide (v2.1.0)

This guide documents PMG's offline-first architecture, dependency resolution sequence, reproducible packaging instructions, and troubleshooting steps.

---

## 1. Dependency Resolution Priority

PMG enforces a strict 6-tier fallback sequence when resolving runtime dependencies:

```
Tier 1: Existing Installed Python Packages
   ↓ (if missing)
Tier 2: Bundled / Cached Offline Wheels (PMG_OFFLINE_CACHE, ~/.pmg/cache, ./wheels)
   ↓ (if missing)
Tier 3: Standalone Executable (pmg / pmg.exe binary)
   ↓ (if missing)
Tier 4: Network Installation (pip install, only when PMG_OFFLINE_MODE is not active)
   ↓ (if missing)
Tier 5: Compatible Alternative Backend (e.g. stdlib xml.etree for read-only inspection)
   ↓ (if missing)
Tier 6: Explicit Unsupported Report (never fabricate execution or silent failures)
```

---

## 2. Core vs. Optional Dependencies

| Category | Packages | Purpose | Mandatory for Compilation? |
| :--- | :--- | :--- | :--- |
| **Core Engine** | `python-pptx>=0.6.21`<br>`lxml>=4.9.0`<br>`Pillow>=9.0.0` | 16:9 canvas geometry, editable shapes, OpenXML Morph transitions, and Animation Pane timing trees. | **YES** |
| **Visual CV & Media** | `opencv-python>=4.5.0`<br>`numpy>=1.21.0`<br>`scipy>=1.7.0` | Video motion reconstruction, AI visual review, and image decomposition. | **NO (Optional)** |
| **Windows COM** | `pywin32>=300` | Local Microsoft PowerPoint COM automation and clean open verification. | **NO (Optional)** |

---

## 3. Creating an Offline Wheel Bundle

To package PMG for completely offline air-gapped systems or isolated sandboxes without internet:

### Step 1: Download Pure-Python & Platform-Specific Wheels
Run the following on an internet-connected machine matching your target OS and Python version:

```bash
# For minimal core offline bundle (pure python / common abi):
pip download --dest ./wheels python-pptx lxml Pillow

# For complete offline bundle including CV analysis:
pip download --dest ./wheels -e ".[full]"
```

### Step 2: Distribute Wheels to the Target Environment
Place the `.whl` files in one of PMG's automatic discovery locations:
1. Environment variable: `export PMG_OFFLINE_CACHE=/path/to/wheels`
2. User home cache: `~/.pmg/cache/`
3. Repository workspace: `./wheels/`

### Step 3: Verify with PMG Diagnostics
```bash
python -m pmg doctor
```

PMG's `pmg.dependencies.resolve_dependencies()` will inspect the directory, match the ABI/platform tags against the host, and confirm readiness.

---

## 4. Troubleshooting Offline Environments

### "No module named 'cv2'"
- **Cause:** OpenCV is not installed in the sandbox.
- **Resolution:** In PMG v2.1.0, OpenCV is strictly optional. Slide compilation, Morph transitions, and Level 1 structural validation work without `cv2`. Video motion reconstruction and AI visual review will report `SKIPPED`.

### "No module named 'win32com'"
- **Cause:** Host is non-Windows (Linux/macOS) or `pywin32` is not installed.
- **Resolution:** PMG automatically detects this and enters Mode A (Sandbox) or Mode B (macOS). Level 3 playback validation is accurately reported as `NOT_TESTED — PowerPoint Desktop unavailable`. Compilation is unaffected.

### "Network access is prohibited"
- **Cause:** Running inside an isolated sandbox with no outbound socket access.
- **Resolution:** Set `export PMG_OFFLINE_MODE=1`. PMG disables all network probes and uses purely local cached resources.

