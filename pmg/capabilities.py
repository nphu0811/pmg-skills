"""
pmg.capabilities
================
Universal capability detection and environment modeling for PMG Skills (v2.1.0).

Determines runtime capabilities across:
- Mode A: Web / Sandbox (no PowerPoint, no net/exec, offline)
- Mode B: Local Desktop (Windows, macOS, Linux with/without PowerPoint)
- Mode C: Remote Worker (configured & authorized Windows validation worker)

Provides deterministic detection, explicit available/unavailable/unknown states,
safe dependency injection for testing, and zero mandatory dependency on Microsoft
PowerPoint for slide compilation.
"""

import os
import sys
import shutil
import socket
import platform
import importlib
from enum import Enum
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Callable, Tuple


class CapabilityState(str, Enum):
    AVAILABLE = "available"
    UNAVAILABLE = "unavailable"
    UNKNOWN = "unknown"


class ExecutionMode(str, Enum):
    WEB_SANDBOX = "web_sandbox"
    LOCAL_DESKTOP = "local_desktop"
    REMOTE_WORKER = "remote_worker"


@dataclass
class EnvironmentCapabilities:
    """Detailed capability model describing the current runtime environment."""
    platform: str
    os_name: str
    python_available: bool = True
    python_version: str = field(default_factory=lambda: sys.version.split()[0])
    network_available: bool = False
    powerpoint_available: bool = False
    com_available: bool = False
    renderer_available: bool = False
    renderer_type: Optional[str] = None  # "powerpoint", "libreoffice", None
    pptx_available: bool = False
    lxml_available: bool = False
    pillow_available: bool = False
    numpy_available: bool = False
    cv2_available: bool = False
    scipy_available: bool = False
    ffmpeg_available: bool = False
    ffprobe_available: bool = False
    remote_worker_configured: bool = False
    remote_worker_available: bool = False
    remote_worker_url: Optional[str] = None
    available_backends: List[str] = field(default_factory=list)
    execution_mode: ExecutionMode = ExecutionMode.LOCAL_DESKTOP
    state_details: Dict[str, CapabilityState] = field(default_factory=dict)
    notes: List[str] = field(default_factory=list)

    @property
    def can_compile_native_pptx(self) -> bool:
        """Creating native PowerPoint animations requires ONLY python-pptx and lxml."""
        return self.pptx_available and self.lxml_available

    @property
    def can_validate_structural(self) -> bool:
        """Level 1 structural validation requires only standard zipfile, lxml/xml."""
        return self.lxml_available or True  # XML parser available in stdlib

    @property
    def can_validate_visual(self) -> bool:
        """Level 2 visual validation requires a rendering backend (PowerPoint or LibreOffice)."""
        return self.renderer_available

    @property
    def can_validate_playback(self) -> bool:
        """Level 3 playback validation requires PowerPoint Desktop (local COM or remote worker)."""
        return (self.com_available and self.powerpoint_available) or self.remote_worker_available

    def to_dict(self) -> Dict[str, Any]:
        return {
            "platform": self.platform,
            "os_name": self.os_name,
            "python_available": self.python_available,
            "python_version": self.python_version,
            "network_available": self.network_available,
            "powerpoint_available": self.powerpoint_available,
            "com_available": self.com_available,
            "renderer_available": self.renderer_available,
            "renderer_type": self.renderer_type,
            "pptx_available": self.pptx_available,
            "lxml_available": self.lxml_available,
            "pillow_available": self.pillow_available,
            "numpy_available": self.numpy_available,
            "cv2_available": self.cv2_available,
            "scipy_available": self.scipy_available,
            "ffmpeg_available": self.ffmpeg_available,
            "ffprobe_available": self.ffprobe_available,
            "remote_worker_configured": self.remote_worker_configured,
            "remote_worker_available": self.remote_worker_available,
            "remote_worker_url": self.remote_worker_url,
            "available_backends": list(self.available_backends),
            "execution_mode": self.execution_mode.value,
            "can_compile_native_pptx": self.can_compile_native_pptx,
            "can_validate_structural": self.can_validate_structural,
            "can_validate_visual": self.can_validate_visual,
            "can_validate_playback": self.can_validate_playback,
            "state_details": {k: v.value for k, v in self.state_details.items()},
            "notes": list(self.notes),
        }

    def format_diagnostic_report(self) -> str:
        """Produces a clean human-readable diagnostic report for 'pmg doctor'."""
        lines = [
            "==================================================",
            "PMG Skills Universal Runtime - Environment Diagnostics",
            "==================================================",
            f"Execution Mode:       {self.execution_mode.value.upper()}",
            f"Platform:             {self.platform} ({self.os_name})",
            f"Python Runtime:       {self.python_version} ({'AVAILABLE' if self.python_available else 'UNAVAILABLE'})",
            f"Network Access:       {'AVAILABLE' if self.network_available else 'UNAVAILABLE / RESTRICTED'}",
            "",
            "Core Compilation Stack (Level 0 - PowerPoint Not Required):",
            f"  python-pptx:        {'INSTALLED' if self.pptx_available else 'MISSING'}",
            f"  lxml:               {'INSTALLED' if self.lxml_available else 'MISSING'}",
            f"  Pillow:             {'INSTALLED' if self.pillow_available else 'MISSING'}",
            f"  Native Compilation: {'READY' if self.can_compile_native_pptx else 'NOT READY (Missing pptx or lxml)'}",
            "",
            "Validation & Playback Backends:",
            f"  Level 1 Structural: {'READY' if self.can_validate_structural else 'NOT READY'}",
            f"  Level 2 Visual:     {'READY (' + str(self.renderer_type) + ')' if self.renderer_available else 'SKIPPED (No renderer)'}",
            f"  Level 3 Playback:   {'READY (Local COM)' if (self.com_available and self.powerpoint_available) else ('READY (Remote Worker)' if self.remote_worker_available else 'NOT AVAILABLE (PowerPoint Desktop not found)')}",
            "",
            "Optional Media & CV Tools:",
            f"  OpenCV (cv2):       {'INSTALLED' if self.cv2_available else 'NOT INSTALLED (Video analysis disabled)'}",
            f"  NumPy:              {'INSTALLED' if self.numpy_available else 'NOT INSTALLED'}",
            f"  SciPy:              {'INSTALLED' if self.scipy_available else 'NOT INSTALLED'}",
            f"  FFmpeg / FFprobe:   {'INSTALLED' if (self.ffmpeg_available and self.ffprobe_available) else 'NOT INSTALLED'}",
            "",
            "Remote Worker (Mode C):",
            f"  Configured:         {'YES (' + str(self.remote_worker_url) + ')' if self.remote_worker_configured else 'NO'}",
            f"  Status:             {'REACHABLE' if self.remote_worker_available else 'NOT ACTIVE'}",
            "",
            f"Active Backends:      {', '.join(self.available_backends) if self.available_backends else 'None'}",
            "==================================================",
        ]
        if self.notes:
            lines.append("Notes & Remediation:")
            for n in self.notes:
                lines.append(f"  * {n}")
            lines.append("==================================================")
        return "\n".join(lines)


def _check_module(module_name: str) -> bool:
    try:
        importlib.import_module(module_name)
        return True
    except (ImportError, Exception):
        return False


def _probe_windows_powerpoint() -> bool:
    """Checks for PowerPoint COM registration on Windows without launching PowerPoint."""
    if sys.platform != "win32":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID"):
            return True
    except Exception:
        return False


def _probe_macos_powerpoint() -> bool:
    """Checks for Microsoft PowerPoint app bundle on macOS."""
    if sys.platform != "darwin":
        return False
    mac_paths = [
        "/Applications/Microsoft PowerPoint.app",
        os.path.expanduser("~/Applications/Microsoft PowerPoint.app"),
    ]
    return any(os.path.exists(p) for p in mac_paths)


def _probe_renderer() -> Tuple[bool, Optional[str]]:
    """Detects available slide rendering backends."""
    # 1. Windows PowerPoint Desktop COM
    if sys.platform == "win32" and _probe_windows_powerpoint():
        try:
            import win32com.client  # noqa: F401
            return True, "powerpoint"
        except Exception:
            pass

    # 2. Headless LibreOffice
    soffice_path = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice_path:
        return True, "libreoffice"

    # Standard Windows LibreOffice installation path
    if sys.platform == "win32":
        for candidate in [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]:
            if os.path.exists(candidate):
                return True, "libreoffice"

    return False, None


def _probe_network(timeout_sec: float = 0.5) -> bool:
    """Quick non-blocking probe for outbound network connectivity."""
    if os.environ.get("PMG_OFFLINE_MODE", "").lower() in ("1", "true", "yes"):
        return False
    if os.environ.get("PMG_NETWORK_AVAILABLE", "").lower() in ("1", "true", "yes"):
        return True
    if os.environ.get("PMG_NETWORK_AVAILABLE", "").lower() in ("0", "false", "no"):
        return False

    try:
        # Non-blocking probe to common reliable DNS IP
        with socket.create_connection(("8.8.8.8", 53), timeout=timeout_sec):
            return True
    except Exception:
        return False


def detect_environment(
    probes_override: Optional[Dict[str, Any]] = None,
    test_network: bool = False,
    test_remote: bool = False,
    worker_url: Optional[str] = None,
    worker_token: Optional[str] = None,
) -> EnvironmentCapabilities:
    """
    Deterministically detects the host environment and available PMG capabilities.

    Supports dependency injection via `probes_override` for unit testing mock platforms
    (e.g. Linux CI, macOS, offline sandbox) without needing physical hosts.
    """
    probes = probes_override or {}

    plat = probes.get("platform", sys.platform)
    os_name = probes.get("os_name", platform.system())
    py_avail = probes.get("python_available", True)
    py_ver = probes.get("python_version", sys.version.split()[0])

    # Core python packages
    pptx_ok = probes.get("pptx_available", _check_module("pptx"))
    lxml_ok = probes.get("lxml_available", _check_module("lxml"))
    pillow_ok = probes.get("pillow_available", _check_module("PIL"))
    numpy_ok = probes.get("numpy_available", _check_module("numpy"))
    cv2_ok = probes.get("cv2_available", _check_module("cv2"))
    scipy_ok = probes.get("scipy_available", _check_module("scipy"))

    # System binaries
    ffmpeg_ok = probes.get("ffmpeg_available", shutil.which("ffmpeg") is not None)
    ffprobe_ok = probes.get("ffprobe_available", shutil.which("ffprobe") is not None)

    # Windows COM
    if "com_available" in probes:
        com_ok = bool(probes["com_available"])
    elif plat == "win32":
        com_ok = _check_module("win32com.client")
    else:
        com_ok = False

    # PowerPoint availability (never infer solely from OS!)
    if "powerpoint_available" in probes:
        ppt_ok = bool(probes["powerpoint_available"])
    elif plat == "win32":
        ppt_ok = com_ok and _probe_windows_powerpoint()
    elif plat == "darwin":
        ppt_ok = _probe_macos_powerpoint()
    else:
        ppt_ok = False

    # Renderer availability
    if "renderer_available" in probes:
        ren_ok = bool(probes["renderer_available"])
        ren_type = probes.get("renderer_type", "custom" if ren_ok else None)
    else:
        ren_ok, ren_type = _probe_renderer()

    # Network availability (explicitly probed or configured; never blindly assumed!)
    if "network_available" in probes:
        net_ok = bool(probes["network_available"])
    elif test_network or os.environ.get("PMG_TEST_NETWORK", "") == "1":
        net_ok = _probe_network()
    else:
        # If not testing network, check explicit environment variable or default to False in offline/sandbox
        val = os.environ.get("PMG_NETWORK_AVAILABLE", "")
        if val.lower() in ("1", "true", "yes"):
            net_ok = True
        elif val.lower() in ("0", "false", "no"):
            net_ok = False
        else:
            net_ok = False

    # Remote worker (Mode C)
    w_url = worker_url or probes.get("remote_worker_url") or os.environ.get("PMG_QA_WORKER_URL")
    w_tok = worker_token or os.environ.get("PMG_QA_WORKER_TOKEN")
    worker_cfg = bool(w_url)
    worker_ok = False

    if "remote_worker_available" in probes:
        worker_ok = bool(probes["remote_worker_available"])
    elif worker_cfg and (test_remote or os.environ.get("PMG_TEST_WORKER", "") == "1"):
        try:
            from pmg.worker.client import PMGWorkerClient
            client = PMGWorkerClient(worker_url=w_url, worker_token=w_tok)
            status = client.get_remote_status(timeout_sec=2)
            worker_ok = status.get("reachable", False)
        except Exception:
            worker_ok = False

    # Available backends
    backends: List[str] = []
    if pptx_ok and lxml_ok:
        backends.append("native_compiler")
    if ren_ok and ren_type:
        backends.append(f"{ren_type}_renderer")
    if plat == "win32" and com_ok and ppt_ok:
        backends.append("windows_powerpoint_com")
    if plat == "darwin" and ppt_ok:
        backends.append("macos_powerpoint_automation")
    if worker_ok:
        backends.append("remote_windows_worker")

    # Determine execution mode
    forced_mode = os.environ.get("PMG_EXECUTION_MODE", "").lower()
    if forced_mode == "sandbox" or forced_mode == "web_sandbox":
        mode = ExecutionMode.WEB_SANDBOX
    elif forced_mode == "remote" or forced_mode == "remote_worker":
        mode = ExecutionMode.REMOTE_WORKER
    elif worker_cfg and (worker_ok or not (plat == "win32" and ppt_ok)):
        mode = ExecutionMode.REMOTE_WORKER
    elif ppt_ok:
        mode = ExecutionMode.LOCAL_DESKTOP
    elif not ppt_ok or plat == "linux" or os.environ.get("PMG_SANDBOX", "") == "1":
        mode = ExecutionMode.WEB_SANDBOX
    else:
        mode = ExecutionMode.LOCAL_DESKTOP

    # State details
    states: Dict[str, CapabilityState] = {
        "python": CapabilityState.AVAILABLE if py_avail else CapabilityState.UNAVAILABLE,
        "native_compiler": CapabilityState.AVAILABLE if (pptx_ok and lxml_ok) else CapabilityState.UNAVAILABLE,
        "powerpoint": CapabilityState.AVAILABLE if ppt_ok else CapabilityState.UNAVAILABLE,
        "renderer": CapabilityState.AVAILABLE if ren_ok else CapabilityState.UNAVAILABLE,
        "network": CapabilityState.AVAILABLE if net_ok else CapabilityState.UNAVAILABLE,
        "remote_worker": CapabilityState.AVAILABLE if worker_ok else (
            CapabilityState.UNAVAILABLE if worker_cfg else CapabilityState.UNAVAILABLE
        ),
    }

    notes: List[str] = []
    if not (pptx_ok and lxml_ok):
        notes.append("Install 'python-pptx' and 'lxml' to enable native PowerPoint generation.")
    if not ppt_ok:
        notes.append("Microsoft PowerPoint Desktop is not installed. Native slide compilation works normally; Level 3 playback validation will be reported as NOT_TESTED.")
    if not ren_ok:
        notes.append("No slide renderer found. Level 2 visual validation will be reported as SKIPPED.")
    if not net_ok:
        notes.append("Network access is offline or restricted. Offline-first dependency resolution active.")

    return EnvironmentCapabilities(
        platform=plat,
        os_name=os_name,
        python_available=py_avail,
        python_version=py_ver,
        network_available=net_ok,
        powerpoint_available=ppt_ok,
        com_available=com_ok,
        renderer_available=ren_ok,
        renderer_type=ren_type,
        pptx_available=pptx_ok,
        lxml_available=lxml_ok,
        pillow_available=pillow_ok,
        numpy_available=numpy_ok,
        cv2_available=cv2_ok,
        scipy_available=scipy_ok,
        ffmpeg_available=ffmpeg_ok,
        ffprobe_available=ffprobe_ok,
        remote_worker_configured=worker_cfg,
        remote_worker_available=worker_ok,
        remote_worker_url=w_url if worker_cfg else None,
        available_backends=backends,
        execution_mode=mode,
        state_details=states,
        notes=notes,
    )
