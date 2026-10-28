"""
pmg.version
===========
Engine version, runtime dependency verification, and provenance tracking for PMG Skills (v1.5.1).
"""

import os
import sys
import shutil
import platform
import subprocess
import importlib
import json
from importlib import metadata
from datetime import datetime, timezone
from typing import Dict, Any, List

VERSION = "2.0.1"


def get_repo_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_git_commit() -> str:
    """Returns the current Git commit SHA or 'unknown' if outside a git worktree."""
    installed = _installed_vcs_info()
    if installed.get("commit_id"):
        return installed["commit_id"]
    try:
        res = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=get_repo_root(),
            capture_output=True,
            text=True,
            timeout=5,
        )
        if res.returncode == 0 and res.stdout.strip():
            return res.stdout.strip()
    except Exception:
        pass
    return "unknown"


def get_git_branch() -> str:
    """Returns the current Git branch name."""
    installed = _installed_vcs_info()
    if installed.get("commit_id"):
        return installed.get("requested_revision", "installed")
    try:
        res = subprocess.run(
            ["git", "rev-parse", "--abbrev-ref", "HEAD"],
            cwd=get_repo_root(),
            capture_output=True,
            text=True,
            timeout=5,
        )
        if res.returncode == 0 and res.stdout.strip():
            return res.stdout.strip()
    except Exception:
        pass
    return "unknown"


def _installed_vcs_info() -> Dict[str, Any]:
    """Read pip's actual pinned VCS origin when running an installed package."""
    if os.path.isdir(os.path.join(get_repo_root(), ".git")):
        return {}
    try:
        distribution = metadata.distribution("pmg-skills")
        installed_module = distribution.locate_file("pmg/version.py")
        if os.path.realpath(installed_module) != os.path.realpath(__file__):
            return {}
        direct_url = json.loads(distribution.read_text("direct_url.json") or "{}")
        return direct_url.get("vcs_info", {})
    except (metadata.PackageNotFoundError, ValueError, OSError, TypeError):
        return {}


def verify_runtime_dependencies(require_windows_com: bool = False) -> Dict[str, Any]:
    """
    Verifies required Python packages and system binaries at runtime (Upgrade C).
    Returns a structured diagnostic report with actionable remediation instructions.
    """
    core_pkgs = {
        "pptx": "python-pptx>=0.6.21",
        "lxml": "lxml>=4.9.0",
        "PIL": "Pillow>=10.0.0",
        "numpy": "numpy>=1.24.0",
        "cv2": "opencv-python>=4.8.0",
    }
    missing_pkgs: List[str] = []
    installed_pkgs: Dict[str, bool] = {}

    for mod_name, pip_spec in core_pkgs.items():
        try:
            importlib.import_module(mod_name)
            installed_pkgs[mod_name] = True
        except ImportError:
            installed_pkgs[mod_name] = False
            missing_pkgs.append(pip_spec)

    win32_ok = False
    if sys.platform == "win32":
        try:
            importlib.import_module("win32com.client")
            win32_ok = True
        except ImportError:
            if require_windows_com:
                missing_pkgs.append("pywin32>=306")

    ffmpeg_ok = shutil.which("ffmpeg") is not None
    ffprobe_ok = shutil.which("ffprobe") is not None

    return {
        "ok": len(missing_pkgs) == 0,
        "engine_version": VERSION,
        "platform": sys.platform,
        "installed_packages": installed_pkgs,
        "missing_packages": missing_pkgs,
        "win32com_available": win32_ok,
        "ffmpeg_available": ffmpeg_ok,
        "ffprobe_available": ffprobe_ok,
        "remediation": (
            f"pip install {' '.join(missing_pkgs)}" if missing_pkgs else "All required Python dependencies are installed."
        ),
    }


def get_engine_provenance(
    entrypoint: str = "pmg.orchestrator",
    validation_method: str = "local_powerpoint_com",
    worker_identity: str = "local",
) -> Dict[str, Any]:
    """Returns full provenance metadata for the PMG execution."""
    return {
        "engine_name": "pmg-skills",
        "engine_version": VERSION,
        "git_commit": get_git_commit(),
        "git_branch": get_git_branch(),
        "canonical_entrypoint": entrypoint,
        "validation_method": validation_method,
        "worker_identity": worker_identity,
        "python_version": sys.version.split()[0],
        "platform": platform.platform(),
        "os_system": platform.system(),
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
    }
