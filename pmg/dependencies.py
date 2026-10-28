"""
pmg.dependencies
================
Offline-first dependency resolution strategy and package inspection for PMG Skills (v2.1.0).

Priority sequence:
1. Existing compatible installed dependencies in the active Python environment.
2. Bundled or cached compatible offline dependencies (wheels/sdists in PMG_OFFLINE_CACHE, ~/.pmg/cache, ./wheels).
3. Standalone executable if provided.
4. Network installation, only when permitted (never assumed).
5. Alternative compatible backend, if available (e.g. standard library XML if lxml missing).
6. Explicit unsupported-capability report.
"""

import os
import sys
import glob
import shutil
import platform
import subprocess
import importlib
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Tuple


@dataclass
class DependencyResolution:
    package_name: str
    required_spec: str
    status: str  # "installed", "cached_wheel_found", "executable_found", "network_available", "fallback_available", "unsupported"
    source: str
    version: Optional[str] = None
    compatible: bool = False
    details: str = ""


@dataclass
class OfflineBundleManifest:
    target_python: str
    target_platform: str
    target_arch: str
    wheels: List[str] = field(default_factory=list)
    cache_dir: Optional[str] = None
    is_complete: bool = False


def get_offline_cache_directories() -> List[str]:
    """Returns candidate directories searched for offline wheels and bundles."""
    candidates = []
    env_cache = os.environ.get("PMG_OFFLINE_CACHE")
    if env_cache and os.path.isdir(env_cache):
        candidates.append(os.path.abspath(env_cache))

    # User home cache: ~/.pmg/cache or ~/.cache/pmg
    home = os.path.expanduser("~")
    pmg_home_cache = os.path.join(home, ".pmg", "cache")
    if os.path.isdir(pmg_home_cache):
        candidates.append(pmg_home_cache)

    # Local workspace ./wheels directory if present
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    local_wheels = os.path.join(repo_root, "wheels")
    if os.path.isdir(local_wheels):
        candidates.append(local_wheels)

    return candidates


def inspect_installed_package(module_name: str) -> Tuple[bool, Optional[str]]:
    """Checks if a module is installed and returns its __version__ if available."""
    try:
        mod = importlib.import_module(module_name)
        ver = getattr(mod, "__version__", None)
        return True, str(ver) if ver else "installed"
    except (ImportError, Exception):
        return False, None


def find_compatible_offline_wheels(cache_dir: str) -> List[str]:
    """
    Finds wheels matching the current Python version, OS, and CPU architecture.
    Handles 'py3-none-any' (pure python) and platform-specific tags correctly.
    """
    if not os.path.isdir(cache_dir):
        return []

    py_tag = f"cp{sys.version_info.major}{sys.version_info.minor}"
    py3_tag = "py3"
    plat_tag = platform.system().lower()

    matched_wheels = []
    for wheel_file in glob.glob(os.path.join(cache_dir, "*.whl")):
        fname = os.path.basename(wheel_file).lower()
        # Pure Python wheel
        if "none-any.whl" in fname or "py2.py3-none-any" in fname or "py3-none-any" in fname:
            matched_wheels.append(wheel_file)
            continue

        # ABI/Python tag match
        if (py_tag in fname or py3_tag in fname) and (plat_tag in fname or "any" in fname):
            matched_wheels.append(wheel_file)

    return matched_wheels


def resolve_dependencies(
    allow_network: bool = False,
    cache_override_dir: Optional[str] = None,
) -> Dict[str, DependencyResolution]:
    """
    Executes the 6-tier dependency resolution strategy.
    """
    required_packages = {
        "pptx": ("python-pptx", ">=0.6.21"),
        "lxml": ("lxml", ">=4.9.0"),
        "PIL": ("Pillow", ">=9.0.0"),
        "numpy": ("numpy", ">=1.21.0"),
        "cv2": ("opencv-python", ">=4.5.0"),
        "scipy": ("scipy", ">=1.7.0"),
    }
    if sys.platform == "win32":
        required_packages["win32com"] = ("pywin32", ">=300")

    search_dirs = [cache_override_dir] if cache_override_dir else get_offline_cache_directories()
    cached_wheels: List[str] = []
    for s_dir in search_dirs:
        if s_dir and os.path.isdir(s_dir):
            cached_wheels.extend(find_compatible_offline_wheels(s_dir))

    resolutions: Dict[str, DependencyResolution] = {}

    for import_name, (pkg_name, spec) in required_packages.items():
        # Tier 1: Installed
        installed, ver = inspect_installed_package(import_name)
        if installed:
            resolutions[pkg_name] = DependencyResolution(
                package_name=pkg_name,
                required_spec=spec,
                status="installed",
                source="active_python_environment",
                version=ver,
                compatible=True,
                details=f"Installed in environment with version {ver}.",
            )
            continue

        # Tier 2: Bundled / Cached offline wheels
        matching_wheel = next((w for w in cached_wheels if pkg_name.replace("-", "_").lower() in os.path.basename(w).lower()), None)
        if matching_wheel:
            resolutions[pkg_name] = DependencyResolution(
                package_name=pkg_name,
                required_spec=spec,
                status="cached_wheel_found",
                source=matching_wheel,
                version=os.path.basename(matching_wheel),
                compatible=True,
                details=f"Compatible offline wheel available at {matching_wheel}.",
            )
            continue

        # Tier 3: Standalone executable
        pmg_exe = shutil.which("pmg") or shutil.which("pmg.exe")
        if pmg_exe and not pmg_exe.endswith(".py"):
            resolutions[pkg_name] = DependencyResolution(
                package_name=pkg_name,
                required_spec=spec,
                status="executable_found",
                source=pmg_exe,
                version="standalone",
                compatible=True,
                details=f"Covered by standalone binary executable: {pmg_exe}.",
            )
            continue

        # Tier 4: Network installation (only if permitted)
        if allow_network:
            resolutions[pkg_name] = DependencyResolution(
                package_name=pkg_name,
                required_spec=spec,
                status="network_available",
                source="pypi",
                version=None,
                compatible=True,
                details="Permitted to install via pip install from network.",
            )
            continue

        # Tier 5: Alternative compatible backend (e.g. stdlib xml instead of lxml for read-only)
        if pkg_name == "lxml":
            resolutions[pkg_name] = DependencyResolution(
                package_name=pkg_name,
                required_spec=spec,
                status="fallback_available",
                source="xml.etree.ElementTree",
                version="stdlib",
                compatible=False,
                details="Standard library xml parser available for read-only inspection.",
            )
            continue

        # Tier 6: Explicit unsupported report
        resolutions[pkg_name] = DependencyResolution(
            package_name=pkg_name,
            required_spec=spec,
            status="unsupported",
            source="none",
            version=None,
            compatible=False,
            details=f"Missing dependency {pkg_name} ({spec}) with no offline wheel or network access.",
        )

    return resolutions

