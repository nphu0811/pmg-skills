"""
pmg.version
===========
Engine version and provenance tracking for PMG Skills.
"""

import os
import sys
import platform
import subprocess
from datetime import datetime, timezone
from typing import Dict, Any

VERSION = "1.1.0"


def get_repo_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_git_commit() -> str:
    """Returns the current Git commit SHA or 'unknown' if outside a git worktree."""
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


def get_engine_provenance(entrypoint: str = "pmg.orchestrator") -> Dict[str, Any]:
    """Returns full provenance metadata for the PMG execution."""
    return {
        "engine_name": "pmg-skills",
        "engine_version": VERSION,
        "git_commit": get_git_commit(),
        "git_branch": get_git_branch(),
        "canonical_entrypoint": entrypoint,
        "python_version": sys.version.split()[0],
        "platform": platform.platform(),
        "os_system": platform.system(),
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
    }

