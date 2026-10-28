"""
pmg.worker
==========
Windows PowerPoint Validation Worker package (Upgrade A — PMG v1.5.1).
Allows AI agents on Linux, macOS, cloud containers, or Windows to run real
Microsoft PowerPoint playback validation either locally or via a remote Windows worker.
"""

from pmg.worker.result_schema import (
    WorkerHealthStatus,
    WorkerJobRequest,
    WorkerJobResult,
)
from pmg.worker.security import (
    verify_bearer_token,
    validate_safe_pptx_upload,
    redact_secrets,
    IsolatedJobWorkspace,
)
from pmg.worker.powerpoint_runner import (
    check_worker_health,
    run_worker_native_export,
)
from pmg.worker.video_capture import run_worker_slideshow_capture
from pmg.worker.jobs import WorkerJobManager
from pmg.worker.server import PMGValidationWorkerServer
from pmg.worker.client import PMGWorkerClient

__all__ = [
    "WorkerHealthStatus",
    "WorkerJobRequest",
    "WorkerJobResult",
    "verify_bearer_token",
    "validate_safe_pptx_upload",
    "redact_secrets",
    "IsolatedJobWorkspace",
    "check_worker_health",
    "run_worker_native_export",
    "run_worker_slideshow_capture",
    "WorkerJobManager",
    "PMGValidationWorkerServer",
    "PMGWorkerClient",
]

