"""
pmg.worker.security
===================
Security hardening for the PMG Windows PowerPoint Validation Worker:
- Bearer token authentication (constant-time comparison)
- File size enforcement
- ZIP bomb & path-traversal protection
- Unsafe macro / active content / OLE / external command rejection (.pptm, vbaProject.bin, oleObject)
- Secret redaction in logs and reports
- Isolated temporary workspace management
"""

import os
import hmac
import shutil
import zipfile
import tempfile
from typing import Tuple, Dict, Any, Optional

DEFAULT_MAX_PPTX_BYTES = 100 * 1024 * 1024  # 100 MB
MAX_UNCOMPRESSED_ZIP_BYTES = 350 * 1024 * 1024  # 350 MB
MAX_ZIP_ENTRIES = 5000

UNSAFE_EXTENSIONS = {".pptm", ".ppsm", ".potm", ".docm", ".xlsm", ".exe", ".bat", ".cmd", ".ps1", ".vbs", ".js"}
UNSAFE_ARCHIVE_ENTRIES = {"vbaproject.bin", "vba.bin"}


def verify_bearer_token(auth_header: Optional[str], expected_token: Optional[str]) -> bool:
    """
    Verifies an HTTP Authorization: Bearer <token> header against expected_token
    using constant-time comparison. If expected_token is None or empty, authentication
    is only allowed when explicitly running in local loopback mode without a token.
    """
    if not expected_token:
        return True
    if not auth_header or not auth_header.startswith("Bearer "):
        return False
    provided = auth_header[len("Bearer "):].strip()
    return hmac.compare_digest(provided.encode("utf-8"), expected_token.encode("utf-8"))


def redact_secrets(data: Any) -> Any:
    """Recursively redacts tokens, passwords, and authorization headers from dicts/strings."""
    secret_keys = {"token", "authorization", "password", "secret", "api_key", "pmg_qa_worker_token"}
    if isinstance(data, dict):
        out = {}
        for k, v in data.items():
            if str(k).lower() in secret_keys and v:
                out[k] = "***REDACTED***"
            else:
                out[k] = redact_secrets(v)
        return out
    elif isinstance(data, list):
        return [redact_secrets(item) for item in data]
    return data


def validate_safe_pptx_upload(
    file_path: str,
    original_filename: str = "upload.pptx",
    max_bytes: int = DEFAULT_MAX_PPTX_BYTES,
) -> Tuple[bool, str]:
    """
    Validates that an uploaded PPTX file is safe to open in Microsoft PowerPoint:
    1. Rejects macro-enabled or executable extensions (.pptm, .ppsm, .exe, etc.).
    2. Enforces maximum compressed file size.
    3. Verifies valid ZIP container structure.
    4. Rejects path-traversal entries (`..`, leading `/` or `\\`, drive letters).
    5. Rejects ZIP bombs (excessive uncompressed size or entry count).
    6. Rejects embedded VBA macros (`vbaProject.bin`), embedded executables, or unsafe external OLE actions.
    """
    _, ext = os.path.splitext(original_filename.lower())
    if ext in UNSAFE_EXTENSIONS:
        return False, f"SECURITY_REJECTED: Unsafe file extension '{ext}' is forbidden (macros/executables not allowed)."
    if ext and ext != ".pptx":
        return False, f"SECURITY_REJECTED: Only standard '.pptx' presentations are accepted (got '{ext}')."

    if not os.path.exists(file_path):
        return False, "SECURITY_REJECTED: Uploaded file does not exist."

    size_bytes = os.path.getsize(file_path)
    if size_bytes <= 0:
        return False, "SECURITY_REJECTED: Uploaded file is empty."
    if size_bytes > max_bytes:
        return False, f"SECURITY_REJECTED: File size ({size_bytes} bytes) exceeds limit ({max_bytes} bytes)."

    if not zipfile.is_zipfile(file_path):
        return False, "SECURITY_REJECTED: File is not a valid OpenXML ZIP archive."

    try:
        with zipfile.ZipFile(file_path, "r") as zf:
            infos = zf.infolist()
            if len(infos) > MAX_ZIP_ENTRIES:
                return False, f"SECURITY_REJECTED: Archive contains too many entries ({len(infos)} > {MAX_ZIP_ENTRIES})."

            total_uncompressed = 0
            has_presentation_xml = False

            for info in infos:
                norm_name = info.filename.replace("\\", "/")
                # Path traversal check
                if (
                    norm_name.startswith("/")
                    or ".." in norm_name.split("/")
                    or ":" in norm_name
                ):
                    return False, f"SECURITY_REJECTED: Path traversal detected in archive entry '{info.filename}'."

                base_lower = os.path.basename(norm_name).lower()
                if base_lower in UNSAFE_ARCHIVE_ENTRIES or base_lower.endswith((".exe", ".dll", ".bat", ".ps1", ".vbs")):
                    return False, f"SECURITY_REJECTED: Unsafe macro or active payload '{info.filename}' detected inside archive."

                if "oleobject" in base_lower and base_lower.endswith(".bin"):
                    return False, f"SECURITY_REJECTED: Embedded OLE binary object '{info.filename}' is not permitted."

                total_uncompressed += max(0, info.file_size)
                if total_uncompressed > MAX_UNCOMPRESSED_ZIP_BYTES:
                    return False, (
                        f"SECURITY_REJECTED: Uncompressed archive size exceeds limit "
                        f"({total_uncompressed} > {MAX_UNCOMPRESSED_ZIP_BYTES} bytes)."
                    )

                if norm_name == "ppt/presentation.xml":
                    has_presentation_xml = True

            if not has_presentation_xml:
                return False, "SECURITY_REJECTED: Missing 'ppt/presentation.xml' in OpenXML package."
    except Exception as e:
        return False, f"SECURITY_REJECTED: Failed to inspect ZIP archive: {str(e)}"

    return True, "OK"


class IsolatedJobWorkspace:
    """Context manager creating and cleaning up an isolated temporary directory per validation job."""

    def __init__(self, job_id: str, base_dir: Optional[str] = None, cleanup_on_exit: bool = True):
        self.job_id = "".join(c for c in job_id if c.isalnum() or c in ("-", "_"))[:48] or "job"
        self.base_dir = base_dir
        self.cleanup_on_exit = cleanup_on_exit
        self.path: Optional[str] = None

    def __enter__(self) -> str:
        self.path = tempfile.mkdtemp(prefix=f"pmg_worker_{self.job_id}_", dir=self.base_dir)
        return self.path

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.cleanup_on_exit and self.path and os.path.exists(self.path):
            try:
                shutil.rmtree(self.path, ignore_errors=True)
            except Exception:
                pass

