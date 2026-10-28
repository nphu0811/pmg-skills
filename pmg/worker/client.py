"""
pmg.worker.client
=================
Client for the PMG Windows PowerPoint Validation Worker supporting:
- PMG_QA_WORKER_URL
- PMG_QA_WORKER_TOKEN
- PMG_QA_MODE (local, remote, auto, static)
"""

import os
import json
import urllib.request
import urllib.error
import urllib.parse
from typing import Optional, Dict, Any

from pmg.models import ValidationStatus
from pmg.worker.result_schema import WorkerJobResult


class PMGWorkerClient:
    """
    Communicates with a remote or local PMG Windows PowerPoint Validation Worker over HTTP/HTTPS.
    Enforces Bearer token authentication and bounded request timeouts.
    """

    def __init__(
        self,
        worker_url: Optional[str] = None,
        worker_token: Optional[str] = None,
        timeout_sec: int = 260,
    ):
        self.worker_url = (worker_url or os.environ.get("PMG_QA_WORKER_URL") or "").rstrip("/")
        self.worker_token = (
            worker_token if worker_token is not None else os.environ.get("PMG_QA_WORKER_TOKEN")
        )
        self.timeout_sec = timeout_sec

    def _headers(self, extra: Optional[Dict[str, str]] = None) -> Dict[str, str]:
        hdrs: Dict[str, str] = {"Accept": "application/json"}
        if self.worker_token:
            hdrs["Authorization"] = f"Bearer {self.worker_token}"
        if extra:
            hdrs.update(extra)
        return hdrs

    def get_remote_status(self, timeout_sec: float = 10.0) -> Dict[str, Any]:
        """Queries GET /health on the configured remote worker."""
        if not self.worker_url:
            return {
                "status": "blocked",
                "reachable": False,
                "error": "BLOCKED: PMG_QA_WORKER_URL is not configured.",
            }
        req = urllib.request.Request(
            f"{self.worker_url}/health",
            headers=self._headers(),
            method="GET",
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout_sec) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                data["reachable"] = True
                data["worker_url"] = self.worker_url
                return data
        except urllib.error.HTTPError as he:
            return {
                "status": "blocked",
                "reachable": False,
                "http_status": he.code,
                "error": f"BLOCKED: Remote worker returned HTTP {he.code}.",
            }
        except Exception as e:
            return {
                "status": "blocked",
                "reachable": False,
                "error": f"BLOCKED: Remote worker unreachable at {self.worker_url}: {str(e)}",
            }

    def validate_remote(
        self,
        pptx_path: str,
        hold_duration_ms: int = 1200,
        timeout_sec: Optional[int] = None,
    ) -> WorkerJobResult:
        """
        Submits a PPTX file to POST /validate on the remote Windows worker and returns
        a structured WorkerJobResult. If the worker is unconfigured or unreachable,
        returns validation_status = BLOCKED (never PASS).
        """
        eff_timeout = timeout_sec or self.timeout_sec
        if not self.worker_url:
            return WorkerJobResult(
                job_id="unconfigured",
                state="failed",
                validation_status=ValidationStatus.BLOCKED.value,
                worker_identity="none",
                error="BLOCKED: PMG_QA_WORKER_URL is not configured for remote PowerPoint validation.",
            )

        if not os.path.exists(pptx_path):
            return WorkerJobResult(
                job_id="missing_file",
                state="failed",
                validation_status=ValidationStatus.FAIL.value,
                worker_identity=self.worker_url,
                error=f"PPTX file does not exist: {pptx_path}",
            )

        filename = os.path.basename(pptx_path)
        with open(pptx_path, "rb") as f:
            payload = f.read()

        qs = urllib.parse.urlencode(
            {
                "filename": filename,
                "hold_ms": str(hold_duration_ms),
                "timeout_sec": str(eff_timeout),
            }
        )
        url = f"{self.worker_url}/validate?{qs}"
        hdrs = self._headers(
            {
                "Content-Type": "application/octet-stream",
                "X-PMG-Filename": filename,
                "X-PMG-Hold-Ms": str(hold_duration_ms),
                "X-PMG-Timeout-Sec": str(eff_timeout),
            }
        )
        req = urllib.request.Request(url, data=payload, headers=hdrs, method="POST")

        try:
            with urllib.request.urlopen(req, timeout=eff_timeout + 15) as resp:
                resp_data = json.loads(resp.read().decode("utf-8"))
                return WorkerJobResult(**{k: v for k, v in resp_data.items() if k in WorkerJobResult.__dataclass_fields__})
        except urllib.error.HTTPError as he:
            body_txt = ""
            try:
                body_txt = he.read().decode("utf-8", errors="replace")
                parsed_err = json.loads(body_txt)
                if isinstance(parsed_err, dict) and "validation_status" in parsed_err:
                    return WorkerJobResult(
                        job_id=parsed_err.get("job_id", "http_error"),
                        state="failed",
                        validation_status=parsed_err.get("validation_status", ValidationStatus.BLOCKED.value),
                        worker_identity=self.worker_url,
                        error=parsed_err.get("error", f"HTTP {he.code}"),
                    )
            except Exception:
                pass
            status_val = ValidationStatus.FAIL.value if he.code in (400, 413) else ValidationStatus.BLOCKED.value
            return WorkerJobResult(
                job_id="http_error",
                state="failed",
                validation_status=status_val,
                worker_identity=self.worker_url,
                error=f"Remote worker HTTP {he.code}: {body_txt or he.reason}",
            )
        except Exception as e:
            return WorkerJobResult(
                job_id="network_error",
                state="failed",
                validation_status=ValidationStatus.BLOCKED.value,
                worker_identity=self.worker_url,
                error=f"BLOCKED: Failed to reach remote Windows validation worker at {self.worker_url}: {str(e)}",
            )


    def compile_remote(
        self,
        project_spec: Any,
        output_pptx_path: str,
        timeout_sec: Optional[int] = None,
    ) -> bool:
        """
        Remote compilation adapter (distinguished from playback validation).
        Submits presentation spec to POST /compile and saves resulting PPTX artifact.
        If remote worker is not configured or unreachable, returns False without crashing local execution.
        """
        if not self.worker_url:
            return False

        eff_timeout = timeout_sec or min(60, self.timeout_sec)
        payload = json.dumps(project_spec).encode("utf-8") if isinstance(project_spec, (dict, list)) else (
            open(project_spec, "rb").read() if os.path.isfile(str(project_spec)) else str(project_spec).encode("utf-8")
        )

        url = f"{self.worker_url}/compile"
        hdrs = self._headers({"Content-Type": "application/json"})
        req = urllib.request.Request(url, data=payload, headers=hdrs, method="POST")

        try:
            with urllib.request.urlopen(req, timeout=eff_timeout) as resp:
                data = resp.read()
                if data.startswith(b"PK\x03\x04"):
                    os.makedirs(os.path.dirname(os.path.abspath(output_pptx_path)), exist_ok=True)
                    with open(output_pptx_path, "wb") as f:
                        f.write(data)
                    return True
            return False
        except Exception:
            return False
