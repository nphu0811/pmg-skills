"""
pmg.worker.server
=================
HTTP / HTTPS server for the PMG Windows PowerPoint Validation Worker (`pmg worker serve`).
Endpoints:
- GET  /health               -> WorkerHealthStatus JSON
- GET  /jobs/<job_id>        -> WorkerJobResult JSON
- POST /jobs/<job_id>/cancel -> Cancel a running job
- POST /validate             -> Submit a PPTX file (raw binary or base64 JSON) for full PowerPoint validation
"""

import os
import ssl
import json
import base64
import tempfile
import threading
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from typing import Optional, Tuple

from pmg.models import ValidationStatus
from pmg.version import VERSION
from pmg.worker.jobs import WorkerJobManager
from pmg.worker.powerpoint_runner import check_worker_health
from pmg.worker.security import verify_bearer_token, DEFAULT_MAX_PPTX_BYTES


class _WorkerRequestHandler(BaseHTTPRequestHandler):
    server_version = f"PMGValidationWorker/{VERSION}"

    def log_message(self, format: str, *args) -> None:
        # Quiet structured logging without leaking headers or tokens
        pass

    def _send_json(self, status_code: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _check_auth(self) -> bool:
        expected_token = getattr(self.server, "auth_token", None)
        auth_hdr = self.headers.get("Authorization")
        if not verify_bearer_token(auth_hdr, expected_token):
            self._send_json(
                401,
                {
                    "error": "UNAUTHORIZED: Valid Bearer token required.",
                    "validation_status": ValidationStatus.BLOCKED.value,
                },
            )
            return False
        return True

    def do_GET(self) -> None:
        if not self._check_auth():
            return

        parsed = urlparse(self.path)
        job_mgr: WorkerJobManager = getattr(self.server, "job_manager")

        if parsed.path == "/health":
            health = check_worker_health(
                active_jobs=job_mgr.active_job_count(),
                max_file_size_mb=int(job_mgr.max_file_size_bytes / (1024 * 1024)),
            )
            self._send_json(200, health.to_dict())
            return

        if parsed.path.startswith("/jobs/"):
            job_id = parsed.path.split("/jobs/", 1)[1].strip("/")
            job = job_mgr.get_job(job_id)
            if job is None:
                self._send_json(404, {"error": f"Job '{job_id}' not found."})
            else:
                self._send_json(200, job.to_dict())
            return

        self._send_json(404, {"error": f"Unknown endpoint '{parsed.path}'."})

    def do_POST(self) -> None:
        if not self._check_auth():
            return

        parsed = urlparse(self.path)
        job_mgr: WorkerJobManager = getattr(self.server, "job_manager")

        if parsed.path.startswith("/jobs/") and parsed.path.endswith("/cancel"):
            parts = parsed.path.strip("/").split("/")
            job_id = parts[1] if len(parts) >= 2 else ""
            ok = job_mgr.cancel_job(job_id)
            self._send_json(200 if ok else 404, {"job_id": job_id, "cancelled": ok})
            return

        if parsed.path == "/validate":
            content_len = int(self.headers.get("Content-Length", "0") or "0")
            if content_len <= 0:
                self._send_json(
                    400,
                    {
                        "error": "Empty request body.",
                        "validation_status": ValidationStatus.FAIL.value,
                    },
                )
                return
            if content_len > job_mgr.max_file_size_bytes + (2 * 1024 * 1024):
                self._send_json(
                    413,
                    {
                        "error": f"SECURITY_REJECTED: Payload size ({content_len} bytes) exceeds maximum allowed limit.",
                        "validation_status": ValidationStatus.FAIL.value,
                    },
                )
                return

            raw_body = self.rfile.read(content_len)
            qs = parse_qs(parsed.query)
            filename = self.headers.get("X-PMG-Filename") or (qs.get("filename", ["submission.pptx"])[0])
            hold_ms = int(self.headers.get("X-PMG-Hold-Ms") or (qs.get("hold_ms", ["1200"])[0]))
            timeout_sec = int(self.headers.get("X-PMG-Timeout-Sec") or (qs.get("timeout_sec", ["240"])[0]))

            content_type = (self.headers.get("Content-Type") or "").lower()
            pptx_bytes = raw_body
            if "application/json" in content_type:
                try:
                    data = json.loads(raw_body.decode("utf-8"))
                    filename = data.get("filename", filename)
                    hold_ms = int(data.get("hold_duration_ms", hold_ms))
                    timeout_sec = int(data.get("timeout_sec", timeout_sec))
                    b64_str = data.get("pptx_base64", "")
                    pptx_bytes = base64.b64decode(b64_str)
                except Exception as e:
                    self._send_json(
                        400,
                        {
                            "error": f"Invalid JSON payload: {str(e)}",
                            "validation_status": ValidationStatus.FAIL.value,
                        },
                    )
                    return

            _, ext = os.path.splitext(filename)
            if not ext:
                ext = ".pptx"
            fd, tmp_upload_path = tempfile.mkstemp(prefix="pmg_upload_", suffix=ext)
            try:
                with os.fdopen(fd, "wb") as f:
                    f.write(pptx_bytes)

                res = job_mgr.execute_job_sync(
                    pptx_source_path=tmp_upload_path,
                    original_filename=filename,
                    hold_duration_ms=hold_ms,
                    timeout_sec=timeout_sec,
                )
                status_code = 200 if res.error is None or not res.error.startswith("SECURITY_REJECTED") else 400
                self._send_json(status_code, res.to_dict())
            finally:
                if os.path.exists(tmp_upload_path):
                    try:
                        os.remove(tmp_upload_path)
                    except Exception:
                        pass
            return

        self._send_json(404, {"error": f"Unknown endpoint '{parsed.path}'."})


class PMGValidationWorkerServer:
    """
    Reusable HTTP/HTTPS server hosting the PMG Windows PowerPoint Validation Worker.
    Can run either in the foreground (`serve_forever`) or in a background thread for tests.
    """

    def __init__(
        self,
        host: str = "127.0.0.1",
        port: int = 8765,
        auth_token: Optional[str] = None,
        max_file_size_bytes: int = DEFAULT_MAX_PPTX_BYTES,
        ssl_certfile: Optional[str] = None,
        ssl_keyfile: Optional[str] = None,
    ):
        self.host = host
        self.port = port
        self.auth_token = auth_token if auth_token is not None else os.environ.get("PMG_QA_WORKER_TOKEN")
        self.job_manager = WorkerJobManager(
            worker_identity=f"pmg-windows-worker@{host}:{port}",
            max_file_size_bytes=max_file_size_bytes,
        )
        self.httpd = ThreadingHTTPServer((self.host, self.port), _WorkerRequestHandler)
        setattr(self.httpd, "auth_token", self.auth_token)
        setattr(self.httpd, "job_manager", self.job_manager)
        self.scheme = "http"

        if ssl_certfile and os.path.exists(ssl_certfile):
            ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            ctx.load_cert_chain(certfile=ssl_certfile, keyfile=ssl_keyfile)
            self.httpd.socket = ctx.wrap_socket(self.httpd.socket, server_side=True)
            self.scheme = "https"

        self.actual_port = self.httpd.server_address[1]
        self._thread: Optional[threading.Thread] = None

    @property
    def url(self) -> str:
        return f"{self.scheme}://{self.host}:{self.actual_port}"

    def start_background(self) -> str:
        self._thread = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        self._thread.start()
        return self.url

    def serve_forever(self) -> None:
        self.httpd.serve_forever()

    def shutdown(self) -> None:
        try:
            self.httpd.shutdown()
            self.httpd.server_close()
        except Exception:
            pass
        if self._thread is not None:
            self._thread.join(timeout=3.0)


    def start(self, background: bool = False) -> str:
        if background:
            return self.start_background()
        self.serve_forever()
        return self.url

    def stop(self) -> None:
        self.shutdown()
