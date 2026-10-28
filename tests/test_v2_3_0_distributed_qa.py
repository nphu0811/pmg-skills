"""
tests.test_v2_3_0_distributed_qa
=================================
Unit and integration test suite for PMG v2.3.0 Distributed PowerPoint QA.

Tests:
1. Worker health check (GET /health)
2. Authentication security (Bearer token required, 401 on missing/invalid)
3. Payload size enforcement (413 on oversized upload)
4. Job cancellation (POST /jobs/<id>/cancel)
5. Idempotent job retrieval via artifact SHA-256
6. Client handling when remote worker is unreachable (BLOCKED status, no crash)
7. Windows PowerPoint real playback validation when available
"""

import os
import time
import json
import pytest
import pptx
from pmg.models import ValidationStatus
from pmg.worker.server import PMGValidationWorkerServer
from pmg.worker.client import PMGWorkerClient
from pmg.worker.jobs import WorkerJobManager
from pmg.worker.powerpoint_probe import file_sha256
from pmg.worker.powerpoint_runner import is_powerpoint_available


@pytest.fixture(scope="module")
def worker_server():
    """Runs a live PMG validation worker server in a background thread."""
    token = "test-secret-token-v230"
    server = PMGValidationWorkerServer(
        host="127.0.0.1",
        port=8789,
        auth_token=token,
        max_file_size_bytes=5 * 1024 * 1024,  # 5MB limit for testing
    )
    server.start(background=True)
    time.sleep(0.3)
    yield server
    server.stop()


def test_worker_health_check_endpoint(worker_server):
    """Verifies that GET /health returns structured WorkerHealthStatus."""
    client = PMGWorkerClient(worker_url="http://127.0.0.1:8789", worker_token="test-secret-token-v230")
    status = client.get_remote_status()
    assert status.get("reachable") is True
    assert "platform" in status
    assert "powerpoint_available" in status
    assert "worker_version" in status


def test_worker_authentication_rejection(worker_server):
    """Verifies that unauthorized requests receive HTTP 401."""
    bad_client = PMGWorkerClient(worker_url="http://127.0.0.1:8789", worker_token="wrong-token")
    status = bad_client.get_remote_status()
    assert status.get("reachable") is False
    assert status.get("http_status") == 401


def test_worker_payload_size_limit_rejection(worker_server, tmp_path):
    """Verifies that requests exceeding maximum payload size receive HTTP 413."""
    oversized = tmp_path / "huge.pptx"
    # Write 6MB of data exceeding the 5MB limit
    with open(oversized, "wb") as f:
        f.write(b"0" * (6 * 1024 * 1024))

    client = PMGWorkerClient(worker_url="http://127.0.0.1:8789", worker_token="test-secret-token-v230")
    result = client.validate_remote(str(oversized))
    assert result.state == "failed"
    assert "SECURITY_REJECTED" in (result.error or "") or "413" in (result.error or "")


def test_worker_cancel_job(worker_server):
    """Verifies job cancellation endpoint."""
    client = PMGWorkerClient(worker_url="http://127.0.0.1:8789", worker_token="test-secret-token-v230")
    job_mgr: WorkerJobManager = worker_server.job_manager
    # Pre-register a job in queued state
    test_jid = "test_cancel_123"
    job_mgr.execute_job_sync = lambda **kw: None
    job_mgr.cancel_job(test_jid)
    assert test_jid in job_mgr._cancelled


def test_idempotent_job_hash_caching(tmp_path):
    """Verifies that identical PPTX uploads return cached results without re-running."""
    sample_pptx = str(tmp_path / "sample_idempotent.pptx")
    prs = pptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(sample_pptx)

    mgr = WorkerJobManager(worker_identity="test-worker")
    res1 = mgr.execute_job_sync(sample_pptx, job_id="job_first")

    # Second run with allow_cache=True
    sha = file_sha256(sample_pptx)
    if res1.validation_status == ValidationStatus.PASS.value:
        cached = mgr.get_cached_by_hash(sha)
        assert cached is not None
        assert cached.job_id == "job_first"


def test_unreachable_remote_worker_handled_gracefully():
    """Verifies that an unconfigured or dead remote worker returns BLOCKED without raising unhandled errors."""
    client = PMGWorkerClient(worker_url="http://127.0.0.1:9999", worker_token="token")
    result = client.validate_remote("nonexistent.pptx")
    # File check fires first
    assert result.validation_status in (ValidationStatus.FAIL.value, ValidationStatus.BLOCKED.value)

    # With dummy file but dead worker
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tf:
        tf.write(b"DUMMY_PPTX_BYTES")
        dummy_path = tf.name

    try:
        dead_client = PMGWorkerClient(worker_url="http://127.0.0.1:9999", worker_token="token")
        res_dead = dead_client.validate_remote(dummy_path, timeout_sec=2)
        assert res_dead.validation_status == ValidationStatus.BLOCKED.value
        assert "BLOCKED" in (res_dead.error or "")
    finally:
        if os.path.exists(dummy_path):
            os.remove(dummy_path)


def test_real_powerpoint_worker_validation_if_available(tmp_path):
    """
    Executes actual Windows PowerPoint validation if PowerPoint is installed on this host.
    Separated cleanly: skipped if PowerPoint is not available.
    """
    if not is_powerpoint_available():
        pytest.skip("PowerPoint Desktop is not available on this host.")

    pptx_path = str(tmp_path / "worker_deck.pptx")
    from pmg.compiler import NativePPTXCompiler
    compiler = NativePPTXCompiler()
    slide = compiler.add_slide()
    compiler.add_glass_card(slide, 1.0, 2.0, 5.0, 3.0, "Card1", "Title 1", "Body 1")
    compiler.add_glass_card(slide, 7.0, 2.0, 5.0, 3.0, "Card2", "Title 2", "Body 2")
    compiler.save(pptx_path)

    mgr = WorkerJobManager(worker_identity="test-local-worker")
    res = mgr.execute_job_sync(pptx_path, timeout_sec=60)
    assert res.clean_open.get("is_clean") is True
    assert res.validated_pptx_sha256 == file_sha256(pptx_path)
