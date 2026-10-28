"""
Tests for PMG v1.5.1:
- Upgrade A: Windows PowerPoint Validation Worker (Server, Client, Auth, Security, Unreachable -> BLOCKED, Live Worker Playback)
- Upgrade B: 24-Case Animation Benchmark Suite (All 24 cases across Groups A, B, C, D + Live Playback on Cases 21 & 24)
- Upgrade C: Runtime Dependency Check, Provenance & Embedded `ppt/pmg_execution_manifest.json` in delivered PPTX
"""

import os
import json
import zipfile
import pytest

from pmg.version import VERSION, verify_runtime_dependencies
from pmg.models import ValidationStatus
from pmg.orchestrator import PMGOrchestrator
from pmg.worker import (
    check_worker_health,
    validate_safe_pptx_upload,
    redact_secrets,
    PMGValidationWorkerServer,
    PMGWorkerClient,
)
from benchmarks import BENCHMARK_CASES, BenchmarkRunner


def test_v1_5_1_version_and_runtime_dependencies():
    assert VERSION in ("1.5.1", "1.5.2", "1.6.0", "1.7.0", "1.8.0")
    deps = verify_runtime_dependencies(require_windows_com=True)
    assert deps["ok"] is True
    assert deps["engine_version"] in ("1.5.1", "1.5.2", "1.6.0", "1.7.0", "1.8.0")
    assert deps["ffmpeg_available"] is True
    assert deps["ffprobe_available"] is True

    health = check_worker_health()
    assert health.status == "healthy"
    assert health.powerpoint_available is True


def test_worker_security_controls_and_redaction(tmp_path):
    # 1. Secret redaction
    raw = {"token": "super_secret_123", "nested": {"Authorization": "Bearer abc", "safe": "ok"}}
    redacted = redact_secrets(raw)
    assert redacted["token"] == "***REDACTED***"
    assert redacted["nested"]["Authorization"] == "***REDACTED***"
    assert redacted["nested"]["safe"] == "ok"

    # 2. Reject macro-enabled .pptm extension
    pptm_file = str(tmp_path / "malicious.pptm")
    with open(pptm_file, "wb") as f:
        f.write(b"PK\x03\x04dummy")
    ok, msg = validate_safe_pptx_upload(pptm_file, original_filename="malicious.pptm")
    assert ok is False
    assert "SECURITY_REJECTED" in msg

    # 3. Reject ZIP with path traversal entry
    traversal_zip = str(tmp_path / "traversal.pptx")
    with zipfile.ZipFile(traversal_zip, "w") as zf:
        zf.writestr("ppt/presentation.xml", "<p:presentation/>")
        zf.writestr("../../evil.txt", "pwn")
    ok2, msg2 = validate_safe_pptx_upload(traversal_zip, original_filename="traversal.pptx")
    assert ok2 is False
    assert "Path traversal" in msg2

    # 4. Reject ZIP with embedded vbaProject.bin macro payload
    macro_zip = str(tmp_path / "macro_inside.pptx")
    with zipfile.ZipFile(macro_zip, "w") as zf:
        zf.writestr("ppt/presentation.xml", "<p:presentation/>")
        zf.writestr("ppt/vbaProject.bin", b"VBA_MACRO_PAYLOAD")
    ok3, msg3 = validate_safe_pptx_upload(macro_zip, original_filename="macro_inside.pptx")
    assert ok3 is False
    assert "Unsafe macro" in msg3


def test_worker_missing_or_unreachable_returns_blocked(tmp_path):
    # Unconfigured worker URL -> BLOCKED (never PASS)
    client_no_url = PMGWorkerClient(worker_url="", worker_token=None)
    st = client_no_url.get_remote_status()
    assert st["status"] == "blocked"
    assert st["reachable"] is False

    res_no_url = client_no_url.validate_remote(str(tmp_path / "nonexistent.pptx"))
    assert res_no_url.validation_status == ValidationStatus.BLOCKED.value

    # Unreachable port -> BLOCKED (never PASS)
    client_dead = PMGWorkerClient(worker_url="http://127.0.0.1:59997", worker_token="tok")
    res_dead = client_dead.validate_remote(__file__, timeout_sec=3)
    assert res_dead.validation_status == ValidationStatus.BLOCKED.value


def test_worker_server_client_end_to_end_and_embedded_manifest(tmp_path):
    """
    Starts a real PMGValidationWorkerServer on localhost with Bearer token authentication,
    verifies 401 rejection on wrong token, and runs a full remote worker validation pass
    through PMGOrchestrator(worker_mode='remote') verifying real PowerPoint playback +
    embedded `ppt/pmg_execution_manifest.json` in the delivered PPTX.
    """
    srv = PMGValidationWorkerServer(host="127.0.0.1", port=0, auth_token="pmg_test_secret_token")
    worker_url = srv.start_background()
    try:
        # 1. Wrong token must be rejected with HTTP 401 -> BLOCKED
        bad_client = PMGWorkerClient(worker_url=worker_url, worker_token="wrong_token")
        bad_status = bad_client.get_remote_status()
        assert bad_status["reachable"] is False
        assert bad_status.get("http_status") == 401

        # 2. Valid token health check must succeed
        good_client = PMGWorkerClient(worker_url=worker_url, worker_token="pmg_test_secret_token")
        good_status = good_client.get_remote_status()
        assert good_status["reachable"] is True
        assert good_status["status"] == "healthy"

        # 3. Run PMGOrchestrator in 'remote' worker mode against the running Windows Worker Server
        out_pptx = str(tmp_path / "remote_validated_case01.pptx")
        case01 = BENCHMARK_CASES[0]
        orchestrator = PMGOrchestrator(qa_root_dir=str(tmp_path / "qa_remote"))
        manifest = orchestrator.run_pipeline(
            output_pptx_path=out_pptx,
            builder_fn=lambda p: case01.builder_fn(p, str(tmp_path / "assets")),
            hold_duration_ms=900,
            max_repair_loops=2,
            require_video_playback=True,
            worker_mode="remote",
            worker_url=worker_url,
            worker_token="pmg_test_secret_token",
            run_id="test_remote_worker_e2e",
        )
        assert manifest.status == ValidationStatus.PASS
        assert manifest.validation_method == "remote_windows_worker"
        assert manifest.delivered_pptx is not None and os.path.exists(manifest.delivered_pptx)

        # 4. Verify Upgrade C embedded provenance manifest inside the delivered .pptx archive
        with zipfile.ZipFile(manifest.delivered_pptx, "r") as zf:
            assert "ppt/pmg_execution_manifest.json" in zf.namelist()
            embedded = json.loads(zf.read("ppt/pmg_execution_manifest.json").decode("utf-8"))
        assert embedded["engine_version"] in ("1.5.1", "1.5.2", "1.6.0", "1.7.0", "1.8.0")
        assert embedded["validation_method"] == "remote_windows_worker"
        assert embedded["final_status"] == "NOT_VERIFIED"
        assert embedded["certification_source"] == "external_companion_manifest"
        assert manifest.final_verification["playback_verified"] is True
    finally:
        srv.shutdown()


def test_24_case_animation_benchmark_suite(tmp_path):
    """
    Verifies all 24 benchmark cases across Groups A, B, C, and D in the Benchmark Suite,
    plus live PowerPoint playback execution on Case 21 (Hero Image Visibility with Outline)
    and Case 24 (Negative Detection + Auto-Repair).
    """
    assert len(BENCHMARK_CASES) == 24
    runner = BenchmarkRunner(benchmarks_dir=str(tmp_path / "benchmarks"))

    # 1. Verify all 24 cases build, validate, extract expected timelines, and write JSON/MD reports
    report_all = runner.run_all(worker_mode="static", hold_duration_ms=800)
    assert report_all["summary"]["total_cases"] == 24
    assert report_all["summary"]["passed_cases"] == 24
    assert report_all["summary"]["pass_rate"] == 100.0

    # Verify Case 24 pre-heal detection was True and repair loop healed it
    c24_static = next(c for c in report_all["cases"] if c["case_id"] == 24)
    assert c24_static["detected_broken_pre_heal"] is True
    assert c24_static["repair_loops_used"] >= 1

    # 2. Run live PowerPoint playback benchmark on Case 21 and Case 24
    live_rep = runner.run_all(case_ids=[21, 24], worker_mode="local", hold_duration_ms=900)
    assert live_rep["summary"]["total_cases"] == 2
    assert live_rep["summary"]["passed_cases"] == 2
    for c in live_rep["cases"]:
        assert c["final_status"] == "PASS"
        assert c["visibility_correctness"] == 1.0

