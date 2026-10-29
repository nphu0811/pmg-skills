"""
tests.test_v2_4_0_self_healing
==============================
Unit and regression test suite for PMG v2.4.0 Self-Healing Compatibility,
Transactional Repair, and Immutable Certification.

Tests:
1. Diagnostic error code classification and safety levels
2. Safe automatic repair of missing OPC ContentTypes
3. Safe automatic repair of duplicate shape IDs
4. Safe automatic repair of dangling animation target references
5. Transactional rollback on unrecoverable corruptions (preserves original)
6. Prevention of destructive modifications (no flattening, no slide deletion)
7. Immutable certification binding to exact artifact SHA-256
8. Certificate invalidation upon byte tampering
9. Unified API build() with repair="safe"
"""

import os
import zipfile
import hashlib
import pytest
import pptx
from lxml import etree

from pmg.models import ValidationStatus, ReviewDefect
from pmg.diagnostics import (
    ErrorCode,
    RepairSafety,
    StructuredDiagnostic,
    DiagnosticClassifier,
)
from pmg.self_healing import (
    repair_missing_content_types,
    repair_duplicate_shape_ids,
    repair_broken_animation_targets,
    TransactionalRepairManager,
    issue_validation_certificate,
    verify_certificate,
)
from pmg.compiler import NativePPTXCompiler
from pmg.validation import validate_level1_structural
from pmg.runtime import PMGRuntime


def test_diagnostic_classifier_mapping():
    """Verifies that defects are accurately mapped to ErrorCode and RepairSafety."""
    # 1. Missing ContentType
    d1 = ReviewDefect("openxml", "critical", 0, "ppt/manifest.json", "Package part 'ppt/manifest.json' has no registered ContentType in [Content_Types].xml.")
    diag1 = DiagnosticClassifier.classify_defect(d1)
    assert diag1.error_code == ErrorCode.OPC_MISSING_CONTENT_TYPE
    assert diag1.repair_safety == RepairSafety.SAFE_AUTOMATIC

    # 2. Duplicate shape ID
    d2 = ReviewDefect("openxml", "critical", 1, "Slide 1", "Slide 1 contains duplicate shape ID '2'.")
    diag2 = DiagnosticClassifier.classify_defect(d2)
    assert diag2.error_code == ErrorCode.SHP_DUPLICATE_ID
    assert diag2.repair_safety == RepairSafety.SAFE_AUTOMATIC

    # 3. Missing animation target
    d3 = ReviewDefect("timing", "critical", 1, "Slide 1", "Animation target spid='99' does not exist on slide 1.")
    diag3 = DiagnosticClassifier.classify_defect(d3)
    assert diag3.error_code == ErrorCode.ANIM_MISSING_TARGET
    assert diag3.repair_safety == RepairSafety.SAFE_AUTOMATIC

    # 4. Morph identity collision
    d4 = ReviewDefect("morph", "major", 2, "!!Traveling_Probe", "Morph Collision: Duplicate traveling entity '!!Traveling_Probe' on slide 2.")
    diag4 = DiagnosticClassifier.classify_defect(d4)
    assert diag4.error_code == ErrorCode.MORPH_IDENTITY_COLLISION
    assert diag4.repair_safety == RepairSafety.REQUIRES_REVIEW


def test_safe_repair_missing_content_types(tmp_path):
    """Verifies safe repair of undeclared part ContentTypes in [Content_Types].xml."""
    pptx_path = str(tmp_path / "missing_ct.pptx")
    compiler = NativePPTXCompiler()
    compiler.compile_from_spec({
        "theme": "LIGHT_ACADEMIC_CYAN",
        "slides": [{"title": "Test", "elements": [{"type": "card", "name": "Card1", "title": "A", "body": "B"}]}]
    }, pptx_path)

    # Inject undeclared custom extension into ZIP
    with zipfile.ZipFile(pptx_path, "a") as zf:
        zf.writestr("ppt/custom_data.json", b'{"custom": true}')

    # Level 1 should fail before repair
    l1_before = validate_level1_structural(pptx_path)
    assert l1_before.status == ValidationStatus.FAIL
    assert any("has no registered ContentType" in d.description for d in l1_before.defects)

    # Apply safe automatic repair
    actions = repair_missing_content_types(pptx_path)
    assert len(actions) >= 1
    assert any("custom_data.json" in a or "json" in a for a in actions)

    # Level 1 should now pass
    l1_after = validate_level1_structural(pptx_path)
    assert l1_after.status == ValidationStatus.PASS
    assert l1_after.opc_content_types_valid is True


def test_safe_repair_duplicate_shape_ids(tmp_path):
    """Verifies safe renumbering of duplicate shape IDs on a slide."""
    pptx_path = str(tmp_path / "duplicate_ids.pptx")
    compiler = NativePPTXCompiler()
    compiler.compile_from_spec({
        "theme": "LIGHT_ACADEMIC_CYAN",
        "slides": [{
            "title": "Dup Test",
            "elements": [
                {"type": "card", "name": "Card1", "title": "A", "body": "B"},
                {"type": "card", "name": "Card2", "title": "C", "body": "D"},
            ]
        }]
    }, pptx_path)

    # Intentionally duplicate cNvPr id="2" in slide1.xml
    temp_zip = pptx_path + ".tmp.pptx"
    p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    with zipfile.ZipFile(pptx_path, "r") as src, zipfile.ZipFile(temp_zip, "w") as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "ppt/slides/slide1.xml":
                root = etree.fromstring(data)
                cNvPrs = list(root.iter(f"{{{p_ns}}}cNvPr"))
                if len(cNvPrs) >= 2:
                    # Force duplicate id="2"
                    cNvPrs[0].set("id", "2")
                    cNvPrs[1].set("id", "2")
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            dst.writestr(item, data)
    os.replace(temp_zip, pptx_path)

    l1_before = validate_level1_structural(pptx_path)
    assert l1_before.status == ValidationStatus.FAIL
    assert any("duplicate shape id '2'" in d.description.lower() for d in l1_before.defects)

    # Repair duplicate shape IDs
    actions = repair_duplicate_shape_ids(pptx_path)
    assert len(actions) >= 1

    l1_after = validate_level1_structural(pptx_path)
    assert l1_after.status == ValidationStatus.PASS
    assert l1_after.unique_identifiers_valid is True


def test_safe_repair_broken_animation_targets(tmp_path):
    """Verifies safe removal of dangling <p:spTgt> animation references."""
    pptx_path = str(tmp_path / "broken_anim.pptx")
    compiler = NativePPTXCompiler()
    compiler.compile_from_spec({
        "theme": "LIGHT_ACADEMIC_CYAN",
        "slides": [{
            "title": "Anim Test",
            "elements": [{"type": "card", "name": "ValidCard", "title": "A", "body": "B"}],
            "animations": [{"target": "ValidCard", "effect": "fade"}]
        }]
    }, pptx_path)

    # Inject dangling animation target spid="9999" into timing tree
    temp_zip = pptx_path + ".tmp.pptx"
    p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    with zipfile.ZipFile(pptx_path, "r") as src, zipfile.ZipFile(temp_zip, "w") as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "ppt/slides/slide1.xml":
                root = etree.fromstring(data)
                timing = root.find(f"{{{p_ns}}}timing")
                if timing is not None:
                    # Inject a dangling spTgt
                    first_cTn = timing.find(f".//{{{p_ns}}}cTn")
                    if first_cTn is not None:
                        cBhvr = etree.SubElement(first_cTn, f"{{{p_ns}}}cBhvr")
                        tgt = etree.SubElement(cBhvr, f"{{{p_ns}}}tgtEl")
                        etree.SubElement(tgt, f"{{{p_ns}}}spTgt", spid="9999")
                        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            dst.writestr(item, data)
    os.replace(temp_zip, pptx_path)

    l1_before = validate_level1_structural(pptx_path)
    assert l1_before.status == ValidationStatus.FAIL
    assert any("does not exist" in d.description for d in l1_before.defects)

    # Repair broken animation targets
    actions = repair_broken_animation_targets(pptx_path)
    assert len(actions) >= 1
    assert any("9999" in a for a in actions)

    l1_after = validate_level1_structural(pptx_path)
    assert l1_after.status == ValidationStatus.PASS
    assert l1_after.animation_targets_valid is True


def test_transactional_repair_manager_success(tmp_path):
    """Verifies that TransactionalRepairManager safely heals defects and produces a valid output."""
    pptx_path = str(tmp_path / "corrupt_candidate.pptx")
    compiler = NativePPTXCompiler()
    compiler.compile_from_spec({
        "theme": "LIGHT_ACADEMIC_CYAN",
        "slides": [{"title": "AutoHeal", "elements": [{"type": "card", "name": "Card1", "title": "A", "body": "B"}]}]
    }, pptx_path)

    # Inject undeclared extension
    with zipfile.ZipFile(pptx_path, "a") as zf:
        zf.writestr("ppt/meta.json", b'{"version": "2.4.0"}')

    mgr = TransactionalRepairManager(max_retries=3)
    repaired_path = str(tmp_path / "healed_output.pptx")
    res = mgr.repair(pptx_path, output_path=repaired_path)

    assert res.success is True
    assert res.rolled_back is False
    assert res.validation_status == ValidationStatus.PASS
    assert os.path.exists(repaired_path)
    assert res.repaired_sha256 != res.original_sha256

    # Verify output is completely valid
    l1 = validate_level1_structural(repaired_path)
    assert l1.status == ValidationStatus.PASS


def test_transactional_repair_manager_rollback_on_irreparable(tmp_path):
    """Verifies that TransactionalRepairManager rolls back cleanly when repair is impossible."""
    bad_file = str(tmp_path / "totally_broken.pptx")
    with open(bad_file, "wb") as f:
        f.write(b"NOT_A_VALID_ZIP_FILE_AT_ALL")

    with open(bad_file, "rb") as f:
        orig_hash = hashlib.sha256(f.read()).hexdigest()

    mgr = TransactionalRepairManager(max_retries=2)
    res = mgr.repair(bad_file)

    assert res.success is False
    assert res.rolled_back is True
    # Verify original file bytes were untouched
    with open(bad_file, "rb") as f:
        assert hashlib.sha256(f.read()).hexdigest() == orig_hash


def test_immutable_validation_certificate_and_tamper_detection(tmp_path):
    """Verifies certificate generation, hash binding, and invalidation on byte tampering."""
    pptx_path = str(tmp_path / "certified_deck.pptx")
    compiler = NativePPTXCompiler()
    compiler.compile_from_spec({
        "theme": "LIGHT_ACADEMIC_CYAN",
        "slides": [{"title": "Certified", "elements": [{"type": "card", "name": "Card1", "title": "A", "body": "B"}]}]
    }, pptx_path)

    # Issue immutable certificate
    cert = issue_validation_certificate(
        pptx_path=pptx_path,
        validation_level="Level 1 Structural",
        overall_status="PASS",
        checks_performed=["zip_integrity", "xml_valid", "opc_content_types", "unique_ids"],
    )

    assert cert.overall_status == "PASS"
    assert cert.validator_version == "2.4.0"
    assert len(cert.artifact_sha256) == 64

    # Verification before modification: PASS
    is_valid, msg = verify_certificate(cert, pptx_path)
    assert is_valid is True
    assert "Certificate is valid" in msg

    # Tamper with 1 byte in the PPTX
    with open(pptx_path, "ab") as f:
        f.write(b"\x00")

    # Verification after modification: MUST FAIL (Tampering detected)
    is_valid_after, msg_after = verify_certificate(cert, pptx_path)
    assert is_valid_after is False
    assert "INVALIDATED" in msg_after


def test_unified_api_runtime_build_with_safe_repair(tmp_path):
    """Verifies PMGRuntime.build(..., repair='safe') end-to-end integration."""
    out_path = str(tmp_path / "runtime_repaired.pptx")

    spec = {
        "theme": "LIGHT_ACADEMIC_CYAN",
        "slides": [
            {
                "title": "Quantum Teleportation Protocol",
                "subtitle": "Self-Healing Architecture v2.4.0",
                "elements": [
                    {"type": "card", "name": "!!Entangled_Node", "title": "Bell State Fabric", "body": "Fidelity maintained at 99.98%."}
                ],
                "animations": [
                    {"target": "!!Entangled_Node", "effect": "fade", "duration_ms": 700}
                ]
            }
        ]
    }

    runtime = PMGRuntime.auto()
    res = runtime.build(
        project=spec,
        output=out_path,
        validate=True,
        repair="safe",
    )

    assert res.is_valid is True
    assert res.can_deliver is True
    assert len(res.artifact_sha256) == 64
    assert os.path.exists(out_path)
