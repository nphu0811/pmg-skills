"""
Automated Regression and Acceptance Tests for PMG Skills v1.5.2:
1. Valid OPC ContentType Declaration for Embedded Manifests:
   - Ensures ppt/pmg_execution_manifest.json is registered in [Content_Types].xml
   - Rejects unmapped parts that could trigger PowerPoint Repair dialog
2. External Companion Manifest Output:
   - Ensures <pptx_path>.manifest.json is generated
3. Strict Delivery Gate:
   - Re-validation of final deliverable post-packaging
   - Clean PowerPoint COM opening test (verify_clean_powerpoint_open)
4. Cinematic Showcase Presets:
   - Spotify Radial App Shell & UI Stream (SpotifyAppShellBuilder)
   - Earth & Space Viewfinder / Targeting Reticle Sequence (EarthSpaceMotionBuilder)
   - University Concentric Blooming Rings & 8-Petal Mục Lục (UniversityBloomBuilder)
   - Universal topic adaptability across light and dark palettes
"""

import os
import json
import zipfile
import pytest
import xml.etree.ElementTree as ET

from pmg import (
    VERSION,
    PMGOrchestrator,
    StaticValidator,
    Themes,
    MotionDeck,
)
from pmg.orchestrator import _embed_manifest_in_pptx
from pmg.validator import verify_clean_powerpoint_open
from pmg.cinematic_showcase import (
    SpotifyAppShellBuilder,
    EarthSpaceMotionBuilder,
    UniversityBloomBuilder,
)


def test_current_reliability_version():
    assert VERSION == "1.7.0"


def test_spotify_app_shell_builder_generation(tmp_path):
    output_pptx = str(tmp_path / "test_spotify.pptx")
    builder = SpotifyAppShellBuilder(
        topic_title="AI MUSIC STUDIO",
        topic_subtitle="NEXT-GEN SOUND PLATFORM",
        author_name="Audio Team",
    )
    res_path = builder.build_deck(output_pptx, num_slides=4)
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 1000

    val_report = StaticValidator().validate(res_path)
    assert val_report["is_valid"], f"Validation errors: {val_report.get('defects')}"
    morph_count = sum(1 for s in val_report["slides"] if s.get("transition_type") == "morph")
    assert morph_count >= 2
    assert any("AppShell" in s for s in val_report["traveling_anchors"])


def test_earth_space_motion_builder_generation(tmp_path):
    output_pptx = str(tmp_path / "test_earth_space.pptx")
    builder = EarthSpaceMotionBuilder(
        topic_title="EXPLORATION SYSTEM",
        topic_subtitle="AUTONOMOUS ORBITAL MISSIONS",
    )
    res_path = builder.build_deck(output_pptx, num_slides=5)
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 1000

    val_report = StaticValidator().validate(res_path)
    assert val_report["is_valid"], f"Validation errors: {val_report.get('defects')}"
    morph_count = sum(1 for s in val_report["slides"] if s.get("transition_type") == "morph")
    assert morph_count >= 3
    assert any("Reticle" in s for s in val_report["traveling_anchors"])


def test_university_bloom_builder_generation(tmp_path):
    output_pptx = str(tmp_path / "test_bloom.pptx")
    builder = UniversityBloomBuilder(
        presentation_title="HOA HỌC XANH",
        subtitle="ĐẠI HỌC QUỐC GIA",
        student_name="Nhóm Nghiên Cứu",
    )
    res_path = builder.build_deck(output_pptx)
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 1000

    val_report = StaticValidator().validate(res_path)
    assert val_report["is_valid"], f"Validation errors: {val_report.get('defects')}"
    morph_count = sum(1 for s in val_report["slides"] if s.get("transition_type") == "morph")
    assert morph_count >= 2
    assert any("Petal" in s for s in val_report["traveling_anchors"])


def test_manifest_content_type_registration(tmp_path):
    output_pptx = str(tmp_path / "test_manifest_opc.pptx")
    deck = MotionDeck(theme=Themes.LIGHT_ACADEMIC_CYAN)
    s1 = deck.add_slide(transition="none")
    s2 = deck.add_slide(transition="morph", duration_ms=1500)
    deck.save(output_pptx)

    dummy_manifest = {
        "engine_version": "1.7.0",
        "status": "PASS",
        "run_id": "test_opc",
    }
    _embed_manifest_in_pptx(output_pptx, dummy_manifest)

    # 1. Verify companion manifest was generated externally
    companion_path = f"{output_pptx}.manifest.json"
    assert os.path.exists(companion_path)
    with open(companion_path, "r", encoding="utf-8") as f:
        comp_data = json.load(f)
    assert comp_data["engine_version"] == "1.7.0"

    # 2. Verify OPC [Content_Types].xml has valid ContentType mapping for json
    with zipfile.ZipFile(output_pptx, "r") as zf:
        assert "[Content_Types].xml" in zf.namelist()
        ct_bytes = zf.read("[Content_Types].xml")
        root = ET.fromstring(ct_bytes)
        default_ns = "http://schemas.openxmlformats.org/package/2006/content-types"

        has_json_mapping = False
        for child in root.findall(f"{{{default_ns}}}Default"):
            ext = child.attrib.get("Extension", "").lower()
            ct = child.attrib.get("ContentType", "").lower()
            if ext == "json" and "json" in ct:
                has_json_mapping = True
                break
        assert has_json_mapping, "Expected [Content_Types].xml to have <Default Extension='json' ContentType='application/json'/>"

    # 3. Verify static validator flags zero unmapped part defects
    val_report = StaticValidator().validate(output_pptx)
    unmapped_defects = [d for d in val_report["defects"] if d.get("remediation_hint") == "declare_part_content_type"]
    assert len(unmapped_defects) == 0, f"Unmapped part defects found: {unmapped_defects}"


def test_unmapped_part_triggers_critical_defect(tmp_path):
    output_pptx = str(tmp_path / "test_unmapped.pptx")
    deck = MotionDeck(theme=Themes.LIGHT_ACADEMIC_CYAN)
    s1 = deck.add_slide(transition="none")
    deck.save(output_pptx)

    # Corrupt by injecting a file with unknown extension into the ZIP without registering in [Content_Types].xml
    temp_zip = str(tmp_path / "corrupted.pptx")
    with zipfile.ZipFile(output_pptx, "r") as zin:
        with zipfile.ZipFile(temp_zip, "w") as zout:
            for item in zin.infolist():
                zout.writestr(item, zin.read(item.filename))
            zout.writestr("custom_data/unknown.xyz", b"hello corrupt")

    val_report = StaticValidator().validate(temp_zip)
    unmapped_defects = [d for d in val_report["defects"] if d.get("remediation_hint") == "declare_part_content_type"]
    assert len(unmapped_defects) > 0, "Expected StaticValidator to detect unmapped OPC part"


def test_verify_clean_powerpoint_open_on_clean_deck(tmp_path):
    output_pptx = str(tmp_path / "test_clean_com.pptx")
    deck = MotionDeck(theme=Themes.LIGHT_ACADEMIC_CYAN)
    deck.add_slide(transition="none")
    deck.add_slide(transition="morph", duration_ms=1000)
    deck.save(output_pptx)

    # Embed valid manifest
    _embed_manifest_in_pptx(output_pptx, {"version": "1.7.0", "status": "PASS"})

    com_result = verify_clean_powerpoint_open(output_pptx)
    assert com_result["is_clean"] is True, f"verify_clean_powerpoint_open failed: {com_result.get('error')}"
