"""
tests/test_v1_4_reverse_engineer.py
===================================
Automated test suite for PMG Skills v1.4.0:
- PPTX Motion Reverse Engineering (`PPTXMotionReverseEngineer`)
- Extraction of object hierarchy, geometry, z-order, roles, transitions, and native <p:timing> animations
- Animation style transfer onto new content and palette (`reconstruct_and_transfer`)
- Exact vs. approximated motion comparison (`compare_reference_and_reconstructed`)
- Native PowerPoint Animation Pane COM verification on reconstructed slide
"""

import os
from pmg.reverse_engineer import (
    PPTXMotionReverseEngineer,
    build_v1_4_reference_smartphone_slide,
)
from pmg.validator import StaticValidator
from pmg.animations import verify_animation_pane_com


def test_v1_4_extract_and_reconstruct_smartphone_slide(tmp_path):
    ref_path = str(tmp_path / "ref_smartphone.pptx")
    recon_path = str(tmp_path / "recon_solar_tablet.pptx")
    spec_json = str(tmp_path / "extracted_spec.json")

    # 1. Build reference smartphone slide
    build_v1_4_reference_smartphone_slide(ref_path)
    assert os.path.exists(ref_path)

    # 2. Extract motion specification
    re_engine = PPTXMotionReverseEngineer()
    spec = re_engine.extract_motion_spec(ref_path)
    spec.save(spec_json)
    assert os.path.exists(spec_json)
    assert len(spec.slides) == 1

    s1 = spec.slides[0]
    assert len(s1.objects) == 13
    assert len(s1.animations) == 15

    # Verify extracted effect types include fly_in, rotate, scale, motion_path, wipe, fade_out
    eff_types = {a.effect_type for a in s1.animations}
    for req in ("appear", "fade", "fly_in", "rotate", "scale", "motion_path", "wipe", "fade_out"):
        assert req in eff_types, f"Missing extracted effect type: {req}"

    # 3. Reconstruct and transfer animation principles onto new content
    re_engine.reconstruct_and_transfer(
        motion_spec=spec,
        output_pptx_path=recon_path,
        content_transfer={
            "Ref_Title": "SolarGrid Field Tablet · Edge Telemetry\nTransferred motion principles from EcoPhone reference.",
            "Hero_Smartphone_Frame": "SOLAR-TAB OS v5",
            "Smartphone_Display_Screen": "MICROGRID SYNC: 100%",
        },
        palette_transfer={
            "hero_device": "064E3B",
        },
    )
    assert os.path.exists(recon_path)

    # 4. Validate reconstructed PPTX statically
    validator = StaticValidator()
    val = validator.validate(recon_path)
    assert val["is_valid"] is True
    assert val["slides"][0]["native_animation_count"] == 15

    # 5. Compare reference and reconstructed motion
    comp = re_engine.compare_reference_and_reconstructed(ref_path, recon_path)
    assert comp["reference_object_count"] == 13
    assert comp["reconstructed_object_count"] == 13
    assert comp["reference_animation_count"] == 15
    assert comp["reconstructed_animation_count"] == 15
    assert comp["exact_animation_matches"] == 15
    assert comp["approximated_animations"] == 0
    assert comp["mean_position_error_in"] == 0.0
    assert comp["mean_duration_error_ms"] == 0.0

    # 6. Verify reconstructed slide in PowerPoint Animation Pane via COM
    com_info = verify_animation_pane_com(recon_path)
    if com_info.get("verified_via_com"):
        assert com_info["slides"][1]["main_sequence_count"] == 15
