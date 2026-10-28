"""
tests/test_v1_5_video_reconstruct.py
====================================
Automated test suite for PMG Skills v1.5.0:
- Video-to-Editable-PowerPoint Motion Reconstruction (`VideoMotionReconstructor`)
- 12-stage CV & motion graphics pipeline (scene boundaries, keyframes, object detection,
  multi-frame trajectory tracking, cubic Bezier motion path fitting, editable shape synthesis,
  and native <p:timing> Animation Pane compilation)
- Anti-shortcut enforcement (0 flattened bitmap frames; 100% editable shapes/text)
- Engineering honesty report verification (`directly_observed_motion`, `estimated_motion`,
  `inferred_visual_properties`, `unsupported_effects`, `approximate_reconstructions`)
"""

import os
from pmg.video_reconstructor import (
    VideoMotionReconstructor,
    generate_controlled_reference_video,
)
from pmg.validator import StaticValidator


def test_v1_5_controlled_video_to_editable_pptx_reconstruction(tmp_path):
    ref_mp4 = str(tmp_path / "controlled_ref.mp4")
    recon_pptx = str(tmp_path / "recon_from_video.pptx")
    work_dir = str(tmp_path / "recon_work")

    # 1. Generate controlled 6.5s animated reference video fixture
    generate_controlled_reference_video(ref_mp4, duration_sec=6.5, fps=30)
    assert os.path.exists(ref_mp4)

    # 2. Run 12-stage Video-to-Editable-PowerPoint reconstruction
    reconstructor = VideoMotionReconstructor(sample_fps=10.0, max_duration_sec=6.5, max_scenes=1)
    report = reconstructor.reconstruct_from_video(
        video_path=ref_mp4,
        output_pptx_path=recon_pptx,
        work_dir=work_dir,
    )
    assert os.path.exists(recon_pptx)
    assert report.scenes_detected >= 1
    assert report.total_editable_objects_reconstructed >= 6
    assert report.total_native_animations_reconstructed >= 6

    # 3. Verify engineering honesty categories in report
    assert len(report.directly_observed_motion) >= 5
    assert len(report.estimated_motion) >= 5
    assert len(report.inferred_visual_properties) >= 5
    assert len(report.unsupported_effects) >= 1
    assert len(report.approximate_reconstructions) >= 5

    # 4. Verify StaticValidator (100% editable shapes, 0 picture frame-flattening)
    validator = StaticValidator()
    val = validator.validate(recon_pptx)
    assert val["is_valid"] is True
    for s in val["slides"]:
        assert s["picture_count"] == 0, "Forbidden frame-flattening detected"
        assert s["editable_object_count"] >= 6
        assert s["native_animation_count"] >= 6

    # 5. Verify PowerPoint Animation Pane via COM
    if report.com_animation_pane_verified:
        assert report.com_animation_pane_count == report.total_native_animations_reconstructed


def test_v1_5_real_user_video_reconstruction_if_present(tmp_path):
    user_video = r"C:\Users\admin\Videos\2026-10-28 09-42-27.mp4"
    if not os.path.exists(user_video):
        return

    recon_pptx = str(tmp_path / "user_video_recon.pptx")
    work_dir = str(tmp_path / "user_video_work")
    reconstructor = VideoMotionReconstructor(sample_fps=8.0, max_duration_sec=6.5, max_scenes=1)
    report = reconstructor.reconstruct_from_video(
        video_path=user_video,
        output_pptx_path=recon_pptx,
        work_dir=work_dir,
    )
    assert os.path.exists(recon_pptx)
    assert report.total_editable_objects_reconstructed >= 4
    assert report.total_native_animations_reconstructed >= 4

    validator = StaticValidator()
    val = validator.validate(recon_pptx)
    assert val["is_valid"] is True
    assert val["slides"][0]["picture_count"] == 0

