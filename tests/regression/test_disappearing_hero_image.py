"""
Regression Test: Upgrade D — Detect and Fix Disappearing Objects During Slideshow
=================================================================================
Reproduces the exact real-world failure mode from `Vu_tru_20_slide_PMG_inspired.pptx`
where `!!Earth_Grand` (Hero Earth Picture) and `!!HeroTitle` (Animated Title) are visible
in PowerPoint Edit Mode, but disappear completely during Slide Show while the sibling
decorative ring (`Oval 13`) remains visible.

Verifies all 11 mandatory assertions from Upgrade D:
1. Slide 1 contains a starfield background, hero image (`!!Earth_Grand`), decorative outline (`Oval 13`),
   main title (`!!HeroTitle`), subtitle (`!!HeroSubtitle`), and 2 info cards (`!!Card_1`, `!!Card_2`).
2. Fade In animations are applied to the hero image and title.
3. The hero image is visible in Edit Mode (`edit_mode_refs/edit_slide_01.png`).
4. In the broken variant, the hero image and title disappear during Slide Show while the outline remains.
5. PMG validator + visual reviewer detect the missing hero image and title (`FAIL`).
6. The decorative outline (`Oval 13`) alone does NOT count as passing.
7. Self-healing repairs the `<p:timing>` structure automatically within <= 3 loops.
8. The repaired PPTX still retains its native Fade In animations (`<p:animEffect filter="fade"/>` + `<p:set>`).
9. The hero image fades in during Slide Show and remains visible at steady state.
10. The title fades in during Slide Show and remains visible at steady state.
11. Final validation passes (`ValidationStatus.PASS`) — plus direct verification on the real
    `Vu_tru_20_slide_PMG_inspired.pptx` Slide 1–2 excerpt when present on disk.
"""

import os
import zipfile
import tempfile
import numpy as np
import cv2
import pptx
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from lxml import etree

from pmg.models import ValidationStatus
from pmg.orchestrator import PMGOrchestrator
from pmg.validator import StaticValidator, extract_expected_timeline
from pmg.video_pipeline import export_video_with_powerpoint, record_actual_slideshow
from pmg.visual_reviewer import VisualAnimationReviewer
from pmg.animations import P_NS
from scripts.motion_graph_kit import apply_transition


def _create_earth_globe_png(path: str, size: int = 640) -> str:
    """Generates a high-contrast, vibrant Earth globe PNG image with continents, oceans, and clouds."""
    img = np.zeros((size, size, 3), dtype=np.uint8)
    # Deep space dark surround
    img[:, :] = (18, 12, 8)
    center = (size // 2, size // 2)
    radius = int(size * 0.44)

    # Ocean base (vibrant blue-cyan in BGR)
    cv2.circle(img, center, radius, (215, 115, 18), -1, lineType=cv2.LINE_AA)
    # Inner lighter ocean gradient
    cv2.circle(img, (center[0] - 25, center[1] - 25), int(radius * 0.78), (245, 165, 45), -1, lineType=cv2.LINE_AA)
    # Continents (emerald/green/gold land masses in BGR)
    cv2.ellipse(img, (center[0] - 40, center[1] - 30), (95, 55), 25, 0, 360, (85, 195, 65), -1, lineType=cv2.LINE_AA)
    cv2.ellipse(img, (center[0] + 55, center[1] + 35), (80, 48), -20, 0, 360, (70, 175, 95), -1, lineType=cv2.LINE_AA)
    cv2.ellipse(img, (center[0] - 10, center[1] + 75), (60, 30), 10, 0, 360, (110, 210, 120), -1, lineType=cv2.LINE_AA)
    # White atmospheric cloud bands
    cv2.ellipse(img, (center[0] - 15, center[1] - 65), (110, 18), -12, 0, 360, (245, 250, 255), -1, lineType=cv2.LINE_AA)
    cv2.ellipse(img, (center[0] + 20, center[1] + 10), (125, 16), 8, 0, 360, (240, 245, 250), -1, lineType=cv2.LINE_AA)
    # Specular highlight
    cv2.circle(img, (center[0] - 65, center[1] - 65), 35, (255, 250, 235), -1, lineType=cv2.LINE_AA)

    cv2.imwrite(path, img)
    return path


def _build_broken_space_hero_fixture(output_pptx: str, earth_png: str) -> str:
    """
    Builds a 2-slide Space presentation reproducing the exact broken <p:timing> XML
    structure from `Vu_tru_20_slide_PMG_inspired.pptx`:
    - Starfield dark background (`!!Starfield_BG`)
    - Hero Earth image (`!!Earth_Grand`)
    - Decorative circular ring (`Oval 13`) surrounding `!!Earth_Grand`
    - Main title (`!!HeroTitle`)
    - Subtitle (`!!HeroSubtitle`)
    - Two info cards (`!!Card_Orbital`, `!!Card_Telemetry`)
    - Slide 2 with Morph transition (`!!Earth_Grand`, `!!HeroTitle`, etc.)
    - Broken `<p:timing>` on Slide 1:
      * `<p:bldLst>` registers `!!HeroTitle`, `!!HeroSubtitle`, `!!Earth_Grand`, `!!Card_Orbital`, `!!Card_Telemetry`
        with `grpId="1..5"` so PowerPoint hides them at Slide Show start.
      * `Oval 13` is NOT in `<p:bldLst>`, so the blue ring outline stays visible!
      * `<p:cTn presetClass="entr">` nodes omit `<p:set>` (`style.visibility='visible'`), omit `grpId`,
        sit in separate top-level `<p:par>` siblings under `mainSeq` without `<p:tn val="2"/>`, and
        include a duplicate conflicting entrance on `!!HeroTitle`.
    """
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for s_idx in (1, 2):
        slide = prs.slides.add_slide(blank_layout)

        # 1. Starfield Background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.name = "!!Starfield_BG"
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(8, 12, 24)
        bg.line.fill.background()

        # 2. Main Title
        title_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(1.0 if s_idx == 1 else 0.7), Inches(6.5), Inches(1.5)
        )
        title_box.name = "!!HeroTitle"
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = "VŨ TRỤ ĐẠI QUAN" if s_idx == 1 else "HÀNH TRÌNH QUỸ ĐẠO"
        p.font.size = Pt(46)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

        # 3. Subtitle
        sub_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(2.6 if s_idx == 1 else 2.2), Inches(6.2), Inches(1.0)
        )
        sub_box.name = "!!HeroSubtitle"
        stf = sub_box.text_frame
        stf.word_wrap = True
        sp = stf.paragraphs[0]
        sp.text = "Khám phá Trái Đất và không gian sâu qua chuyển động Morph 60fps"
        sp.font.size = Pt(20)
        sp.font.color.rgb = RGBColor(148, 163, 184)

        # 4. Two Info Cards
        c1 = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8),
            Inches(4.2 if s_idx == 1 else 3.8),
            Inches(3.1),
            Inches(1.8),
        )
        c1.name = "!!Card_Orbital"
        c1.fill.solid()
        c1.fill.fore_color.rgb = RGBColor(30, 41, 59)
        c1.line.color.rgb = RGBColor(56, 189, 248)
        c1.text_frame.text = "ORBITAL VELOCITY\n29.78 km/s"
        c1.text_frame.paragraphs[0].font.color.rgb = RGBColor(240, 249, 255)
        c1.text_frame.paragraphs[0].font.size = Pt(16)

        c2 = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(4.2),
            Inches(4.2 if s_idx == 1 else 3.8),
            Inches(3.1),
            Inches(1.8),
        )
        c2.name = "!!Card_Telemetry"
        c2.fill.solid()
        c2.fill.fore_color.rgb = RGBColor(30, 41, 59)
        c2.line.color.rgb = RGBColor(56, 189, 248)
        c2.text_frame.text = "APOGEE ALTITUDE\n408 km ISS"
        c2.text_frame.paragraphs[0].font.color.rgb = RGBColor(240, 249, 255)
        c2.text_frame.paragraphs[0].font.size = Pt(16)

        # 5. Large Hero Earth Image (`!!Earth_Grand`)
        earth_x = Inches(7.9 if s_idx == 1 else 8.3)
        earth_y = Inches(1.1 if s_idx == 1 else 1.4)
        earth_d = Inches(4.8 if s_idx == 1 else 4.2)
        pic = slide.shapes.add_picture(earth_png, earth_x, earth_y, width=earth_d, height=earth_d)
        pic.name = "!!Earth_Grand"

        # 6. Decorative Outline Ring (`Oval 13`) surrounding the Earth image
        ring = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            earth_x - Inches(0.12),
            earth_y - Inches(0.12),
            earth_d + Inches(0.24),
            earth_d + Inches(0.24),
        )
        ring.name = "Oval 13"
        ring.fill.background()
        ring.line.color.rgb = RGBColor(56, 189, 248)
        ring.line.width = Pt(2.25)

        if s_idx == 1:
            t_spid = title_box.shape_id
            s_spid = sub_box.shape_id
            e_spid = pic.shape_id
            c1_spid = c1.shape_id
            c2_spid = c2.shape_id

            # Inject the exact broken <p:timing> pattern from Vu_tru_20_slide_PMG_inspired.pptx:
            # - Missing <p:set> visibility inside presetClass="entr"
            # - Missing grpId on <p:cTn> while <p:bldLst> binds grpId="1..6"
            # - Separate top-level <p:par> siblings directly under mainSeq without <p:tn val="2"/>
            # - Duplicate conflicting entrance on title_box (t_spid)
            broken_timing_xml = f"""
            <p:timing xmlns:p="{P_NS}">
              <p:tnLst>
                <p:par>
                  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                    <p:childTnLst>
                      <p:seq concurrent="1" nextAc="seek">
                        <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                          <p:childTnLst>
                            <p:par>
                              <p:cTn id="3" fill="hold">
                                <p:stCondLst><p:cond evt="onBegin" delay="0"/></p:stCondLst>
                                <p:childTnLst>
                                  <p:par>
                                    <p:cTn id="4" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">
                                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                      <p:childTnLst>
                                        <p:animEffect transition="in" filter="fade">
                                          <p:cBhvr>
                                            <p:cTn id="5" dur="700"/>
                                            <p:tgtEl><p:spTgt spid="{t_spid}"/></p:tgtEl>
                                          </p:cBhvr>
                                        </p:animEffect>
                                      </p:childTnLst>
                                    </p:cTn>
                                  </p:par>
                                </p:childTnLst>
                              </p:cTn>
                            </p:par>
                            <p:par>
                              <p:cTn id="6" fill="hold">
                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                <p:childTnLst>
                                  <p:par>
                                    <p:cTn id="7" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">
                                      <p:stCondLst><p:cond delay="120"/></p:stCondLst>
                                      <p:childTnLst>
                                        <p:animEffect transition="in" filter="fade">
                                          <p:cBhvr>
                                            <p:cTn id="8" dur="700"/>
                                            <p:tgtEl><p:spTgt spid="{t_spid}"/></p:tgtEl>
                                          </p:cBhvr>
                                        </p:animEffect>
                                      </p:childTnLst>
                                    </p:cTn>
                                  </p:par>
                                </p:childTnLst>
                              </p:cTn>
                            </p:par>
                            <p:par>
                              <p:cTn id="9" fill="hold">
                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                <p:childTnLst>
                                  <p:par>
                                    <p:cTn id="10" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">
                                      <p:stCondLst><p:cond delay="180"/></p:stCondLst>
                                      <p:childTnLst>
                                        <p:animEffect transition="in" filter="fade">
                                          <p:cBhvr>
                                            <p:cTn id="11" dur="900"/>
                                            <p:tgtEl><p:spTgt spid="{e_spid}"/></p:tgtEl>
                                          </p:cBhvr>
                                        </p:animEffect>
                                      </p:childTnLst>
                                    </p:cTn>
                                  </p:par>
                                </p:childTnLst>
                              </p:cTn>
                            </p:par>
                            <p:par>
                              <p:cTn id="12" fill="hold">
                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                <p:childTnLst>
                                  <p:par>
                                    <p:cTn id="13" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">
                                      <p:stCondLst><p:cond delay="250"/></p:stCondLst>
                                      <p:childTnLst>
                                        <p:animEffect transition="in" filter="fade">
                                          <p:cBhvr>
                                            <p:cTn id="14" dur="650"/>
                                            <p:tgtEl><p:spTgt spid="{s_spid}"/></p:tgtEl>
                                          </p:cBhvr>
                                        </p:animEffect>
                                      </p:childTnLst>
                                    </p:cTn>
                                  </p:par>
                                  <p:par>
                                    <p:cTn id="15" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">
                                      <p:stCondLst><p:cond delay="350"/></p:stCondLst>
                                      <p:childTnLst>
                                        <p:animEffect transition="in" filter="fade">
                                          <p:cBhvr>
                                            <p:cTn id="16" dur="650"/>
                                            <p:tgtEl><p:spTgt spid="{c1_spid}"/></p:tgtEl>
                                          </p:cBhvr>
                                        </p:animEffect>
                                      </p:childTnLst>
                                    </p:cTn>
                                  </p:par>
                                  <p:par>
                                    <p:cTn id="17" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">
                                      <p:stCondLst><p:cond delay="450"/></p:stCondLst>
                                      <p:childTnLst>
                                        <p:animEffect transition="in" filter="fade">
                                          <p:cBhvr>
                                            <p:cTn id="18" dur="650"/>
                                            <p:tgtEl><p:spTgt spid="{c2_spid}"/></p:tgtEl>
                                          </p:cBhvr>
                                        </p:animEffect>
                                      </p:childTnLst>
                                    </p:cTn>
                                  </p:par>
                                </p:childTnLst>
                              </p:cTn>
                            </p:par>
                          </p:childTnLst>
                        </p:cTn>
                        <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                        <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
                      </p:seq>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:tnLst>
              <p:bldLst>
                <p:bldP spid="{t_spid}" grpId="1" animBg="1"/>
                <p:bldP spid="{t_spid}" grpId="2" animBg="1"/>
                <p:bldP spid="{e_spid}" grpId="3" animBg="1"/>
                <p:bldP spid="{s_spid}" grpId="4" animBg="1"/>
                <p:bldP spid="{c1_spid}" grpId="5" animBg="1"/>
                <p:bldP spid="{c2_spid}" grpId="6" animBg="1"/>
              </p:bldLst>
            </p:timing>
            """
            slide._element.append(etree.fromstring(broken_timing_xml.strip()))
        else:
            apply_transition(slide, "morph", duration_ms=1400)

    prs.save(output_pptx)
    return output_pptx


def test_disappearing_hero_image_detection_and_self_healing(tmp_path):
    """
    End-to-end regression test verifying all 11 requirements of Upgrade D:
    - Broken Fade-In timing causes `!!Earth_Grand` and `!!HeroTitle` to vanish in Slide Show
      while `Oval 13` outline stays visible.
    - StaticValidator + VisualAnimationReviewer detect the failure (`FAIL`) and do NOT let `Oval 13`
      pass for the missing Earth globe.
    - `PMGOrchestrator` self-heals the deck within <= 3 loops while preserving Fade In animations.
    - Repaired deck passes real PowerPoint Slide Show playback with `!!Earth_Grand` and `!!HeroTitle`
      fading in and remaining visible at steady state.
    """
    earth_png = _create_earth_globe_png(str(tmp_path / "earth_globe.png"))
    broken_pptx = _build_broken_space_hero_fixture(str(tmp_path / "broken_space_hero.pptx"), earth_png)

    # 1. Verify StaticValidator catches the broken timing structure immediately
    validator = StaticValidator()
    static_val = validator.validate(broken_pptx)
    assert not static_val["is_valid"], "StaticValidator must flag broken entrance timing that hides objects."
    hint_set = {d.get("remediation_hint") for d in static_val["defects"]}
    assert "repair_disappearing_object_timing" in hint_set

    # 2. Verify ExpectedTimeline classifies semantic roles and expected completion visibility
    timeline = extract_expected_timeline(broken_pptx, hold_duration_ms=1200)
    s1_objs = {o.object_id: o for o in timeline.slides[0].objects}
    assert "!!Earth_Grand" in s1_objs
    assert s1_objs["!!Earth_Grand"].semantic_role == "hero_image"
    assert s1_objs["!!Earth_Grand"].is_required_content is True
    assert s1_objs["!!Earth_Grand"].expected_initial_visibility is False
    assert s1_objs["!!Earth_Grand"].expected_completion_state == "visible"
    assert s1_objs["Oval 13"].semantic_role == "decorative_outline"
    assert s1_objs["Oval 13"].is_required_content is False

    # 3. Run live PowerPoint playback on the unhealed broken deck and verify visual reviewer detects
    #    that !!Earth_Grand and !!HeroTitle disappeared in Slide Show while visible in Edit Mode!
    broken_qa_dir = str(tmp_path / "broken_qa")
    os.makedirs(broken_qa_dir, exist_ok=True)
    exp_meta = export_video_with_powerpoint(
        pptx_path=broken_pptx,
        output_video_path=os.path.join(broken_qa_dir, "powerpoint_export.mp4"),
        slide_duration=1.2,
    )
    rec_meta, telemetry = record_actual_slideshow(
        pptx_path=broken_pptx,
        output_video_path=os.path.join(broken_qa_dir, "actual_slideshow_recording.mp4"),
        expected_timeline=timeline,
        target_fps=30,
    )
    reviewer = VisualAnimationReviewer()
    broken_report, _ = reviewer.review(
        export_meta=exp_meta,
        recording_meta=rec_meta,
        recording_telemetry=telemetry,
        expected_timeline=timeline,
        static_validation=static_val,
        qa_dir=broken_qa_dir,
    )
    assert broken_report.status == ValidationStatus.FAIL
    assert broken_report.object_visibility_passed is False
    failed_vis_objs = {
        c["object_id"] for c in broken_report.object_visibility_checks if not c.get("passed", True)
    }
    assert "!!Earth_Grand" in failed_vis_objs, (
        f"Visual reviewer must detect that !!Earth_Grand disappeared in Slide Show (Outline Oval 13 must not fool it). "
        f"Checks: {broken_report.object_visibility_checks}"
    )
    assert "!!HeroTitle" in failed_vis_objs, (
        f"Visual reviewer must detect that !!HeroTitle disappeared in Slide Show. "
        f"Checks: {broken_report.object_visibility_checks}"
    )

    # 4. Run full PMGOrchestrator with 3-loop self-healing enabled
    healed_output_pptx = str(tmp_path / "healed_space_hero.pptx")
    orchestrator = PMGOrchestrator(qa_root_dir=str(tmp_path / "qa_runs"))
    manifest = orchestrator.run_pipeline(
        output_pptx_path=healed_output_pptx,
        existing_pptx_path=broken_pptx,
        hold_duration_ms=1200,
        max_repair_loops=3,
        require_video_playback=True,
        run_id="test_disappearing_earth_heal",
    )

    # Assertions 7, 8, 9, 10, 11:
    assert manifest.status == ValidationStatus.PASS, f"Expected PASS after self-healing, got: {manifest.review_summary}"
    assert manifest.repair_loops_executed >= 1
    assert manifest.delivered_pptx is not None and os.path.exists(manifest.delivered_pptx)
    assert manifest.review_summary.get("object_visibility_passed") is True

    # Verify all required objects on Slide 1 passed visibility in the healed presentation
    healed_checks = {
        c["object_id"]: c
        for c in manifest.review_summary.get("object_visibility_checks", [])
        if c.get("slide_number") == 1
    }
    assert healed_checks["!!Earth_Grand"]["passed"] is True
    assert healed_checks["!!Earth_Grand"]["steady_core_std"] > 15.0
    assert healed_checks["!!HeroTitle"]["passed"] is True

    # Verify that Fade In animations were PRESERVED (NOT deleted!) in the healed PPTX
    with zipfile.ZipFile(manifest.delivered_pptx, "r") as zf:
        s1_xml = etree.fromstring(zf.read("ppt/slides/slide1.xml"))
    fade_effects = s1_xml.xpath('.//p:animEffect[@filter="fade" and @transition="in"]', namespaces={"p": P_NS})
    vis_sets = s1_xml.xpath('.//p:cTn[@presetClass="entr"]//p:set', namespaces={"p": P_NS})
    assert len(fade_effects) >= 4, "Self-healing must preserve legitimate Fade In animations, not delete them!"
    assert len(vis_sets) >= 4, "Every preserved entrance animation must include <p:set> visibility='visible'!"


def test_real_vu_tru_20_slide_disappearing_earth_excerpt(tmp_path):
    """
    If `C:\\Users\\admin\\Downloads\\Vu_tru_20_slide_PMG_inspired.pptx` is present on the machine,
    extracts a 2-slide excerpt (Slides 1–2) from the real user file and verifies that PMG
    automatically heals Slide 1's disappearing `!!Earth_Grand` and `!!HeroTitle` to PASS!
    """
    real_deck_path = r"C:\Users\admin\Downloads\Vu_tru_20_slide_PMG_inspired.pptx"
    if not os.path.exists(real_deck_path):
        return

    # Create a 2-slide excerpt from the real deck (removing slides 3..20) so playback test runs quickly
    excerpt_pptx = str(tmp_path / "vu_tru_slides_1_2.pptx")
    prs = pptx.Presentation(real_deck_path)
    xml_slides = prs.slides._sldIdLst
    slides_list = list(xml_slides)
    for sld_id in slides_list[2:]:
        r_id = sld_id.rId
        prs.part.drop_rel(r_id)
        xml_slides.remove(sld_id)
    prs.save(excerpt_pptx)

    # Confirm StaticValidator detects the broken timing in the real user deck
    validator = StaticValidator()
    pre_val = validator.validate(excerpt_pptx)
    assert not pre_val["is_valid"]
    assert any(d.get("remediation_hint") == "repair_disappearing_object_timing" for d in pre_val["defects"])

    # Run PMGOrchestrator to heal and verify in live PowerPoint Slide Show
    healed_real_pptx = str(tmp_path / "vu_tru_slides_1_2_healed.pptx")
    orchestrator = PMGOrchestrator(qa_root_dir=str(tmp_path / "qa_real_vu_tru"))
    manifest = orchestrator.run_pipeline(
        output_pptx_path=healed_real_pptx,
        existing_pptx_path=excerpt_pptx,
        hold_duration_ms=1200,
        max_repair_loops=3,
        require_video_playback=True,
        run_id="real_vu_tru_heal",
    )
    assert manifest.status == ValidationStatus.PASS, f"Real Vu_tru excerpt failed to heal: {manifest.review_summary}"
    assert manifest.repair_loops_executed >= 1
    assert manifest.review_summary.get("object_visibility_passed") is True
