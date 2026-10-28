"""
benchmarks.cases
================
Defines all 24 canonical animation benchmark cases across Groups A, B, C, and D
for the PMG Animation Benchmark Suite (v1.5.1 Upgrade B).
"""

import os
from dataclasses import dataclass
from typing import Callable, List, Dict, Any
import numpy as np
import cv2
import pptx
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from lxml import etree

from pmg.animations import SlideAnimationTimeline, P_NS
from scripts.motion_graph_kit import apply_transition


@dataclass
class BenchmarkCaseSpec:
    case_id: int
    slug: str
    group: str
    title: str
    description: str
    expected_effects: List[str]
    is_negative_or_repair_test: bool
    builder_fn: Callable[[str, str], str]


def _ensure_planet_png(assets_dir: str, name: str = "planet_hero.png", hue_shift: int = 0) -> str:
    os.makedirs(assets_dir, exist_ok=True)
    path = os.path.join(assets_dir, name)
    if os.path.exists(path):
        return path
    size = 420
    img = np.zeros((size, size, 3), dtype=np.uint8)
    img[:, :] = (16, 12, 8)
    c = (size // 2, size // 2)
    r = int(size * 0.43)
    base_bgr = (max(40, (215 + hue_shift) % 255), max(60, (125 + hue_shift * 2) % 255), max(30, (45 + hue_shift * 3) % 255))
    cv2.circle(img, c, r, base_bgr, -1, lineType=cv2.LINE_AA)
    cv2.ellipse(img, (c[0] - 25, c[1] - 20), (75, 42), 20, 0, 360, (90, 205, 85), -1, lineType=cv2.LINE_AA)
    cv2.ellipse(img, (c[0] + 35, c[1] + 28), (65, 36), -15, 0, 360, (240, 245, 255), -1, lineType=cv2.LINE_AA)
    cv2.circle(img, (c[0] - 45, c[1] - 45), 24, (255, 250, 240), -1, lineType=cv2.LINE_AA)
    cv2.imwrite(path, img)
    return path


def _new_base_slide(prs: pptx.Presentation, title_text: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.name = "!!Canvas_BG"
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(248, 250, 252)
    bg.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(10.5), Inches(1.0))
    tb.name = "!!Slide_Title"
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(34)
    p.font.bold = True
    p.font.color.rgb = RGBColor(15, 23, 42)
    return slide, tb


def _add_card(slide, name: str, label: str, x: float, y: float, w: float = 3.2, h: float = 2.0, rgb=(2, 132, 199)):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.name = name
    shp.fill.solid()
    shp.fill.fore_color.rgb = RGBColor(*rgb)
    shp.line.color.rgb = RGBColor(255, 255, 255)
    shp.text_frame.text = label
    shp.text_frame.paragraphs[0].font.size = Pt(20)
    shp.text_frame.paragraphs[0].font.bold = True
    shp.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    return shp


# --- GROUP A (Cases 1..8): Single-Object Animations ---

def _build_case_01(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 01: Native Fade In Entrance")
    card = _add_card(s1, "!!FadeIn_Target", "FADE IN OBJECT", 4.8, 2.6, 3.8, 2.2)
    tl = SlideAnimationTimeline(1)
    tl.add_effect(card, "fade", trigger="with_previous", duration_ms=700, delay_ms=100)
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_02(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 02: Native Fade Out Exit")
    card = _add_card(s1, "!!FadeOut_Target", "FADE OUT OBJECT", 4.8, 2.6, 3.8, 2.2, rgb=(225, 29, 72))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(card, "fade_out", trigger="with_previous", duration_ms=650, delay_ms=150)
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_03(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 03: Fly In & Fly Out Choreography")
    c_in = _add_card(s1, "!!FlyIn_Card", "FLY IN UP", 2.2, 2.8, 3.4, 2.0, rgb=(5, 150, 105))
    c_out = _add_card(s1, "!!FlyOut_Card", "FLY OUT RIGHT", 7.2, 2.8, 3.4, 2.0, rgb=(217, 119, 6))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(c_in, "fly_in", trigger="with_previous", duration_ms=650, direction="up")
    tl.add_effect(c_out, "fly_out", trigger="after_previous", duration_ms=650, direction="right")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_04(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 04: Directional Wipe Reveal")
    bar = _add_card(s1, "!!Wipe_Bar", "DIRECTIONAL WIPE PROGRESS", 2.0, 3.0, 8.8, 1.6, rgb=(79, 70, 229))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(bar, "wipe", trigger="with_previous", duration_ms=750, direction="left")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_05(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 05: Native 360-Degree Spin Rotation")
    prism = _add_card(s1, "!!Spin_Prism", "ROTATING CORE", 5.2, 2.5, 2.8, 2.8, rgb=(14, 165, 233))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(prism, "spin", trigger="with_previous", duration_ms=800, rotation_deg=360.0)
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_06(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 06: Native Scale / Grow-Shrink")
    node = _add_card(s1, "!!Scale_Node", "SCALE 135%", 4.9, 2.6, 3.5, 2.2, rgb=(16, 185, 129))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(node, "scale", trigger="with_previous", duration_ms=750, scale_factor=(1.35, 1.35))
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_07(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 07: Linear Native Motion Path")
    shuttle = _add_card(s1, "!!Linear_Shuttle", "LINEAR VECTOR", 1.5, 3.0, 2.6, 1.6, rgb=(2, 132, 199))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(shuttle, "motion_path", trigger="with_previous", duration_ms=850, motion_path="M 0 0 L 0.42 0 E")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_08(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 08: Curved Cubic Bezier Motion Path")
    probe = _add_card(s1, "!!Bezier_Probe", "ORBITAL ARC", 1.5, 3.4, 2.5, 1.5, rgb=(139, 92, 246))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(
        probe,
        "motion_path",
        trigger="with_previous",
        duration_ms=900,
        motion_path="M 0 0 C 0.14 -0.22 0.28 0.22 0.42 0 E",
    )
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


# --- GROUP B (Cases 9..14): Multi-Object Choreography ---

def _build_case_09(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 09: Simultaneous Multi-Object Entrance")
    c1 = _add_card(s1, "!!Sim_Card_A", "PARALLEL ALPHA", 1.2, 2.6, 3.2, 2.2)
    c2 = _add_card(s1, "!!Sim_Card_B", "PARALLEL BETA", 5.0, 2.6, 3.2, 2.2, rgb=(16, 185, 129))
    c3 = _add_card(s1, "!!Sim_Card_C", "PARALLEL GAMMA", 8.8, 2.6, 3.2, 2.2, rgb=(79, 70, 229))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(c1, "fade", trigger="with_previous", duration_ms=650, delay_ms=0)
    tl.add_effect(c2, "fade", trigger="with_previous", duration_ms=650, delay_ms=0)
    tl.add_effect(c3, "fade", trigger="with_previous", duration_ms=650, delay_ms=0)
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_10(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 10: Sequential Staggered Cascade")
    c1 = _add_card(s1, "!!Seq_Card_1", "STEP 01", 1.2, 2.6, 3.2, 2.2)
    c2 = _add_card(s1, "!!Seq_Card_2", "STEP 02", 5.0, 2.6, 3.2, 2.2, rgb=(16, 185, 129))
    c3 = _add_card(s1, "!!Seq_Card_3", "STEP 03", 8.8, 2.6, 3.2, 2.2, rgb=(217, 119, 6))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(c1, "fly_in", trigger="with_previous", duration_ms=450, delay_ms=0, direction="up")
    tl.add_effect(c2, "fly_in", trigger="with_previous", duration_ms=450, delay_ms=220, direction="up")
    tl.add_effect(c3, "fly_in", trigger="with_previous", duration_ms=450, delay_ms=440, direction="up")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_11(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, tb = _new_base_slide(prs, "Case 11: With-Previous Synchronized Timing")
    c1 = _add_card(s1, "!!WithPrev_Card", "SYNCHRONIZED PANEL", 4.5, 2.8, 4.2, 2.2)
    tl = SlideAnimationTimeline(1)
    tl.add_effect(tb, "fade", trigger="with_previous", duration_ms=600, delay_ms=0)
    tl.add_effect(c1, "zoom_in", trigger="with_previous", duration_ms=650, delay_ms=100)
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_12(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 12: After-Previous Chain Choreography")
    c1 = _add_card(s1, "!!Chain_A", "FIRST EVENT", 2.0, 2.8, 3.6, 2.0)
    c2 = _add_card(s1, "!!Chain_B", "SECOND EVENT", 7.2, 2.8, 3.6, 2.0, rgb=(5, 150, 105))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(c1, "fade", trigger="with_previous", duration_ms=450)
    tl.add_effect(c2, "wipe", trigger="after_previous", duration_ms=500, direction="left")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_13(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 13: Interactive On-Click Trigger")
    c1 = _add_card(s1, "!!Auto_Intro", "AUTO INTRO", 2.0, 2.8, 3.6, 2.0)
    c2 = _add_card(s1, "!!Click_Reveal", "ON CLICK REVEAL", 7.2, 2.8, 3.6, 2.0, rgb=(225, 29, 72))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(c1, "fade", trigger="with_previous", duration_ms=450)
    tl.add_effect(c2, "zoom_in", trigger="on_click", duration_ms=550)
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_14(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 14: Multiple Effects on Single Object")
    hero = _add_card(s1, "!!MultiEffect_Hero", "ENTER + MOVE + PULSE", 2.0, 2.8, 3.4, 2.0, rgb=(79, 70, 229))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(hero, "fade", trigger="with_previous", duration_ms=400)
    tl.add_effect(hero, "motion_path", trigger="after_previous", duration_ms=550, motion_path="M 0 0 L 0.32 0 E")
    tl.add_effect(hero, "pulse", trigger="after_previous", duration_ms=400)
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


# --- GROUP C (Cases 15..20): Advanced PMG Motion ---

def _build_case_15(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 15: Morph Position Interpolation (Slide 1)")
    _add_card(s1, "!!Morph_Actor", "MORPH ACTOR", 1.5, 2.8, 3.2, 2.0)

    s2, _ = _new_base_slide(prs, "Case 15: Morph Position Interpolation (Slide 2)")
    _add_card(s2, "!!Morph_Actor", "MORPH ACTOR", 8.2, 2.8, 3.2, 2.0)
    apply_transition(s2, "morph", duration_ms=1300)
    prs.save(out_pptx)
    return out_pptx


def _build_case_16(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_base_slide(prs, "Case 16: Morph Scale & Rotation (Slide 1)")
    c1 = _add_card(s1, "!!Morph_Transform", "TRANSFORM", 2.2, 2.8, 2.4, 1.6, rgb=(16, 185, 129))
    c1.rotation = 0.0

    s2, _ = _new_base_slide(prs, "Case 16: Morph Scale & Rotation (Slide 2)")
    c2 = _add_card(s2, "!!Morph_Transform", "TRANSFORM", 6.8, 2.0, 4.4, 3.0, rgb=(16, 185, 129))
    c2.rotation = 18.0
    apply_transition(s2, "morph", duration_ms=1300)
    prs.save(out_pptx)
    return out_pptx


def _build_case_17(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    coords = [(1.2, 4.2), (5.4, 2.4), (9.2, 4.0)]
    for idx, (x, y) in enumerate(coords, 1):
        s, _ = _new_base_slide(prs, f"Case 17: Continuous Traveling Entity (Stop {idx})")
        orb = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(1.4), Inches(1.4))
        orb.name = "!!Traveling_Satellite"
        orb.fill.solid()
        orb.fill.fore_color.rgb = RGBColor(2, 132, 199)
        orb.line.color.rgb = RGBColor(255, 255, 255)
        if idx > 1:
            apply_transition(s, "morph", duration_ms=1200)
    prs.save(out_pptx)
    return out_pptx


def _build_case_18(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    planet_png = _ensure_planet_png(assets_dir, "zoom_planet.png", hue_shift=20)

    s1, _ = _new_base_slide(prs, "Case 18: Wide Orbital Establishing Shot")
    p1 = s1.shapes.add_picture(planet_png, Inches(5.4), Inches(2.6), width=Inches(2.4), height=Inches(2.4))
    p1.name = "!!Camera_Subject"

    s2, _ = _new_base_slide(prs, "Case 18: Macro Push-In Camera Zoom (220%)")
    p2 = s2.shapes.add_picture(planet_png, Inches(4.0), Inches(1.4), width=Inches(5.2), height=Inches(5.2))
    p2.name = "!!Camera_Subject"
    apply_transition(s2, "morph", duration_ms=1300)
    prs.save(out_pptx)
    return out_pptx


def _build_case_19(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    for idx in (1, 2):
        s, _ = _new_base_slide(prs, f"Case 19: 3-Plane Parallax Depth (Frame {idx})")
        shift_bg = 0.0 if idx == 1 else -0.7
        shift_mid = 0.0 if idx == 1 else -1.8
        shift_fg = 0.0 if idx == 1 else -3.0

        _add_card(s, "!!Parallax_BG_Layer", "BG 0.7x", 3.5 + shift_bg, 2.0, 2.6, 1.4, rgb=(148, 163, 184))
        _add_card(s, "!!Parallax_Mid_Layer", "MID 1.0x", 5.5 + shift_mid, 3.2, 3.0, 1.6, rgb=(2, 132, 199))
        _add_card(s, "!!Parallax_FG_Layer", "FG 1.4x", 7.8 + shift_fg, 4.6, 3.4, 1.8, rgb=(15, 23, 42))
        if idx > 1:
            apply_transition(s, "morph", duration_ms=1300)
    prs.save(out_pptx)
    return out_pptx


def _build_case_20(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, tb = _new_base_slide(prs, "Case 20: Complex Overlapping Multi-Track Timeline")
    c1 = _add_card(s1, "!!Track_1", "TRACK 1 FADE", 1.2, 2.5, 3.2, 2.0)
    c2 = _add_card(s1, "!!Track_2", "TRACK 2 WIPE", 5.0, 2.5, 3.2, 2.0, rgb=(5, 150, 105))
    c3 = _add_card(s1, "!!Track_3", "TRACK 3 PATH", 8.8, 2.5, 3.2, 2.0, rgb=(139, 92, 246))
    tl = SlideAnimationTimeline(1)
    tl.add_effect(tb, "fade", trigger="with_previous", duration_ms=500, delay_ms=0)
    tl.add_effect(c1, "fly_in", trigger="with_previous", duration_ms=600, delay_ms=120, direction="up")
    tl.add_effect(c2, "wipe", trigger="with_previous", duration_ms=650, delay_ms=260, direction="left")
    tl.add_effect(c3, "motion_path", trigger="with_previous", duration_ms=700, delay_ms=340, motion_path="M 0 0 L -0.12 0.12 E")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


# --- GROUP D (Cases 21..24): Real-World Failure & Visibility Cases ---

def _build_case_21(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    planet_png = _ensure_planet_png(assets_dir, "hero_earth_21.png", hue_shift=0)
    s1, tb = _new_base_slide(prs, "Case 21: Hero Image Visibility with Sibling Outline")
    pic = s1.shapes.add_picture(planet_png, Inches(7.8), Inches(1.5), width=Inches(4.4), height=Inches(4.4))
    pic.name = "!!Earth_Grand"
    ring = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.68), Inches(1.38), Inches(4.64), Inches(4.64))
    ring.name = "Oval 13"
    ring.fill.background()
    ring.line.color.rgb = RGBColor(2, 132, 199)
    ring.line.width = Pt(2.25)
    c1 = _add_card(s1, "!!Card_Info", "HERO TELEMETRY", 0.8, 2.8, 4.8, 2.2)

    tl = SlideAnimationTimeline(1)
    tl.add_effect(tb, "fade", trigger="with_previous", duration_ms=550, delay_ms=0)
    tl.add_effect(pic, "fade", trigger="with_previous", duration_ms=650, delay_ms=120)
    tl.add_effect(c1, "fade", trigger="with_previous", duration_ms=600, delay_ms=220)
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_22(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, tb = _new_base_slide(prs, "Case 22: Title & Subtitle Text Visibility")
    tb.name = "!!HeroTitle"
    sub = s1.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(9.5), Inches(0.9))
    sub.name = "!!HeroSubtitle"
    p = sub.text_frame.paragraphs[0]
    p.text = "High-contrast typography remains visible after entrance completion"
    p.font.size = Pt(22)
    p.font.color.rgb = RGBColor(51, 65, 85)
    c1 = _add_card(s1, "!!Card_Summary", "KEY METRIC SUMMARY", 0.8, 3.2, 5.5, 2.4)

    tl = SlideAnimationTimeline(1)
    tl.add_effect(tb, "fade", trigger="with_previous", duration_ms=500, delay_ms=0)
    tl.add_effect(sub, "fade", trigger="with_previous", duration_ms=500, delay_ms=150)
    tl.add_effect(c1, "fly_in", trigger="with_previous", duration_ms=600, delay_ms=250, direction="up")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_23(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    img1 = _ensure_planet_png(assets_dir, "planet_a_23.png", hue_shift=10)
    img2 = _ensure_planet_png(assets_dir, "planet_b_23.png", hue_shift=70)
    img3 = _ensure_planet_png(assets_dir, "planet_c_23.png", hue_shift=140)
    s1, tb = _new_base_slide(prs, "Case 23: Multiple Animated Images with Independent Timing")
    p1 = s1.shapes.add_picture(img1, Inches(1.0), Inches(2.4), width=Inches(3.2), height=Inches(3.2))
    p1.name = "!!Planet_Alpha"
    p2 = s1.shapes.add_picture(img2, Inches(5.0), Inches(2.4), width=Inches(3.2), height=Inches(3.2))
    p2.name = "!!Planet_Beta"
    p3 = s1.shapes.add_picture(img3, Inches(9.0), Inches(2.4), width=Inches(3.2), height=Inches(3.2))
    p3.name = "!!Planet_Gamma"

    tl = SlideAnimationTimeline(1)
    tl.add_effect(tb, "fade", trigger="with_previous", duration_ms=450, delay_ms=0)
    tl.add_effect(p1, "fade", trigger="with_previous", duration_ms=550, delay_ms=100)
    tl.add_effect(p2, "zoom_in", trigger="with_previous", duration_ms=550, delay_ms=280)
    tl.add_effect(p3, "fly_in", trigger="with_previous", duration_ms=550, delay_ms=460, direction="up")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_24_broken(out_pptx: str, assets_dir: str) -> str:
    """
    Case 24 (Negative + Automatic Repair Benchmark):
    Deliberately injects the broken <p:timing> structure where `<p:set>` visibility is missing,
    `grpId` is missing, and duplicate conflicting entrance effects target `!!HeroTitle` while
    `Oval 13` outline surrounds `!!Earth_Grand`.
    Verifies that PMG detects the broken deck as FAIL prior to healing and automatically repairs
    it to PASS while preserving the Fade In animations.
    """
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    planet_png = _ensure_planet_png(assets_dir, "hero_earth_24.png", hue_shift=0)
    s1 = prs.slides.add_slide(prs.slide_layouts[6])

    bg = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.name = "!!Starfield_BG"
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(10, 15, 28)
    bg.line.fill.background()

    tb = s1.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(6.2), Inches(1.4))
    tb.name = "!!HeroTitle"
    p = tb.text_frame.paragraphs[0]
    p.text = "CASE 24: AUTO-REPAIR"
    p.font.size = Pt(42)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)

    pic = s1.shapes.add_picture(planet_png, Inches(7.8), Inches(1.3), width=Inches(4.6), height=Inches(4.6))
    pic.name = "!!Earth_Grand"

    ring = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.68), Inches(1.18), Inches(4.84), Inches(4.84))
    ring.name = "Oval 13"
    ring.fill.background()
    ring.line.color.rgb = RGBColor(56, 189, 248)
    ring.line.width = Pt(2.25)

    t_spid = tb.shape_id
    e_spid = pic.shape_id
    broken_xml = f"""
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
                                  <p:cBhvr><p:cTn id="5" dur="650"/><p:tgtEl><p:spTgt spid="{t_spid}"/></p:tgtEl></p:cBhvr>
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
                              <p:stCondLst><p:cond delay="100"/></p:stCondLst>
                              <p:childTnLst>
                                <p:animEffect transition="in" filter="fade">
                                  <p:cBhvr><p:cTn id="8" dur="650"/><p:tgtEl><p:spTgt spid="{t_spid}"/></p:tgtEl></p:cBhvr>
                                </p:animEffect>
                              </p:childTnLst>
                            </p:cTn>
                          </p:par>
                          <p:par>
                            <p:cTn id="9" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">
                              <p:stCondLst><p:cond delay="200"/></p:stCondLst>
                              <p:childTnLst>
                                <p:animEffect transition="in" filter="fade">
                                  <p:cBhvr><p:cTn id="10" dur="700"/><p:tgtEl><p:spTgt spid="{e_spid}"/></p:tgtEl></p:cBhvr>
                                </p:animEffect>
                              </p:childTnLst>
                            </p:cTn>
                          </p:par>
                        </p:childTnLst>
                      </p:cTn>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
      <p:bldLst>
        <p:bldP spid="{t_spid}" grpId="1" animBg="1"/>
        <p:bldP spid="{t_spid}" grpId="2" animBg="1"/>
        <p:bldP spid="{e_spid}" grpId="3" animBg="1"/>
      </p:bldLst>
    </p:timing>
    """
    s1._element.append(etree.fromstring(broken_xml.strip()))
    prs.save(out_pptx)
    return out_pptx


BENCHMARK_CASES: List[BenchmarkCaseSpec] = [
    BenchmarkCaseSpec(1, "case_01_fade_in", "Group A", "Fade In", "Single-object native Fade In entrance.", ["fade"], False, _build_case_01),
    BenchmarkCaseSpec(2, "case_02_fade_out", "Group A", "Fade Out", "Single-object native Fade Out exit.", ["fade_out"], False, _build_case_02),
    BenchmarkCaseSpec(3, "case_03_fly_in_out", "Group A", "Fly In / Fly Out", "Directional Fly In and Fly Out choreography.", ["fly_in", "fly_out"], False, _build_case_03),
    BenchmarkCaseSpec(4, "case_04_wipe", "Group A", "Wipe", "Directional Wipe entrance.", ["wipe"], False, _build_case_04),
    BenchmarkCaseSpec(5, "case_05_rotation", "Group A", "Rotation", "360-degree native Spin emphasis.", ["spin"], False, _build_case_05),
    BenchmarkCaseSpec(6, "case_06_scale_grow_shrink", "Group A", "Scale / Grow-Shrink", "135% native Grow/Shrink emphasis.", ["scale"], False, _build_case_06),
    BenchmarkCaseSpec(7, "case_07_linear_motion_path", "Group A", "Linear Motion Path", "Horizontal linear native motion path.", ["motion_path"], False, _build_case_07),
    BenchmarkCaseSpec(8, "case_08_curved_bezier_motion_path", "Group A", "Curved Bezier Motion Path", "Cubic Bezier orbital arc motion path.", ["motion_path"], False, _build_case_08),
    BenchmarkCaseSpec(9, "case_09_simultaneous", "Group B", "Simultaneous Animations", "Three cards animating in parallel.", ["fade"], False, _build_case_09),
    BenchmarkCaseSpec(10, "case_10_sequential", "Group B", "Sequential Animations", "Three cards cascading sequentially with stagger delays.", ["fly_in"], False, _build_case_10),
    BenchmarkCaseSpec(11, "case_11_with_previous", "Group B", "With Previous Timing", "Synchronized title and zoom-in panel.", ["fade", "zoom_in"], False, _build_case_11),
    BenchmarkCaseSpec(12, "case_12_after_previous", "Group B", "After Previous Timing", "Chained Fade followed by directional Wipe.", ["fade", "wipe"], False, _build_case_12),
    BenchmarkCaseSpec(13, "case_13_on_click", "Group B", "On Click Trigger", "Auto intro card followed by interactive On-Click reveal.", ["fade", "zoom_in"], False, _build_case_13),
    BenchmarkCaseSpec(14, "case_14_multiple_effects_single_object", "Group B", "Multiple Effects on One Object", "Fade In + Linear Motion Path + Pulse on a single actor.", ["fade", "motion_path", "pulse"], False, _build_case_14),
    BenchmarkCaseSpec(15, "case_15_morph_position", "Group C", "Morph Position Transition", "Cross-slide 60fps Morph position translation.", ["morph"], False, _build_case_15),
    BenchmarkCaseSpec(16, "case_16_morph_scale_rotation", "Group C", "Morph Scale & Rotation", "Cross-slide Morph scale expansion and 18-deg rotation.", ["morph"], False, _build_case_16),
    BenchmarkCaseSpec(17, "case_17_traveling_object_across_slides", "Group C", "Traveling Object Across Slides", "3-slide continuous traveling satellite trajectory.", ["morph"], False, _build_case_17),
    BenchmarkCaseSpec(18, "case_18_camera_zoom_simulation", "Group C", "Camera Zoom Simulation", "Wide shot to 220% macro push-in camera zoom.", ["morph"], False, _build_case_18),
    BenchmarkCaseSpec(19, "case_19_multi_plane_parallax", "Group C", "Multi-Plane Parallax", "3-plane speed ratio parallax (BG 0.7x, Mid 1.0x, FG 1.4x).", ["morph"], False, _build_case_19),
    BenchmarkCaseSpec(20, "case_20_complex_overlapping_timeline", "Group C", "Complex Overlapping Timeline", "4-track overlapping Fade, Fly In, Wipe, and Motion Path.", ["fade", "fly_in", "wipe", "motion_path"], False, _build_case_20),
    BenchmarkCaseSpec(21, "case_21_hero_image_visibility", "Group D", "Hero Image Visibility During Animation", "Hero Earth image with sibling Oval 13 decorative outline.", ["fade"], False, _build_case_21),
    BenchmarkCaseSpec(22, "case_22_title_text_visibility", "Group D", "Title & Text Visibility During Animation", "Animated HeroTitle, HeroSubtitle, and summary card visibility.", ["fade", "fly_in"], False, _build_case_22),
    BenchmarkCaseSpec(23, "case_23_multiple_animated_images", "Group D", "Multiple Animated Images with Independent Timing", "Three planetary hero images entering with independent effects.", ["fade", "zoom_in", "fly_in"], False, _build_case_23),
    BenchmarkCaseSpec(24, "case_24_failure_detection_and_repair", "Group D", "Animation Failure Detection & Automatic Repair", "Broken disappearing Earth/Title XML detected as FAIL and healed to PASS.", ["fade"], True, _build_case_24_broken),
]


def get_benchmark_case(case_id: int) -> BenchmarkCaseSpec:
    for c in BENCHMARK_CASES:
        if c.case_id == case_id:
            return c
    raise KeyError(f"Unknown benchmark case ID: {case_id} (valid: 1..24)")

