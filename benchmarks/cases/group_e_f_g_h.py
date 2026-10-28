"""benchmarks.cases.group_e_f_g_h
===============================
Cases 25..44 for PMG Skills v1.8.0 Animation Benchmark Suite.
- Group E (Cases 25..29): Advanced Motion Paths & Dynamics
- Group F (Cases 30..34): Massive Scene Hierarchies & Layered Parallax
- Group G (Cases 35..39): Decomposed Asset Choreography & Reassembly
- Group H (Cases 40..44): End-to-End Cinematic Storyboard & Motion Director
"""
import os
import math
import pptx
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

from pmg.animations import SlideAnimationTimeline
from scripts.motion_graph_kit import apply_transition, link_morph
from pmg.compositing import (
    build_bezier_motion_path,
    build_orbital_motion_path,
    build_waypoint_motion_path,
    generate_massive_constellation,
)
from pmg.storyboard import StoryboardEngine
from pmg.motion_director import MotionDirector
from scripts.motion_graph_kit import Themes


def _new_slide(prs: pptx.Presentation, title: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.name = "!!Canvas_BG"
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(248, 250, 252)
    bg.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(10.5), Inches(1.0))
    tb.name = "!!Slide_Title"
    p = tb.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(15, 23, 42)
    return slide, tb


def _card(slide, name: str, text: str, x: float, y: float, w: float, h: float, rgb=(14, 165, 233)):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.name = name
    shp.fill.solid()
    shp.fill.fore_color.rgb = RGBColor(*rgb)
    shp.line.color.rgb = RGBColor(255, 255, 255)
    shp.text_frame.text = text
    shp.text_frame.paragraphs[0].font.size = Pt(16)
    shp.text_frame.paragraphs[0].font.bold = True
    shp.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    return shp


# --- GROUP E (Cases 25..29): Advanced Motion Paths & Dynamics ---

def _build_case_25(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 25: Orbital Bezier Loop Path")
    orbiter = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.0), Inches(3.2), Inches(1.2), Inches(1.2))
    orbiter.name = "!!Orbiter_Node"
    orbiter.fill.solid()
    orbiter.fill.fore_color.rgb = RGBColor(79, 70, 229)
    orbiter.line.fill.background()

    path_str = build_orbital_motion_path(cx=0.5, cy=0.5, rx=0.25, ry=0.15)
    tl = SlideAnimationTimeline(1)
    tl.add_effect(orbiter, "motion_path", motion_path=path_str, duration_ms=1200, trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_26(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 26: Staggered Radial Blossom")
    tl = SlideAnimationTimeline(1)
    for i in range(8):
        ang = i * (360.0 / 8.0)
        cx = 6.666 + 2.2 * math.cos(math.radians(ang)) - 0.4
        cy = 4.0 + 1.8 * math.sin(math.radians(ang)) - 0.4
        petal = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(cy), Inches(0.8), Inches(0.8))
        petal.name = f"Petal_{i}"
        petal.fill.solid()
        petal.fill.fore_color.rgb = RGBColor(14, 165, 233)
        petal.line.fill.background()
        tl.add_effect(petal, "zoom", duration_ms=400, delay_ms=i * 80, trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_27(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 27: Multi-Segment Waypoint Motion Path")
    carrier = s1.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(2.0), Inches(3.0), Inches(1.2), Inches(1.2))
    carrier.name = "!!Waypoint_Carrier"
    carrier.fill.solid()
    carrier.fill.fore_color.rgb = RGBColor(16, 185, 129)
    carrier.line.fill.background()

    way_path = build_waypoint_motion_path([(0.0, 0.0), (0.2, -0.15), (0.4, 0.15), (0.6, 0.0)])
    tl = SlideAnimationTimeline(1)
    tl.add_effect(carrier, "motion_path", motion_path=way_path, duration_ms=1400, trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_28(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 28: Continuous Micro Pulse and Spin")
    core = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.8), Inches(2.8), Inches(1.6), Inches(1.6))
    core.name = "!!PulseSpin_Core"
    core.fill.solid()
    core.fill.fore_color.rgb = RGBColor(236, 72, 153)
    core.line.color.rgb = RGBColor(255, 255, 255)
    core.line.width = Pt(2)

    tl = SlideAnimationTimeline(1)
    tl.add_effect(core, "pulse", duration_ms=800, scale_factor=(1.25, 1.25), trigger="with_previous")
    tl.add_effect(core, "rotate", duration_ms=1200, rotation_deg=360.0, trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_29(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 29: Auto-Reversing Oscillating Trajectory")
    osc = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.0), Inches(3.2), Inches(2.0), Inches(1.2))
    osc.name = "!!Oscillating_Actor"
    osc.fill.solid()
    osc.fill.fore_color.rgb = RGBColor(245, 158, 11)
    osc.line.fill.background()

    tl = SlideAnimationTimeline(1)
    tl.add_effect(osc, "motion_path", motion_path="M 0 0 L 0.35 0 E", duration_ms=900, auto_reverse=True, repeat_count=2, trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


# --- GROUP F (Cases 30..34): Massive Scene Hierarchies & Layered Parallax ---

def _build_case_30(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 30: Massive Hierarchy (50 Coordinated Nodes)")
    nodes = generate_massive_constellation(num_nodes=50, seed=30)
    for n in nodes:
        dot = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(n["x"]), Inches(n["y"]), Inches(n["size"]), Inches(n["size"]))
        dot.name = f"Node50_{n['id']}"
        dot.fill.solid()
        dot.fill.fore_color.rgb = RGBColor(14, 165, 233)
        dot.line.fill.background()
    prs.save(out_pptx)
    return out_pptx


def _build_case_31(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 31: Massive Hierarchy (100 Coordinated Nodes)")
    nodes = generate_massive_constellation(num_nodes=100, seed=31)
    for n in nodes:
        dot = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(n["x"]), Inches(n["y"]), Inches(n["size"]), Inches(n["size"]))
        dot.name = f"Node100_{n['id']}"
        dot.fill.solid()
        dot.fill.fore_color.rgb = RGBColor(99, 102, 241)
        dot.line.fill.background()
    prs.save(out_pptx)
    return out_pptx


def _build_case_32(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 32: Massive Hierarchy (200 Coordinated Nodes)")
    nodes = generate_massive_constellation(num_nodes=200, seed=32)
    for n in nodes:
        dot = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(n["x"]), Inches(n["y"]), Inches(n["size"]), Inches(n["size"]))
        dot.name = f"Node200_{n['id']}"
        dot.fill.solid()
        dot.fill.fore_color.rgb = RGBColor(16, 185, 129)
        dot.line.fill.background()
    prs.save(out_pptx)
    return out_pptx


def _build_case_33(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    # Slide 1: Initial camera position
    s1, _ = _new_slide(prs, "Case 33: Six-Layer Deep Parallax (Frame A)")
    apply_transition(s1, "none")
    l1 = _card(s1, "!!Parallax_DeepBG", "L1: DEEP BG (0.2x)", 1.0, 2.2, 3.0, 1.2, rgb=(148, 163, 184))
    l2 = _card(s1, "!!Parallax_BG", "L2: BG (0.5x)", 1.0, 3.6, 3.0, 1.2, rgb=(100, 116, 139))
    l3 = _card(s1, "!!Parallax_Mid", "L3: MID (1.0x)", 1.0, 5.0, 3.0, 1.2, rgb=(14, 165, 233))

    # Slide 2: Parallax displacement
    s2, _ = _new_slide(prs, "Case 33: Six-Layer Deep Parallax (Frame B)")
    apply_transition(s2, "morph")
    l1_b = _card(s2, "!!Parallax_DeepBG", "L1: DEEP BG (0.2x)", 1.6, 2.2, 3.0, 1.2, rgb=(148, 163, 184))
    l2_b = _card(s2, "!!Parallax_BG", "L2: BG (0.5x)", 2.5, 3.6, 3.0, 1.2, rgb=(100, 116, 139))
    l3_b = _card(s2, "!!Parallax_Mid", "L3: MID (1.0x)", 4.0, 5.0, 3.0, 1.2, rgb=(14, 165, 233))
    link_morph(l1, l1_b)
    link_morph(l2, l2_b)
    link_morph(l3, l3_b)
    prs.save(out_pptx)
    return out_pptx


def _build_case_34(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 34: Counter-Tracking HUD Overlay (Frame A)")
    apply_transition(s1, "none")
    bg_s1 = _card(s1, "!!Moving_Subject", "TERRAIN SUBJECT", 2.0, 3.0, 3.5, 2.0, rgb=(59, 130, 246))
    hud_s1 = _card(s1, "!!Fixed_HUD", "LOCKED HUD RETICLE", 9.0, 1.5, 3.0, 1.2, rgb=(225, 29, 72))

    s2, _ = _new_slide(prs, "Case 34: Counter-Tracking HUD Overlay (Frame B)")
    apply_transition(s2, "morph")
    bg_s2 = _card(s2, "!!Moving_Subject", "TERRAIN SUBJECT", 6.5, 3.0, 3.5, 2.0, rgb=(59, 130, 246))
    hud_s2 = _card(s2, "!!Fixed_HUD", "LOCKED HUD RETICLE", 9.0, 1.5, 3.0, 1.2, rgb=(225, 29, 72))
    link_morph(bg_s1, bg_s2)
    link_morph(hud_s1, hud_s2)
    prs.save(out_pptx)
    return out_pptx


# --- GROUP G (Cases 35..39): Decomposed Asset Choreography & Reassembly ---

def _build_case_35(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 35: Decomposed Particle Scatter & Gather")
    tl = SlideAnimationTimeline(1)
    for p_i in range(8):
        ang = p_i * (360.0 / 8.0)
        sx = 6.666 + 1.2 * math.cos(math.radians(ang)) - 0.25
        sy = 4.0 + 1.2 * math.sin(math.radians(ang)) - 0.25
        particle = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(sx), Inches(sy), Inches(0.5), Inches(0.5))
        particle.name = f"ScatterParticle_{p_i}"
        particle.fill.solid()
        particle.fill.fore_color.rgb = RGBColor(217, 119, 6)
        particle.line.fill.background()
        tl.add_effect(particle, "fade", duration_ms=450, delay_ms=p_i * 50, trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_36(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 36: 10-Strip Vertical Shutter Reveal")
    tl = SlideAnimationTimeline(1)
    w = 11.333 / 10.0
    for s_i in range(10):
        sx = 1.0 + s_i * w
        strip = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(sx), Inches(2.8), Inches(w - 0.05), Inches(3.5))
        strip.name = f"Strip10_{s_i}"
        strip.fill.solid()
        strip.fill.fore_color.rgb = RGBColor(14, 165, 233)
        strip.line.fill.background()
        tl.add_effect(strip, "wipe", duration_ms=400, delay_ms=s_i * 60, direction="up", trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_37(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 37: Radial Sector Iris Rotation")
    tl = SlideAnimationTimeline(1)
    for b_i in range(6):
        ang = b_i * (360.0 / 6.0)
        bx = 6.666 + 1.5 * math.cos(math.radians(ang)) - 0.6
        by = 4.0 + 1.5 * math.sin(math.radians(ang)) - 0.6
        blade = s1.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(bx), Inches(by), Inches(1.2), Inches(1.2))
        blade.name = f"IrisBlade_{b_i}"
        blade.rotation = int(ang)
        blade.fill.solid()
        blade.fill.fore_color.rgb = RGBColor(168, 43, 42)
        blade.line.fill.background()
        tl.add_effect(blade, "rotate", rotation_deg=60.0, duration_ms=600, delay_ms=b_i * 70, trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_38(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    # Slide 1: Vector polygon in position A
    s1, _ = _new_slide(prs, "Case 38: Vector Polygon Morph Reassembly (Frame A)")
    apply_transition(s1, "none")
    poly_a = s1.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(2.0), Inches(3.0), Inches(2.0), Inches(2.0))
    poly_a.name = "!!Morph_Polygon"
    poly_a.fill.solid()
    poly_a.fill.fore_color.rgb = RGBColor(0, 114, 63)

    # Slide 2: Morph to position B and deform to rectangle
    s2, _ = _new_slide(prs, "Case 38: Vector Polygon Morph Reassembly (Frame B)")
    apply_transition(s2, "morph")
    poly_b = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.0), Inches(3.0), Inches(3.5), Inches(2.0))
    poly_b.name = "!!Morph_Polygon"
    poly_b.fill.solid()
    poly_b.fill.fore_color.rgb = RGBColor(0, 114, 63)
    link_morph(poly_a, poly_b)
    prs.save(out_pptx)
    return out_pptx


def _build_case_39(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 39: Frosted Glass Depth Stack")
    # Base layer
    _card(s1, "Base_Backdrop", "ACRYLIC BACKDROP", 2.0, 2.5, 9.333, 4.0, rgb=(226, 232, 240))
    # Mid translucent card
    c_mid = _card(s1, "Mid_Card", "FROSTED GLASS CARD", 3.0, 3.0, 7.333, 3.0, rgb=(241, 245, 249))
    c_mid.line.color.rgb = RGBColor(14, 165, 233)
    c_mid.line.width = Pt(1.5)
    # Top card
    c_top = _card(s1, "Top_Card", "HERO FOCUS OVERLAY", 4.0, 3.5, 5.333, 2.0, rgb=(255, 255, 255))
    c_top.line.color.rgb = RGBColor(99, 102, 241)
    c_top.line.width = Pt(2.0)
    prs.save(out_pptx)
    return out_pptx


# --- GROUP H (Cases 40..44): End-to-End Cinematic Storyboard & Motion Director ---

def _build_case_40(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 40: Long Intra-Slide Sequence (15s Multi-Act)")
    card1 = _card(s1, "!!Act1_Card", "ACT 1: ESTABLISH (0-3s)", 1.5, 3.0, 3.2, 2.0, rgb=(14, 165, 233))
    card2 = _card(s1, "!!Act2_Card", "ACT 2: TRANSFORMATION (3-8s)", 5.1, 3.0, 3.2, 2.0, rgb=(99, 102, 241))
    card3 = _card(s1, "!!Act3_Card", "ACT 3: RESOLUTION (8-15s)", 8.7, 3.0, 3.2, 2.0, rgb=(16, 185, 129))

    tl = SlideAnimationTimeline(1)
    tl.add_effect(card1, "fade", duration_ms=1000, trigger="with_previous")
    tl.add_effect(card2, "fly_in", duration_ms=1000, trigger="after_previous", direction="up")
    tl.add_effect(card3, "zoom", duration_ms=1000, trigger="after_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_41(out_pptx: str, assets_dir: str) -> str:
    engine = StoryboardEngine()
    plan = engine.plan_presentation("Interplanetary Spacecraft Robotics", num_slides=3)
    director = MotionDirector(Themes.LIGHT_ACADEMIC_CYAN)
    return director.compile_plan_to_deck(plan, out_pptx)


def _build_case_42(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    prev_hero = None
    coords = [(1.5, 3.0, 1.2), (4.5, 2.2, 1.6), (8.0, 4.0, 1.4), (10.5, 3.0, 1.2)]
    for idx, (hx, hy, sz) in enumerate(coords):
        s, _ = _new_slide(prs, f"Case 42: Traveling Entity Scene {idx + 1}")
        apply_transition(s, "none" if idx == 0 else "morph")
        h = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(hx), Inches(hy), Inches(sz), Inches(sz))
        h.name = "!!Thematic_Traveler_Hero"
        h.fill.solid()
        h.fill.fore_color.rgb = RGBColor(14, 165, 233)
        if prev_hero:
            link_morph(prev_hero, h)
        prev_hero = h
    prs.save(out_pptx)
    return out_pptx


def _build_case_43(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 43: Multi-Target Synchronized Parallel Animation")
    tl = SlideAnimationTimeline(1)
    for col in range(3):
        for row in range(4):
            cx = 1.5 + col * 3.8
            cy = 2.4 + row * 1.1
            cell = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(cy), Inches(3.2), Inches(0.9))
            cell.name = f"Cell_{col}_{row}"
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(14, 165, 233)
            cell.line.fill.background()
            tl.add_effect(cell, "wipe", duration_ms=400, delay_ms=row * 100, direction="left", trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx


def _build_case_44(out_pptx: str, assets_dir: str) -> str:
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s1, _ = _new_slide(prs, "Case 44: Clean PowerPoint Open COM Verification")
    tl = SlideAnimationTimeline(1)
    for i in range(20):
        gx = 1.0 + (i % 5) * 2.3
        gy = 2.5 + (i // 5) * 1.1
        box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(gx), Inches(gy), Inches(2.0), Inches(0.8))
        box.name = f"CleanBox_{i}"
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(79, 70, 229)
        box.line.fill.background()
        tl.add_effect(box, "fade", duration_ms=300, delay_ms=(i % 5) * 80, trigger="with_previous")
    tl.compile_to_openxml(s1)
    prs.save(out_pptx)
    return out_pptx
