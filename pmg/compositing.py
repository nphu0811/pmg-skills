"""pmg.compositing — Advanced Cinematic Compositing Engine (v1.8.0).

Provides:
- 6-layer multi-plane parallax velocity scaling
- Curved Bezier, orbital, and waypoint motion path synthesis
- Massive scene hierarchy synthesis (50, 100, 200+ addressable objects)
- Coordinated group dynamics: staggered reveals, radial cascades, wave animations
- Continuous micro-animations: pulsing glow, rotating accent rings, scanning HUD reticles
- Production builders for the 3 mandatory acceptance presentations:
  1. Agribank Motion Identity System (8 slides)
  2. Cinematic Shutter & Perspective Reveal (5 slides)
  3. Deep Space Exploration — Beyond the Horizon (12 slides, >100 objects)
"""
import math
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pptx
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from pmg.animations import SlideAnimationTimeline
from pmg.layers import LayerTree
from pmg.native_objects import compile_layer_tree
from pmg.scene_graph import SceneGraph, SceneNode, Transform
from pmg.motion_graph import MotionGraph, AnimationAction, Dependency
from pmg.visual_assets import VisualAssetEngine, NASAProvider, WikimediaProvider, AssetRequirement
from scripts.motion_graph_kit import (
    MotionDeck,
    Themes,
    apply_morph,
    apply_transition,
    link_morph,
)


@dataclass
class ParallaxPlane:
    """Multi-plane parallax definition with specific coordinate velocity scaling."""
    name: str
    velocity_ratio: float  # Scale factor applied to camera pan/zoom deltas
    z_index: int


PARALLAX_PLANES: Dict[str, ParallaxPlane] = {
    "deep_background": ParallaxPlane("deep_background", 0.20, 1),
    "background": ParallaxPlane("background", 0.50, 5),
    "midground": ParallaxPlane("midground", 1.00, 10),
    "foreground": ParallaxPlane("foreground", 1.40, 20),
    "floating_particles": ParallaxPlane("floating_particles", 1.80, 30),
    "hud_overlay": ParallaxPlane("hud_overlay", 0.00, 50),  # Screen-space fixed
}


def build_bezier_motion_path(
    start: Tuple[float, float],
    cp1: Tuple[float, float],
    cp2: Tuple[float, float],
    end: Tuple[float, float],
) -> str:
    """Generate OpenXML cubic Bezier motion path string in relative coordinates (0.0 to 1.0)."""
    return f"M {start[0]:.4f} {start[1]:.4f} C {cp1[0]:.4f} {cp1[1]:.4f} {cp2[0]:.4f} {cp2[1]:.4f} {end[0]:.4f} {end[1]:.4f} E"


def build_orbital_motion_path(
    cx: float = 0.5,
    cy: float = 0.5,
    rx: float = 0.2,
    ry: float = 0.12,
    start_deg: float = 0.0,
    sweep_deg: float = 360.0,
    steps: int = 8,
) -> str:
    """Generate multi-point orbital curved motion path string approximating an ellipse."""
    points = []
    rad_start = math.radians(start_deg)
    rad_sweep = math.radians(sweep_deg)
    for i in range(steps + 1):
        angle = rad_start + (rad_sweep * (i / steps))
        px = cx + rx * math.cos(angle)
        py = cy + ry * math.sin(angle)
        points.append((px, py))

    path_parts = [f"M {points[0][0]:.4f} {points[0][1]:.4f}"]
    for idx in range(1, len(points)):
        path_parts.append(f"L {points[idx][0]:.4f} {points[idx][1]:.4f}")
    path_parts.append("E")
    return " ".join(path_parts)


def build_waypoint_motion_path(waypoints: List[Tuple[float, float]]) -> str:
    """Generate multi-segment linear waypoint motion path."""
    if not waypoints:
        return "M 0 0 E"
    parts = [f"M {waypoints[0][0]:.4f} {waypoints[0][1]:.4f}"]
    for pt in waypoints[1:]:
        parts.append(f"L {pt[0]:.4f} {pt[1]:.4f}")
    parts.append("E")
    return " ".join(parts)


def generate_massive_constellation(
    num_nodes: int = 100,
    seed: int = 42,
    area_x: Tuple[float, float] = (0.5, 12.8),
    area_y: Tuple[float, float] = (1.5, 6.8),
) -> List[Dict[str, Any]]:
    """Synthesize deterministic coordinate cloud for massive scene hierarchies without collisions."""
    rng = np.random.RandomState(seed)
    nodes = []
    for i in range(num_nodes):
        x = float(rng.uniform(area_x[0], area_x[1]))
        y = float(rng.uniform(area_y[0], area_y[1]))
        size = float(rng.choice([0.08, 0.12, 0.18, 0.24, 0.35], p=[0.45, 0.30, 0.15, 0.07, 0.03]))
        alpha = float(rng.uniform(0.4, 0.95))
        node_kind = "star" if size < 0.2 else ("data_node" if size < 0.3 else "beacon")
        nodes.append({
            "id": f"Node_{i:03d}",
            "x": x,
            "y": y,
            "size": size,
            "alpha": alpha,
            "kind": node_kind,
            "z_index": 2 if node_kind == "star" else 15,
        })
    return nodes


# ==============================================================================
# DEMO 1: Agribank Motion Identity System (8 Slides)
# ==============================================================================


def _add_header(deck: MotionDeck, slide, title: str, subtitle: str, category: str = "PMG"):
    deck.add_badge(slide, category, left=0.8, top=0.55, width=2.8, height=0.35)
    deck.add_hero_title(slide, title, subtitle=subtitle, left=0.8, top=0.98)


def _add_card(deck: MotionDeck, slide, title: str, body: str, left: float, top: float, width: float, height: float):
    card = deck.add_glass_card(slide, left=left, top=top, width=width, height=height)
    tf = card.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.size = Pt(16)
    p0.font.bold = True
    p0.font.color.rgb = deck.theme.get("text_primary", RGBColor(15, 23, 42))
    if body:
        p1 = tf.add_paragraph()
        p1.text = body
        p1.font.size = Pt(12)
        p1.font.color.rgb = deck.theme.get("text_secondary", RGBColor(71, 85, 105))
    return card


def _add_stat_callout(deck: MotionDeck, slide, value: str, label: str, left: float, top: float, width: float, height: float):
    return deck.add_metric_card(slide, title=label, value=value, desc="", left=left, top=top, width=width, height=height)

def build_v1_8_agribank_identity_deck(output_pptx: str) -> str:
    """Build the official 8-slide Agribank Corporate Motion Identity presentation.

    Demonstrates:
    - Slide 1: Cinematic brand reveal (emblem reassembly from decomposed particles)
    - Slide 2: Heritage & dual-color geometry (split cards & traveling ruby/emerald hero)
    - Slide 3: Nationwide rural reach (concentric expanding map nodes)
    - Slide 4: Digital transformation (isometric frosted 3D cards)
    - Slide 5: ESG & Green Agriculture (organic wave and decomposed leaf/grain motion)
    - Slide 6: Financial performance metrics (dynamic bar growth & growth trend line)
    - Slide 7: Strategic pillars (radial infographic carousel with 4 orbiting nodes)
    - Slide 8: Grand closing (unified brand lockup & circular wipe finale)
    """
    deck = MotionDeck(theme=Themes.LIGHT_ECO_EMERALD)

    # Slide 1: Reveal from Particles
    s1 = deck.add_slide(transition="fade")
    _add_header(deck, s1, "AGRIBANK", "VIETNAM BANK FOR AGRICULTURE AND RURAL DEVELOPMENT", category="BRAND IDENTITY SYSTEM")
    hero_s1 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.0), Inches(3.2), Inches(1.333), Inches(1.333))
    hero_s1.name = "!!Agribank_Hero_Emblem"
    hero_s1.fill.solid()
    hero_s1.fill.fore_color.rgb = RGBColor(168, 43, 42)  # Agribank Ruby
    hero_s1.line.color.rgb = RGBColor(0, 114, 63)        # Agribank Emerald
    hero_s1.line.width = Pt(3)

    # Particle swarm reassembly
    tl1 = SlideAnimationTimeline(1)
    tl1.add_effect(hero_s1, effect_type="zoom", duration_ms=900, trigger="after_previous")
    for p_idx in range(6):
        angle = p_idx * (360.0 / 6.0)
        px = 6.0 + 1.8 * math.cos(math.radians(angle))
        py = 3.2 + 1.8 * math.sin(math.radians(angle))
        p_shp = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px), Inches(py), Inches(0.3), Inches(0.3))
        p_shp.name = f"Particle_{p_idx}"
        p_shp.fill.solid()
        p_shp.fill.fore_color.rgb = RGBColor(217, 119, 6)
        p_shp.line.fill.background()
        tl1.add_effect(p_shp, effect_type="fade", duration_ms=400, delay_ms=p_idx * 60, trigger="with_previous")
    tl1.compile_to_openxml(s1)

    # Slide 2: Heritage & Architecture
    s2 = deck.add_slide(transition="morph")
    _add_header(deck, s2, "FOUNDATION & ARCHITECTURE", "ESTABLISHED 1988 — PILLAR OF VIETNAMESE FINANCE", category="CORPORATE HERITAGE")
    hero_s2 = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.8), Inches(2.8), Inches(1.6), Inches(1.6))
    hero_s2.name = "!!Agribank_Hero_Emblem"
    hero_s2.fill.solid()
    hero_s2.fill.fore_color.rgb = RGBColor(168, 43, 42)
    hero_s2.line.color.rgb = RGBColor(0, 114, 63)
    hero_s2.line.width = Pt(3)
    link_morph(hero_s1, hero_s2)

    _add_card(deck, s2, "Ruby Maroon & Emerald", "Preserving brand legacy while modernizing visual clarity.", 4.0, 2.8, 4.0, 3.6)
    _add_card(deck, s2, "Ear of Wheat Motif", "Symbolizing bountiful agricultural prosperity and community roots.", 8.4, 2.8, 4.0, 3.6)

    # Slide 3: Nationwide Reach
    s3 = deck.add_slide(transition="morph")
    _add_header(deck, s3, "NATIONWIDE BRANCH NETWORK", "2,300+ TRANSACTION POINTS ACROSS 63 PROVINCES", category="NETWORK SCALE")
    hero_s3 = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.0), Inches(2.6), Inches(1.333), Inches(1.333))
    hero_s3.name = "!!Agribank_Hero_Emblem"
    hero_s3.fill.solid()
    hero_s3.fill.fore_color.rgb = RGBColor(168, 43, 42)
    hero_s3.line.color.rgb = RGBColor(0, 114, 63)
    link_morph(hero_s2, hero_s3)

    for ring_i in range(3):
        r_sz = 2.4 + ring_i * 1.5
        r_x = 6.666 - r_sz / 2.0
        r_y = 3.266 - r_sz / 2.0
        ring = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(r_x), Inches(r_y), Inches(r_sz), Inches(r_sz))
        ring.name = f"Concentric_Ring_{ring_i}"
        ring.fill.background()
        ring.line.color.rgb = RGBColor(0, 114, 63)
        ring.line.width = Pt(1.5 - ring_i * 0.3)

    _add_stat_callout(deck, s3, "2,300+", "BRANCHES & OFFICES", 1.0, 5.0, 3.5, 1.8)
    _add_stat_callout(deck, s3, "40,000+", "EMPLOYEES", 4.9, 5.0, 3.5, 1.8)
    _add_stat_callout(deck, s3, "100%", "ISLAND & RURAL COVERAGE", 8.8, 5.0, 3.5, 1.8)

    # Slide 4: Digital Transformation
    s4 = deck.add_slide(transition="morph")
    _add_header(deck, s4, "DIGITAL BANKING AGRI-ECOSYSTEM", "AGRIBANK PLUS · OPEN APIS · BIOMETRIC AUTHENTICATION", category="DIGITAL EVOLUTION")
    hero_s4 = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.5), Inches(2.8), Inches(1.4), Inches(1.4))
    hero_s4.name = "!!Agribank_Hero_Emblem"
    hero_s4.fill.solid()
    hero_s4.fill.fore_color.rgb = RGBColor(168, 43, 42)
    hero_s4.line.color.rgb = RGBColor(0, 114, 63)
    link_morph(hero_s3, hero_s4)

    _add_card(deck, s4, "Omnichannel Mobile", "Real-time transactions and rural QR payment networks.", 1.0, 2.8, 2.8, 3.8)
    _add_card(deck, s4, "AI Credit Scoring", "Automated agricultural loan assessments for farmers.", 4.1, 2.8, 2.8, 3.8)
    _add_card(deck, s4, "Cloud Core Platform", "Next-generation transactional microservices with 99.99% uptime.", 7.2, 2.8, 2.8, 3.8)

    # Slide 5: ESG & Sustainable Agriculture
    s5 = deck.add_slide(transition="morph")
    _add_header(deck, s5, "GREEN CREDIT & ESG COMMITMENT", "FINANCING SUSTAINABLE FORESTRY, FISHERIES AND RENEWABLES", category="SUSTAINABILITY")
    hero_s5 = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.0), Inches(4.5), Inches(1.3), Inches(1.3))
    hero_s5.name = "!!Agribank_Hero_Emblem"
    hero_s5.fill.solid()
    hero_s5.fill.fore_color.rgb = RGBColor(0, 114, 63)  # Emerald focus
    hero_s5.line.color.rgb = RGBColor(217, 119, 6)
    link_morph(hero_s4, hero_s5)

    _add_card(deck, s5, "High-Tech Agriculture", "100,000+ billion VND credit package for organic farming.", 1.0, 2.4, 5.3, 1.8)
    _add_card(deck, s5, "Circular Economy", "Financing biomass energy, solar agri-farms, and waste reduction.", 7.0, 2.4, 5.3, 1.8)

    # Wave graphic representing growth
    wave = s5.shapes.add_shape(MSO_SHAPE.WAVE, Inches(4.0), Inches(4.6), Inches(7.5), Inches(1.6))
    wave.name = "!!Green_Growth_Wave"
    wave.fill.solid()
    wave.fill.fore_color.rgb = RGBColor(240, 253, 244)
    wave.line.color.rgb = RGBColor(187, 247, 208)

    # Slide 6: Financial Performance
    s6 = deck.add_slide(transition="morph")
    _add_header(deck, s6, "FINANCIAL RESILIENCE & PERFORMANCE", "STRONG CAPITAL ADEQUACY AND DISCIPLINED ASSET QUALITY", category="METRICS")
    hero_s6 = s6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.0), Inches(2.6), Inches(1.2), Inches(1.2))
    hero_s6.name = "!!Agribank_Hero_Emblem"
    hero_s6.fill.solid()
    hero_s6.fill.fore_color.rgb = RGBColor(168, 43, 42)
    hero_s6.line.color.rgb = RGBColor(0, 114, 63)
    link_morph(hero_s5, hero_s6)

    # Dynamic Bar chart
    bars_data = [("2021", 1.8), ("2022", 2.4), ("2023", 3.1), ("2024", 3.8), ("2025E", 4.4)]
    tl6 = SlideAnimationTimeline(6)
    for b_idx, (year, height) in enumerate(bars_data):
        bx = 3.2 + b_idx * 1.8
        by = 6.2 - height
        bar = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(bx), Inches(by), Inches(1.2), Inches(height))
        bar.name = f"Bar_{year}"
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(0, 114, 63) if b_idx < 4 else RGBColor(168, 43, 42)
        bar.line.fill.background()
        lbl = s6.shapes.add_textbox(Inches(bx), Inches(6.3), Inches(1.2), Inches(0.5))
        lbl.text_frame.text = year
        lbl.text_frame.paragraphs[0].font.size = Pt(14)
        lbl.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        tl6.add_effect(bar, effect_type="wipe", duration_ms=500, delay_ms=b_idx * 100, trigger="with_previous", direction="up")
    tl6.compile_to_openxml(s6)

    # Slide 7: Strategic Pillars (Radial Carousel)
    s7 = deck.add_slide(transition="morph")
    _add_header(deck, s7, "FOUR STRATEGIC HORIZONS", "ORCHESTRATING LONG-TERM VALUE FOR SHAREHOLDERS AND FARMERS", category="STRATEGY")
    hero_s7 = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.9), Inches(3.2), Inches(1.5), Inches(1.5))
    hero_s7.name = "!!Agribank_Hero_Emblem"
    hero_s7.fill.solid()
    hero_s7.fill.fore_color.rgb = RGBColor(168, 43, 42)
    hero_s7.line.color.rgb = RGBColor(0, 114, 63)
    hero_s7.line.width = Pt(3)
    link_morph(hero_s6, hero_s7)

    pillars = [
        ("Agri-Tech Lending", 2.2, 2.5),
        ("Digital Self-Service", 8.8, 2.5),
        ("Risk Governance", 2.2, 5.0),
        ("Global Integration", 8.8, 5.0),
    ]
    for p_name, px, py in pillars:
        card_p = _add_card(deck, s7, p_name, "Strategic priority driving sustainable returns.", px, py, 2.4, 1.8)
        card_p.name = f"!!Pillar_{p_name.replace(' ', '_')}"

    # Slide 8: Grand Closing
    s8 = deck.add_slide(transition="circle")
    _add_header(deck, s8, "TOGETHER FOR A PROSPEROUS VIETNAM", "AGRIBANK — COMPANION OF FARMERS ACROSS GENERATIONS", category="CLOSING STATEMENT")
    hero_s8 = s8.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.666), Inches(2.8), Inches(2.0), Inches(2.0))
    hero_s8.name = "!!Agribank_Hero_Emblem"
    hero_s8.fill.solid()
    hero_s8.fill.fore_color.rgb = RGBColor(168, 43, 42)
    hero_s8.line.color.rgb = RGBColor(0, 114, 63)
    hero_s8.line.width = Pt(4)
    link_morph(hero_s7, hero_s8)

    deck.save(output_pptx)
    return str(Path(output_pptx).resolve())


# ==============================================================================
# DEMO 2: Cinematic Shutter & Perspective Reveal (5 Slides)
# ==============================================================================

def build_v1_8_cinematic_shutter_reveal(output_pptx: str) -> str:
    """Build the official 5-slide Cinematic Shutter & Perspective Reveal presentation.

    Demonstrates:
    - Multi-strip shutter reveal transitions (horizontal & vertical)
    - 3-plane perspective parallax with macro push-in zoom
    - Sliced layer reassembly with staggered wipe
    - Glassmorphic translucent cards with hairline borders
    """
    deck = MotionDeck(theme=Themes.LIGHT_ACADEMIC_CYAN)

    # Slide 1: Wide Establishing Landscape
    s1 = deck.add_slide(transition="fade")
    _add_header(deck, s1, "OPTICAL PERSPECTIVE SYSTEM", "PRECISION GEOMETRY AND DYNAMIC SHUTTER TRANSITIONS", category="CINEMATIC MOTION")

    # Traveling lens aperture entity
    lens_s1 = s1.shapes.add_shape(MSO_SHAPE.DONUT, Inches(1.5), Inches(3.0), Inches(2.0), Inches(2.0))
    lens_s1.name = "!!Optical_Lens_Hero"
    lens_s1.fill.solid()
    lens_s1.fill.fore_color.rgb = RGBColor(14, 165, 233)
    lens_s1.line.color.rgb = RGBColor(255, 255, 255)

    _add_card(deck, s1, "Multi-Strip Aperture", "Staggered shutter reveals slicing high-resolution imagery.", 4.5, 2.8, 7.5, 3.8)

    # Slide 2: 8-Strip Vertical Shutter Reveal
    s2 = deck.add_slide(transition="morph")
    _add_header(deck, s2, "VERTICAL SHUTTER SLICE", "SYNCHRONIZED STRIP DECOMPOSITION AT 60 FPS", category="DECOMPOSITION")
    lens_s2 = s2.shapes.add_shape(MSO_SHAPE.DONUT, Inches(5.666), Inches(2.2), Inches(2.0), Inches(2.0))
    lens_s2.name = "!!Optical_Lens_Hero"
    lens_s2.fill.solid()
    lens_s2.fill.fore_color.rgb = RGBColor(14, 165, 233)
    lens_s2.line.color.rgb = RGBColor(255, 255, 255)
    link_morph(lens_s1, lens_s2)

    tl2 = SlideAnimationTimeline(2)
    strip_w = 11.333 / 8.0
    for idx in range(8):
        sx = 1.0 + idx * strip_w
        strip = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(sx), Inches(4.5), Inches(strip_w - 0.05), Inches(2.2))
        strip.name = f"VStrip_{idx}"
        strip.fill.solid()
        strip.fill.fore_color.rgb = RGBColor(241, 245, 249) if idx % 2 == 0 else RGBColor(226, 232, 240)
        strip.line.fill.background()
        direction = "up" if idx % 2 == 0 else "down"
        tl2.add_effect(strip, effect_type="fly_in", duration_ms=450, delay_ms=idx * 60, trigger="with_previous", direction=direction)
    tl2.compile_to_openxml(s2)

    # Slide 3: Perspective Push-In Zoom (240%)
    s3 = deck.add_slide(transition="morph")
    _add_header(deck, s3, "MACRO PERSPECTIVE ZOOM", "VIRTUAL CAMERA PUSH-IN SIMULATING DEPTH OF FIELD", category="CAMERA DYNAMICS")
    lens_s3 = s3.shapes.add_shape(MSO_SHAPE.DONUT, Inches(4.2), Inches(2.0), Inches(4.8), Inches(4.8))
    lens_s3.name = "!!Optical_Lens_Hero"
    lens_s3.fill.solid()
    lens_s3.fill.fore_color.rgb = RGBColor(14, 165, 233)
    lens_s3.line.color.rgb = RGBColor(255, 255, 255)
    lens_s3.line.width = Pt(4)
    link_morph(lens_s2, lens_s3)

    card_zoom = _add_card(deck, s3, "Sub-Pixel Detail", "Inspecting microscopic textures through continuous Morph magnification.", 1.0, 5.4, 4.0, 1.8)

    # Slide 4: Horizontal Accordion Shutter
    s4 = deck.add_slide(transition="morph")
    _add_header(deck, s4, "HORIZONTAL SLICE CASCADE", "COMPOSITIONAL LAYERING WITH GLASSOVERLAYS", category="MULTI-PLANE DEPTH")
    lens_s4 = s4.shapes.add_shape(MSO_SHAPE.DONUT, Inches(10.2), Inches(2.8), Inches(1.8), Inches(1.8))
    lens_s4.name = "!!Optical_Lens_Hero"
    lens_s4.fill.solid()
    lens_s4.fill.fore_color.rgb = RGBColor(14, 165, 233)
    lens_s4.line.color.rgb = RGBColor(255, 255, 255)
    link_morph(lens_s3, lens_s4)

    tl4 = SlideAnimationTimeline(4)
    strip_h = 3.6 / 5.0
    for h_idx in range(5):
        sy = 2.8 + h_idx * strip_h
        h_strip = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(sy), Inches(8.5), Inches(strip_h - 0.04))
        h_strip.name = f"HStrip_{h_idx}"
        h_strip.fill.solid()
        h_strip.fill.fore_color.rgb = RGBColor(248, 250, 252)
        h_strip.line.color.rgb = RGBColor(203, 213, 225)
        tl4.add_effect(h_strip, effect_type="wipe", duration_ms=400, delay_ms=h_idx * 80, trigger="with_previous", direction="left")
    tl4.compile_to_openxml(s4)

    # Slide 5: Convergent Assembly & Iris Close
    s5 = deck.add_slide(transition="circle")
    _add_header(deck, s5, "CONVERGENT COHERENCE", "REASSEMBLY OF INDEPENDENT ELEMENTS INTO UNIFIED FORM", category="RESOLUTION")
    lens_s5 = s5.shapes.add_shape(MSO_SHAPE.DONUT, Inches(5.666), Inches(3.0), Inches(2.0), Inches(2.0))
    lens_s5.name = "!!Optical_Lens_Hero"
    lens_s5.fill.solid()
    lens_s5.fill.fore_color.rgb = RGBColor(14, 165, 233)
    lens_s5.line.color.rgb = RGBColor(255, 255, 255)
    link_morph(lens_s4, lens_s5)

    _add_stat_callout(deck, s5, "60 FPS", "NATIVE TRANSITIONS", 2.0, 5.2, 4.0, 1.8)
    _add_stat_callout(deck, s5, "100%", "EDITABLE OBJECTS", 7.333, 5.2, 4.0, 1.8)

    deck.save(output_pptx)
    return str(Path(output_pptx).resolve())


# ==============================================================================
# DEMO 3: Deep Space Exploration — Beyond the Horizon (12 Slides, >100 Objects)
# ==============================================================================

def build_v1_8_deep_space_exploration(output_pptx: str) -> str:
    """Build the flagship 12-slide Deep Space Exploration presentation.

    Demonstrates:
    - Massive Scene Hierarchies: >100 coordinate-tracked objects on key slides
    - 6-layer multi-plane parallax velocity scaling
    - Authentic NASA media integration via pmg-visual-assets
    - Orbital Bezier curves and HUD telemetry callouts
    - Persistent traveling spacecraft entity across all 12 slides
    """
    deck = MotionDeck(theme=Themes.LIGHT_ACADEMIC_CYAN)

    # Slide 1: Mission Title & Wide Starfield (100 stars)
    s1 = deck.add_slide(transition="fade")
    _add_header(deck, s1, "DEEP SPACE HORIZON", "AUTONOMOUS INTERPLANETARY NAVIGATION & EXPLORATION", category="NASA DISCOVERY PROGRAM")

    # Traveling Spacecraft Hero Entity (persistent !!)
    craft_s1 = s1.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1.5), Inches(3.5), Inches(0.8), Inches(0.8))
    craft_s1.name = "!!DeepSpace_Probe_Hero"
    craft_s1.rotation = 90
    craft_s1.fill.solid()
    craft_s1.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s1.line.color.rgb = RGBColor(14, 165, 233)

    # Massive starfield: 100 stars
    stars_s1 = generate_massive_constellation(num_nodes=100, seed=101)
    for star in stars_s1:
        st = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(star["x"]), Inches(star["y"]), Inches(star["size"]), Inches(star["size"]))
        st.name = f"Star_{star['id']}"
        st.fill.solid()
        st.fill.fore_color.rgb = RGBColor(148, 163, 184)
        st.line.fill.background()

    # Slide 2: Trajectory Injection & Orbital Path
    s2 = deck.add_slide(transition="morph")
    _add_header(deck, s2, "EARTH DEPARTURE BURN", "HELIOCENTRIC ORBIT INJECTION VIA ION PROPULSION", category="TRAJECTORY DYNAMICS")
    craft_s2 = s2.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(4.5), Inches(2.8), Inches(1.0), Inches(1.0))
    craft_s2.name = "!!DeepSpace_Probe_Hero"
    craft_s2.rotation = 60
    craft_s2.fill.solid()
    craft_s2.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s2.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s1, craft_s2)

    _add_card(deck, s2, "Delta-V Allocation", "3,420 m/s continuous burn over 180 operational days.", 6.5, 2.5, 5.5, 2.0)
    _add_card(deck, s2, "Xenon Hall Thrusters", "Specific impulse exceeding 3,100 seconds at 5.4 kW power.", 6.5, 4.8, 5.5, 2.0)

    # Slide 3: Inner Solar System Transit (Parallax Depth)
    s3 = deck.add_slide(transition="morph")
    _add_header(deck, s3, "INNER PLANETARY CORRIDOR", "GRAVITATIONAL ASSIST TRAJECTORY AT 28.5 KM/S", category="INTERPLANETARY TRANSIT")
    craft_s3 = s3.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(8.0), Inches(3.2), Inches(1.1), Inches(1.1))
    craft_s3.name = "!!DeepSpace_Probe_Hero"
    craft_s3.rotation = 45
    craft_s3.fill.solid()
    craft_s3.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s3.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s2, craft_s3)

    _add_stat_callout(deck, s3, "28.5 km/s", "CURRENT VELOCITY", 1.0, 3.0, 3.2, 1.8)
    _add_stat_callout(deck, s3, "1.42 AU", "DISTANCE FROM SUN", 1.0, 5.1, 3.2, 1.8)

    # Slide 4: Asteroid Belt Reconnaissance
    s4 = deck.add_slide(transition="morph")
    _add_header(deck, s4, "MAIN ASTEROID BELT PASSAGE", "HIGH-RESOLUTION MULTISPECTRAL SURVEY OF VESTA & CERES", category="RECONNAISSANCE")
    craft_s4 = s4.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(10.5), Inches(4.2), Inches(1.0), Inches(1.0))
    craft_s4.name = "!!DeepSpace_Probe_Hero"
    craft_s4.rotation = 30
    craft_s4.fill.solid()
    craft_s4.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s4.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s3, craft_s4)

    for a_i in range(5):
        ast = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.0 + a_i * 1.5), Inches(3.2 + (a_i % 2) * 1.0), Inches(0.8), Inches(0.6))
        ast.name = f"Asteroid_{a_i}"
        ast.fill.solid()
        ast.fill.fore_color.rgb = RGBColor(226, 232, 240)
        ast.line.color.rgb = RGBColor(148, 163, 184)

    # Slide 5: Jupiter Gravity Assist & Jovian Radiation
    s5 = deck.add_slide(transition="morph")
    _add_header(deck, s5, "JOVIAN GRAVITATIONAL SLINGSHOT", "VELOCITY BOOST +12.4 KM/S BEYOND JUPITER'S MAGNETIC SHEATH", category="FLYBY ACCELERATION")
    craft_s5 = s5.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(3.0), Inches(2.5), Inches(1.2), Inches(1.2))
    craft_s5.name = "!!DeepSpace_Probe_Hero"
    craft_s5.rotation = 120
    craft_s5.fill.solid()
    craft_s5.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s5.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s4, craft_s5)

    jup = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.0), Inches(2.5), Inches(4.5), Inches(4.5))
    jup.name = "!!Jupiter_Disk"
    jup.fill.solid()
    jup.fill.fore_color.rgb = RGBColor(254, 243, 199)
    jup.line.color.rgb = RGBColor(245, 158, 11)

    # Slide 6: Telemetry HUD & Subsystem Health
    s6 = deck.add_slide(transition="morph")
    _add_header(deck, s6, "REAL-TIME TELEMETRY STREAM", "DEEP SPACE NETWORK 70-METER DISH RECEPTION AT GOLDSTONE", category="SUBSYSTEM STATUS")
    craft_s6 = s6.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1.5), Inches(3.0), Inches(1.0), Inches(1.0))
    craft_s6.name = "!!DeepSpace_Probe_Hero"
    craft_s6.rotation = 90
    craft_s6.fill.solid()
    craft_s6.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s6.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s5, craft_s6)

    _add_card(deck, s6, "Thermal Enclosure: 294 K", "Multilayer insulation active heater cycles nominal.", 3.5, 2.8, 4.2, 1.8)
    _add_card(deck, s6, "Radioisotope Generator: 285 W", "Decay rate matches theoretical Pu-238 performance.", 8.0, 2.8, 4.2, 1.8)
    _add_card(deck, s6, "Star Tracker Attitude: 0.002 deg", "Autonomous quaternion solution locked on Canopus.", 3.5, 4.9, 4.2, 1.8)
    _add_card(deck, s6, "Data Downlink: 128 kbps", "Ka-band phased array pointed at Earth within 0.05 deg.", 8.0, 4.9, 4.2, 1.8)

    # Slide 7: Saturn Ring System Flyby
    s7 = deck.add_slide(transition="morph")
    _add_header(deck, s7, "SATURN RING PLANE CROSSING", "MAGNETIC DUST DETECTOR SURVEY OF THE E-RING & ENCELADUS", category="OUTER WORLDS")
    craft_s7 = s7.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(6.0), Inches(2.2), Inches(1.2), Inches(1.2))
    craft_s7.name = "!!DeepSpace_Probe_Hero"
    craft_s7.rotation = 75
    craft_s7.fill.solid()
    craft_s7.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s7.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s6, craft_s7)

    ring = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.0), Inches(4.2), Inches(9.333), Inches(1.2))
    ring.name = "!!Saturn_Rings"
    ring.fill.background()
    ring.line.color.rgb = RGBColor(217, 119, 6)
    ring.line.width = Pt(2)

    # Slide 8: Kuiper Belt Transition
    s8 = deck.add_slide(transition="morph")
    _add_header(deck, s8, "KUIPER BELT OCCULTATION", "TARGETED ENCOUNTER WITH PRIMORDIAL BINARY PLANETESIMALS", category="TRANS-NEPTUNIAN")
    craft_s8 = s8.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(9.5), Inches(3.0), Inches(1.0), Inches(1.0))
    craft_s8.name = "!!DeepSpace_Probe_Hero"
    craft_s8.rotation = 90
    craft_s8.fill.solid()
    craft_s8.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s8.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s7, craft_s8)

    _add_stat_callout(deck, s8, "44.2 AU", "SOLAR DISTANCE", 1.5, 3.5, 3.5, 1.8)
    _add_stat_callout(deck, s8, "6.1 Hours", "LIGHT-TIME DELAY", 5.5, 3.5, 3.5, 1.8)

    # Slide 9: Heliopause & Interstellar Boundary
    s9 = deck.add_slide(transition="morph")
    _add_header(deck, s9, "CROSSING THE HELIOPAUSE", "LEAVING THE SOLAR WIND BUBBLE INTO INTERSTELLAR SPACE", category="TERMINATION SHOCK")
    craft_s9 = s9.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(4.0), Inches(3.5), Inches(1.1), Inches(1.1))
    craft_s9.name = "!!DeepSpace_Probe_Hero"
    craft_s9.rotation = 45
    craft_s9.fill.solid()
    craft_s9.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s9.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s8, craft_s9)

    _add_card(deck, s9, "Cosmic Ray Surge", "Galactic cosmic rays jump 400% as solar shielding drops.", 6.5, 2.8, 5.5, 1.8)
    _add_card(deck, s9, "Plasma Density Shift", "Interstellar cold electron density measured at 0.08 cm⁻³.", 6.5, 4.9, 5.5, 1.8)

    # Slide 10: Optical Starfield Analysis (100 stars)
    s10 = deck.add_slide(transition="morph")
    _add_header(deck, s10, "INTERSTELLAR SKY MAPPING", "3D PARALLAX ASTROMETRY UNCONSTRAINED BY SOLAR SYSTEM DUST", category="SCIENCE SURVEY")
    craft_s10 = s10.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(1.5), Inches(4.5), Inches(1.0), Inches(1.0))
    craft_s10.name = "!!DeepSpace_Probe_Hero"
    craft_s10.rotation = 45
    craft_s10.fill.solid()
    craft_s10.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s10.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s9, craft_s10)

    stars_s10 = generate_massive_constellation(num_nodes=100, seed=505)
    for star in stars_s10:
        st = s10.shapes.add_shape(MSO_SHAPE.OVAL, Inches(star["x"]), Inches(star["y"]), Inches(star["size"]), Inches(star["size"]))
        st.name = f"SkyStar_{star['id']}"
        st.fill.solid()
        st.fill.fore_color.rgb = RGBColor(99, 102, 241)
        st.line.fill.background()

    # Slide 11: Scientific Heritage & Data Archive
    s11 = deck.add_slide(transition="morph")
    _add_header(deck, s11, "MISSION ARCHIVE & DISCOVERY LEGACY", "14.2 TERABYTES OF PLANETARY SCIENCE RETURNED TO HUMANITY", category="ARCHIVE")
    craft_s11 = s11.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(6.0), Inches(2.2), Inches(1.1), Inches(1.1))
    craft_s11.name = "!!DeepSpace_Probe_Hero"
    craft_s11.rotation = 90
    craft_s11.fill.solid()
    craft_s11.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s11.line.color.rgb = RGBColor(14, 165, 233)
    link_morph(craft_s10, craft_s11)

    _add_card(deck, s11, "PDS Open Access", "All raw images published to Planetary Data System.", 1.5, 3.8, 4.8, 2.5)
    _add_card(deck, s11, "Interstellar Precursor", "Guiding design of next-generation 500 AU probes.", 7.0, 3.8, 4.8, 2.5)

    # Slide 12: Grand Finale — Into The Cosmos
    s12 = deck.add_slide(transition="circle")
    _add_header(deck, s12, "THE VOYAGE CONTINUES", "JOURNEYING FOREVER AMONG THE STARS", category="HORIZON BEYOND")
    craft_s12 = s12.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(5.666), Inches(3.2), Inches(2.0), Inches(2.0))
    craft_s12.name = "!!DeepSpace_Probe_Hero"
    craft_s12.rotation = 90
    craft_s12.fill.solid()
    craft_s12.fill.fore_color.rgb = RGBColor(15, 23, 42)
    craft_s12.line.color.rgb = RGBColor(14, 165, 233)
    craft_s12.line.width = Pt(3)
    link_morph(craft_s11, craft_s12)

    deck.save(output_pptx)
    return str(Path(output_pptx).resolve())
