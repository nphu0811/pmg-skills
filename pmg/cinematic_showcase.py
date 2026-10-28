"""
pmg.cinematic_showcase
======================
Preserved cinematic presentation builders and reusable choreography patterns.

Implements the iconic presentation patterns extracted directly from DataPowerPoint references:
1. SpotifyAppShellBuilder:
   - Full interactive Spotify Application Shell / Dashboard unlock with Morph transitions (from MAC-LENIN.pptx)
   - Slide 1: Brand/Topic Monolith with expanding ambient glow orb & hero emblem entrance
   - Slide 2: Morph convergence (orb contracts & docks into corner, title & subtitle glide in from margins)
   - Slide 3: Full App Shell Unlock (Sidebar navigation rail with macOS 3-dots, active category indicator,
     top bar with history controls, action pill, user profile avatar, and central viewport with headline & pill buttons)
   - Slide 4: In-Frame Viewport Stream (Shell remains pinned in place while active nav switches and viewport morphs)
   - Slide 5: Multi-card Team / Feature Grid stream

2. EarthSpaceMotionBuilder:
   - 21-Slide Cinematic Motion System (from EARTH AND SPACE.pptx)
   - Cosmic / Macro establishing shot with celestial body / topic sphere traversing across canvas
   - Camera Pan & secondary orbital traverse
   - Targeting Reticle / Viewfinder Zoom Sequence: 4 L-shaped corner brackets [ ] locking onto target region
   - Viewfinder Expansion: reticle brackets expand to frame 3 highlight cards
   - Focal Question / Mystery Node Transition: centered circular node ("BUT SOMETHING IS MISSING") expanding into macro zoom
   - 4-Quadrant Rotating Segmented Radar Wheel Navigator: 4-slice circular wheel rotating 90 deg per slide (0->90->180->270)
   - Orbiting Hero Model with Glowing Radial HUD Telemetry Dial: 8-node circular dial revolving as slides advance

3. UniversityBloomBuilder:
   - Master Morph Concentric Bloom & 8-Petal Table of Contents (from PROMPT.pdf / PROMPT-CHUYỂN-NỘI-DUNG.docx)
   - Concentric expanding rings + diamond corner accents
   - 8 curved bezier petals blooming radially around center hub
   - Morph into 8-item two-column Table of Contents (Mục Lục)

All builders are fully parameterized and adapt cleanly to ANY topic, palette, and domain.
"""

import os
import math
from typing import List, Dict, Any, Optional, Tuple

import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml

from scripts.motion_graph_kit import (
    MotionDeck,
    Themes,
    CinematicColors,
    apply_transition,
    apply_morph,
)
from pmg.animations import SlideAnimationTimeline


# ---------------------------------------------------------------------------
# Helpers for shape manipulation and OpenXML styling
# ---------------------------------------------------------------------------

def _set_shape_rotation(shape, rot_deg: float) -> None:
    if rot_deg != 0:
        try:
            shape._element.spPr.xfrm.set("rot", str(int(rot_deg * 60000)))
        except Exception:
            pass


def _apply_solid_fill(shape, rgb: RGBColor, alpha: Optional[float] = None) -> None:
    if alpha is None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb
    else:
        hex_clr = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
        val_int = int(alpha * 100000)
        xml = (
            f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:srgbClr val="{hex_clr}"><a:alpha val="{val_int}"/></a:srgbClr>'
            f'</a:solidFill>'
        )
        shape._element.spPr.append(parse_xml(xml))


def _apply_hairline_border(shape, rgb: RGBColor, width_pt: float = 1.0, alpha: Optional[float] = None) -> None:
    hex_clr = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    width_emu = int(width_pt * 12700)
    alpha_xml = f'<a:alpha val="{int(alpha * 100000)}"/>' if alpha is not None else ""
    xml = (
        f'<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{width_emu}">'
        f'<a:solidFill><a:srgbClr val="{hex_clr}">{alpha_xml}</a:srgbClr></a:solidFill>'
        f'</a:ln>'
    )
    shape._element.spPr.append(parse_xml(xml))


def _set_slide_background(slide, color_hex: str) -> None:
    bg_xml = (
        f'<p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<p:bgPr><a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill><a:effectLst/></p:bgPr>'
        f'</p:bg>'
    )
    cSld = slide._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
    if cSld is not None:
        existing = cSld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}bg")
        if existing is not None:
            cSld.remove(existing)
        cSld.insert(0, parse_xml(bg_xml))


# ===========================================================================
# 1. SPOTIFY APP SHELL & RADIAL EXPANSION BUILDER (from MAC-LENIN.pptx)
# ===========================================================================

class SpotifyAppShellBuilder:
    """
    Constructs a 4-to-5 slide Spotify-style Application Shell & Radial Morph unlock.
    Universal: works for music, tech keynotes, SaaS apps, university team projects, or portfolios.
    """

    def __init__(
        self,
        topic_title: str = "SPOTIFY",
        topic_subtitle: str = "TEAM 8 PRESENTS",
        author_name: str = "Mr. Thinh",
        primary_accent_hex: str = "1ED760",      # Spotify signature vibrant green
        dark_bg_hex: str = "121212",             # Deep matte charcoal
        sidebar_bg_hex: str = "0A080F",          # Darker sidebar
        surface_card_hex: str = "181818",        # Subtle elevated card surface
        text_primary_rgb: RGBColor = RGBColor(255, 255, 255),
        text_muted_rgb: RGBColor = RGBColor(179, 179, 179),
        nav_items: Optional[List[str]] = None,
    ):
        self.topic_title = topic_title
        self.topic_subtitle = topic_subtitle
        self.author_name = author_name
        self.accent_hex = primary_accent_hex
        self.accent_rgb = RGBColor.from_string(primary_accent_hex)
        self.dark_bg_hex = dark_bg_hex
        self.sidebar_bg_hex = sidebar_bg_hex
        self.surface_card_hex = surface_card_hex
        self.text_primary = text_primary_rgb
        self.text_muted = text_muted_rgb
        self.nav_items = nav_items or [
            "Welcome", "About Us", "Our Teams", "Agenda", "Our Services",
            "Achievements", "Our Goals", "Gallery", "Pricing", "Contact Us",
        ]

    def build_deck(self, output_path: str, num_slides: int = 4) -> str:
        deck = MotionDeck(theme=Themes.SPOTIFY_CYBER)
        num_slides = max(3, min(num_slides, 5))

        # -------------------------------------------------------------------
        # SLIDE 1: The Brand / Topic Monolith
        # -------------------------------------------------------------------
        s1 = deck.add_slide(transition="none")
        _set_slide_background(s1, self.dark_bg_hex)

        # Giant expanding ambient orb
        orb1 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2.86), Inches(-2.33), Inches(17.92), Inches(10.95))
        orb1.name = "!!Spotify_AmbientOrb"
        _apply_solid_fill(orb1, self.accent_rgb, alpha=0.92)
        orb1.line.fill.background()

        # Hero emblem / topic icon (in center)
        hero_icon = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.83), Inches(2.61), Inches(1.67), Inches(1.66))
        hero_icon.name = "!!Spotify_HeroIcon"
        _apply_solid_fill(hero_icon, RGBColor(255, 255, 255))
        hero_icon.line.fill.background()
        tf = hero_icon.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = "●"
        p.font.size = Pt(36)
        p.font.color.rgb = self.accent_rgb
        p.alignment = PP_ALIGN.CENTER

        # Staged off-screen elements for Slide 2 convergence
        t1 = s1.shapes.add_textbox(Inches(14.5), Inches(3.14), Inches(4.5), Inches(1.2))
        t1.name = "!!AppShell_Title"
        t1.text_frame.text = self.topic_title

        sub1 = s1.shapes.add_textbox(Inches(14.5), Inches(4.2), Inches(4.5), Inches(0.8))
        sub1.name = "!!AppShell_Subtitle"
        sub1.text_frame.text = self.topic_subtitle

        pill_l = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(-2.5), Inches(6.6), Inches(2.17), Inches(0.64))
        pill_l.name = "!!AppShell_PillL"
        _apply_solid_fill(pill_l, self.accent_rgb)
        pill_l.line.fill.background()

        pill_r = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(14.0), Inches(0.4), Inches(2.17), Inches(0.64))
        pill_r.name = "!!AppShell_PillR"
        _apply_solid_fill(pill_r, self.accent_rgb)
        pill_r.line.fill.background()

        # Add intra-slide entrance timeline to slide 1
        tline1 = SlideAnimationTimeline(slide_number=1)
        tline1.add_effect(hero_icon, effect_type="fly_in", direction="up", duration_ms=600, trigger="with_previous")
        tline1.compile_to_openxml(s1)

        # -------------------------------------------------------------------
        # SLIDE 2: Convergence & Title Lockup
        # -------------------------------------------------------------------
        s2 = deck.add_slide(transition="morph", duration_ms=2000)
        _set_slide_background(s2, self.dark_bg_hex)

        # Orb 1 contracts into top-left corner
        orb2_tl = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-3.27), Inches(-2.82), Inches(6.18), Inches(6.26))
        orb2_tl.name = "!!Spotify_AmbientOrb"
        _apply_solid_fill(orb2_tl, self.accent_rgb, alpha=0.95)
        orb2_tl.line.fill.background()

        # Secondary orb docks into bottom-right corner
        orb2_br = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.70), Inches(4.55), Inches(6.18), Inches(6.26))
        orb2_br.name = "!!Spotify_AmbientOrb2"
        _apply_solid_fill(orb2_br, self.accent_rgb, alpha=0.95)
        orb2_br.line.fill.background()

        # Hero emblem docks
        hero_icon2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.83), Inches(2.61), Inches(1.67), Inches(1.66))
        hero_icon2.name = "!!Spotify_HeroIcon"
        _apply_solid_fill(hero_icon2, RGBColor(255, 255, 255))
        hero_icon2.line.fill.background()
        tf = hero_icon2.text_frame
        p = tf.paragraphs[0]
        p.text = "●"
        p.font.size = Pt(36)
        p.font.color.rgb = self.accent_rgb
        p.alignment = PP_ALIGN.CENTER

        # Title and subtitle glide in laterally from margin
        t2 = s2.shapes.add_textbox(Inches(7.73), Inches(2.52), Inches(5.0), Inches(1.11))
        t2.name = "!!AppShell_Title"
        tf2 = t2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = self.topic_title
        p2.font.size = Pt(44)
        p2.font.bold = True
        p2.font.color.rgb = self.accent_rgb

        sub2 = s2.shapes.add_textbox(Inches(7.73), Inches(3.63), Inches(5.0), Inches(0.64))
        sub2.name = "!!AppShell_Subtitle"
        tf_s2 = sub2.text_frame
        p_s2 = tf_s2.paragraphs[0]
        p_s2.text = self.topic_subtitle
        p_s2.font.size = Pt(20)
        p_s2.font.bold = True
        p_s2.font.color.rgb = RGBColor(255, 255, 255)

        # Action pills dock into position
        p_dock_l = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.53), Inches(6.59), Inches(2.17), Inches(0.64))
        p_dock_l.name = "!!AppShell_PillL"
        _apply_solid_fill(p_dock_l, self.accent_rgb)
        p_dock_l.line.fill.background()

        p_dock_r = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.79), Inches(0.31), Inches(2.17), Inches(0.64))
        p_dock_r.name = "!!AppShell_PillR"
        _apply_solid_fill(p_dock_r, self.accent_rgb)
        p_dock_r.line.fill.background()

        # Staged sidebar background waiting off-screen left
        sb_wait = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(-3.5), Inches(0.0), Inches(3.29), Inches(7.50))
        sb_wait.name = "!!AppShell_Sidebar"
        _apply_solid_fill(sb_wait, RGBColor.from_string(self.sidebar_bg_hex))
        sb_wait.line.fill.background()

        # -------------------------------------------------------------------
        # SLIDE 3: The Full Interactive Application Shell Unlock
        # -------------------------------------------------------------------
        s3 = deck.add_slide(transition="morph", duration_ms=2000)
        _set_slide_background(s3, "191428")

        self._render_app_shell_framework(s3, active_index=0)

        # Viewport Stage: Hero Welcome headline
        vp_title = s3.shapes.add_textbox(Inches(3.52), Inches(1.71), Inches(7.5), Inches(1.11))
        vp_title.name = "!!AppShell_Title"
        tf_vpt = vp_title.text_frame
        p_vpt = tf_vpt.paragraphs[0]
        p_vpt.text = "WELCOME"
        p_vpt.font.size = Pt(48)
        p_vpt.font.bold = True
        p_vpt.font.color.rgb = self.accent_rgb

        # Narrative paragraph
        desc = s3.shapes.add_textbox(Inches(3.52), Inches(2.83), Inches(9.20), Inches(1.6))
        desc.name = "!!AppShell_ViewportDesc"
        tf_desc = desc.text_frame
        tf_desc.word_wrap = True
        p_desc = tf_desc.paragraphs[0]
        p_desc.text = (
            f"Welcome to {self.topic_title}. Built with a native fluid interactive motion architecture. "
            f"Every section, control, and card transitions seamlessly across slides using native hardware-accelerated Morph."
        )
        p_desc.font.size = Pt(13)
        p_desc.font.color.rgb = self.text_muted

        # Action pill buttons: Play, Follow, ...
        play_btn = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.63), Inches(4.83), Inches(2.17), Inches(0.64))
        play_btn.name = "!!AppShell_PlayBtn"
        _apply_solid_fill(play_btn, self.accent_rgb)
        play_btn.line.fill.background()
        tf_pb = play_btn.text_frame
        p_pb = tf_pb.paragraphs[0]
        p_pb.text = "Play"
        p_pb.font.size = Pt(16)
        p_pb.font.bold = True
        p_pb.font.color.rgb = RGBColor(0, 0, 0)
        p_pb.alignment = PP_ALIGN.CENTER

        follow_btn = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.24), Inches(4.83), Inches(2.17), Inches(0.64))
        follow_btn.name = "!!AppShell_FollowBtn"
        follow_btn.fill.background()
        _apply_hairline_border(follow_btn, RGBColor(255, 255, 255), width_pt=1.0)
        tf_fb = follow_btn.text_frame
        p_fb = tf_fb.paragraphs[0]
        p_fb.text = "Follow"
        p_fb.font.size = Pt(16)
        p_fb.font.bold = True
        p_fb.font.color.rgb = RGBColor(255, 255, 255)
        p_fb.alignment = PP_ALIGN.CENTER

        more_btn = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.86), Inches(4.95), Inches(0.8), Inches(0.45))
        more_btn.name = "!!AppShell_MoreBtn"
        more_btn.fill.background()
        tf_mb = more_btn.text_frame
        p_mb = tf_mb.paragraphs[0]
        p_mb.text = "● ● ●"
        p_mb.font.size = Pt(11)
        p_mb.font.color.rgb = self.text_muted
        p_mb.alignment = PP_ALIGN.CENTER

        if num_slides >= 4:
            # ---------------------------------------------------------------
            # SLIDE 4: In-Frame Viewport Stream (About Us / Multi-Card)
            # ---------------------------------------------------------------
            s4 = deck.add_slide(transition="morph", duration_ms=2000)
            _set_slide_background(s4, "1B162E")
            self._render_app_shell_framework(s4, active_index=1)

            # Viewport Header with album / feature media card
            feat_card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.58), Inches(1.27), Inches(1.54), Inches(1.54))
            feat_card.name = "!!AppShell_FeatureCard"
            _apply_solid_fill(feat_card, self.accent_rgb, alpha=0.3)
            _apply_hairline_border(feat_card, self.accent_rgb, width_pt=1.0)
            tf_fc = feat_card.text_frame
            p_fc = tf_fc.paragraphs[0]
            p_fc.text = "★"
            p_fc.font.size = Pt(36)
            p_fc.font.color.rgb = self.accent_rgb
            p_fc.alignment = PP_ALIGN.CENTER

            vp_sub = s4.shapes.add_textbox(Inches(5.31), Inches(1.73), Inches(3.5), Inches(0.35))
            vp_sub.name = "!!AppShell_VpSub"
            vp_sub.text_frame.text = "Introduction"
            vp_sub.text_frame.paragraphs[0].font.size = Pt(12)
            vp_sub.text_frame.paragraphs[0].font.color.rgb = self.text_muted

            vp_title4 = s4.shapes.add_textbox(Inches(5.31), Inches(2.08), Inches(5.0), Inches(0.84))
            vp_title4.name = "!!AppShell_Title"
            tf4 = vp_title4.text_frame
            p4 = tf4.paragraphs[0]
            p4.text = "About Us"
            p4.font.size = Pt(36)
            p4.font.bold = True
            p4.font.color.rgb = self.accent_rgb

            # 2-column description cards inside the viewport
            c1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.52), Inches(3.40), Inches(4.35), Inches(3.40))
            c1.name = "!!AppShell_ColCard1"
            _apply_solid_fill(c1, RGBColor(24, 24, 24), alpha=0.85)
            _apply_hairline_border(c1, RGBColor(45, 45, 45), width_pt=0.75)
            tf_c1 = c1.text_frame
            tf_c1.word_wrap = True
            p_c1_h = tf_c1.paragraphs[0]
            p_c1_h.text = "Core Mission"
            p_c1_h.font.size = Pt(16)
            p_c1_h.font.bold = True
            p_c1_h.font.color.rgb = RGBColor(255, 255, 255)
            p_c1_b = tf_c1.add_paragraph()
            p_c1_b.text = (
                "Empowering human creativity and problem-solving through clean, high-performance visual motion systems. "
                "Every element remains editable and dynamic."
            )
            p_c1_b.font.size = Pt(11)
            p_c1_b.font.color.rgb = self.text_muted

            c2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.15), Inches(3.40), Inches(4.55), Inches(3.40))
            c2.name = "!!AppShell_ColCard2"
            _apply_solid_fill(c2, RGBColor(24, 24, 24), alpha=0.85)
            _apply_hairline_border(c2, RGBColor(45, 45, 45), width_pt=0.75)
            tf_c2 = c2.text_frame
            tf_c2.word_wrap = True
            p_c2_h = tf_c2.paragraphs[0]
            p_c2_h.text = "Shared Values"
            p_c2_h.font.size = Pt(16)
            p_c2_h.font.bold = True
            p_c2_h.font.color.rgb = RGBColor(255, 255, 255)
            p_c2_b = tf_c2.add_paragraph()
            p_c2_b.text = (
                "Collaboration, precision, and continuous iteration. "
                "Delivering presentations that behave like full desktop software applications."
            )
            p_c2_b.font.size = Pt(11)
            p_c2_b.font.color.rgb = self.text_muted

        if num_slides >= 5:
            # ---------------------------------------------------------------
            # SLIDE 5: Team Roster / Feature Grid Stream
            # ---------------------------------------------------------------
            s5 = deck.add_slide(transition="morph", duration_ms=2000)
            _set_slide_background(s5, "181424")
            self._render_app_shell_framework(s5, active_index=2)

            t5 = s5.shapes.add_textbox(Inches(3.56), Inches(0.70), Inches(6.5), Inches(0.8))
            t5.name = "!!AppShell_Title"
            tf5 = t5.text_frame
            p5 = tf5.paragraphs[0]
            p5.text = "Our Team Members"
            p5.font.size = Pt(32)
            p5.font.bold = True
            p5.font.color.rgb = self.accent_rgb

            # 4 horizontal member cards
            card_w = 2.15
            gap = 0.22
            base_l = 3.56
            members = [
                ("Duy Long", "Team Leader", "01"),
                ("Thành Nam", "Vice Leader", "02"),
                ("Trọng Đức", "Content", "03"),
                ("Huy Hoàng", "Design", "04"),
            ]
            for m_idx, (m_name, m_role, m_num) in enumerate(members):
                cl = base_l + m_idx * (card_w + gap)
                mc = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cl), Inches(1.75), Inches(card_w), Inches(4.80))
                mc.name = f"!!AppShell_MemberCard_{m_idx}"
                _apply_solid_fill(mc, RGBColor(24, 24, 24), alpha=0.90)
                _apply_hairline_border(mc, RGBColor(50, 50, 50), width_pt=0.75)

                # Avatar node
                av = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cl + 0.42), Inches(2.15), Inches(1.30), Inches(1.30))
                av.name = f"!!AppShell_MemberAv_{m_idx}"
                _apply_solid_fill(av, self.accent_rgb, alpha=0.3)
                _apply_hairline_border(av, self.accent_rgb, width_pt=1.0)
                tf_av = av.text_frame
                tf_av.paragraphs[0].text = m_num
                tf_av.paragraphs[0].font.size = Pt(20)
                tf_av.paragraphs[0].font.bold = True
                tf_av.paragraphs[0].font.color.rgb = self.accent_rgb
                tf_av.paragraphs[0].alignment = PP_ALIGN.CENTER

                # Name & role
                mt = s5.shapes.add_textbox(Inches(cl + 0.1), Inches(3.70), Inches(card_w - 0.2), Inches(1.2))
                mt.name = f"!!AppShell_MemberTxt_{m_idx}"
                tf_mt = mt.text_frame
                tf_mt.word_wrap = True
                p_mn = tf_mt.paragraphs[0]
                p_mn.text = m_name
                p_mn.font.size = Pt(15)
                p_mn.font.bold = True
                p_mn.font.color.rgb = RGBColor(255, 255, 255)
                p_mn.alignment = PP_ALIGN.CENTER
                p_mr = tf_mt.add_paragraph()
                p_mr.text = m_role
                p_mr.font.size = Pt(11)
                p_mr.font.color.rgb = self.text_muted
                p_mr.alignment = PP_ALIGN.CENTER

        deck.save(output_path)
        return output_path

    def _render_app_shell_framework(self, slide, active_index: int = 0) -> None:
        """
        Renders the persistent Application Shell:
        - Left Sidebar Navigation Rail (w=3.29, h=7.50)
        - macOS window dots
        - Navigation items with active accent highlight
        - Top bar with back/forward controls, upgrade pill, and user profile avatar
        """
        # 1. Left Sidebar background
        sb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.0), Inches(0.0), Inches(3.29), Inches(7.50))
        sb.name = "!!AppShell_Sidebar"
        _apply_solid_fill(sb, RGBColor.from_string(self.sidebar_bg_hex))
        sb.line.fill.background()

        # 2. Window Control Dots (macOS style)
        dot_colors = [RGBColor(255, 95, 86), RGBColor(255, 189, 46), RGBColor(39, 201, 63)]
        for d_i, d_x in enumerate([0.27, 0.40, 0.53]):
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(d_x), Inches(0.24), Inches(0.08), Inches(0.08))
            dot.name = f"!!AppShell_WinDot_{d_i}"
            _apply_solid_fill(dot, dot_colors[d_i])
            dot.line.fill.background()

        # 3. Sidebar Navigation Items
        base_y = 0.92
        gap_y = 0.50
        for idx, item_label in enumerate(self.nav_items[:8]):
            y_pos = base_y + idx * gap_y
            is_active = (idx == active_index)

            # Icon circle / indicator
            ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.44), Inches(y_pos + 0.04), Inches(0.20), Inches(0.20))
            ic.name = f"!!AppShell_NavIcon_{idx}"
            if is_active:
                _apply_solid_fill(ic, self.accent_rgb)
            else:
                _apply_solid_fill(ic, RGBColor(90, 90, 90))
            ic.line.fill.background()

            # Item label
            tb = slide.shapes.add_textbox(Inches(0.95), Inches(y_pos), Inches(2.1), Inches(0.35))
            tb.name = f"!!AppShell_NavTxt_{idx}"
            tf = tb.text_frame
            p = tf.paragraphs[0]
            p.text = item_label
            p.font.size = Pt(12)
            p.font.bold = is_active
            p.font.color.rgb = self.accent_rgb if is_active else self.text_muted

        # Separator line
        sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.37), Inches(3.50), Inches(2.55), Inches(0.01))
        sep.name = "!!AppShell_SepLine"
        _apply_solid_fill(sep, RGBColor(60, 60, 60))
        sep.line.fill.background()

        # 4. Top Navigation Bar
        # Back/Forward history buttons
        for b_idx, b_x in enumerate([3.56, 4.05]):
            btn = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(b_x), Inches(0.32), Inches(0.34), Inches(0.34))
            btn.name = f"!!AppShell_HistoryBtn_{b_idx}"
            _apply_solid_fill(btn, RGBColor(20, 20, 25))
            btn.line.fill.background()
            tf_b = btn.text_frame
            p_b = tf_b.paragraphs[0]
            p_b.text = "<" if b_idx == 0 else ">"
            p_b.font.size = Pt(11)
            p_b.font.bold = True
            p_b.font.color.rgb = self.text_muted
            p_b.alignment = PP_ALIGN.CENTER

        # Upgrade / Action Pill
        upg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.96), Inches(0.32), Inches(1.15), Inches(0.34))
        upg.name = "!!AppShell_TopUpgrade"
        upg.fill.background()
        _apply_hairline_border(upg, RGBColor(120, 120, 120), width_pt=0.75)
        tf_u = upg.text_frame
        p_u = tf_u.paragraphs[0]
        p_u.text = "Upgrade"
        p_u.font.size = Pt(11)
        p_u.font.color.rgb = RGBColor(255, 255, 255)
        p_u.alignment = PP_ALIGN.CENTER

        # User Profile Pill
        prof = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.40), Inches(0.32), Inches(1.55), Inches(0.34))
        prof.name = "!!AppShell_TopProfile"
        _apply_solid_fill(prof, RGBColor(20, 20, 25))
        prof.line.fill.background()

        av_top = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(11.45), Inches(0.35), Inches(0.28), Inches(0.28))
        av_top.name = "!!AppShell_TopAv"
        _apply_solid_fill(av_top, self.accent_rgb)
        av_top.line.fill.background()

        prof_txt = slide.shapes.add_textbox(Inches(11.75), Inches(0.34), Inches(1.15), Inches(0.30))
        prof_txt.name = "!!AppShell_TopUserTxt"
        p_pt = prof_txt.text_frame.paragraphs[0]
        p_pt.text = f"{self.author_name} ▼"
        p_pt.font.size = Pt(11)
        p_pt.font.color.rgb = RGBColor(255, 255, 255)


# ===========================================================================
# 2. EARTH AND SPACE 21-SLIDE CINEMATIC SYSTEM (from EARTH AND SPACE.pptx)
# ===========================================================================

class EarthSpaceMotionBuilder:
    """
    Constructs high-end cinematic science/tech/explainer presentations based on
    the motion mechanics of EARTH AND SPACE.pptx:
    - Cosmic establishing shot & planetary traverse
    - Viewfinder / Targeting Reticle Zoom Sequence (corner brackets framing subject)
    - Expansion into framed culture/topic highlight cards
    - Centered Focal Mystery / Transition Node ("BUT SOMETHING IS MISSING" macro zoom)
    - 4-Quadrant Rotating Segmented Radar Wheel Navigator (0 -> 90 -> 180 -> 270 deg)
    - Orbiting Hero Model with Glowing Radial HUD Dial
    """

    def __init__(
        self,
        topic_title: str = "EARTH AND SPACE",
        topic_subtitle: str = "A CINEMATIC EXPLORATION",
        theme_accent_hex: str = "00B0F0",  # Cyan accent
        stars_bg_hex: str = "05050A",      # Deep cosmic black
    ):
        self.topic_title = topic_title
        self.topic_subtitle = topic_subtitle
        self.accent_hex = theme_accent_hex
        self.accent_rgb = RGBColor.from_string(theme_accent_hex)
        self.stars_bg_hex = stars_bg_hex

    def build_deck(self, output_path: str, num_slides: int = 8) -> str:
        deck = MotionDeck(theme=Themes.COSMOS_DEEP)
        num_slides = max(4, min(num_slides, 10))

        # -------------------------------------------------------------------
        # SLIDE 1: Cosmic Establishing Shot (Globe entering from bottom)
        # -------------------------------------------------------------------
        s1 = deck.add_slide(transition="none")
        _set_slide_background(s1, self.stars_bg_hex)

        # Starry dust nodes
        self._add_cosmic_particles(s1)

        # Hero celestial body / globe entering from bottom
        globe = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.30), Inches(5.80), Inches(8.73), Inches(8.34))
        globe.name = "!!Space_Globe"
        _apply_solid_fill(globe, RGBColor(2, 132, 199))
        _apply_hairline_border(globe, RGBColor(56, 189, 248), width_pt=2.0)

        t1 = s1.shapes.add_textbox(Inches(2.5), Inches(1.50), Inches(8.33), Inches(1.5))
        t1.name = "!!Space_HeroTitle"
        tf1 = t1.text_frame
        tf1.word_wrap = True
        p1 = tf1.paragraphs[0]
        p1.text = "WELCOME TO OUR\nPRESENTATION"
        p1.font.size = Pt(44)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(255, 255, 255)
        p1.alignment = PP_ALIGN.CENTER

        # -------------------------------------------------------------------
        # SLIDE 2: Camera Pan & Moon Orbit
        # -------------------------------------------------------------------
        s2 = deck.add_slide(transition="morph", duration_ms=2000)
        _set_slide_background(s2, self.stars_bg_hex)
        self._add_cosmic_particles(s2)

        # Globe shifts to top-left flank
        globe2 = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-3.67), Inches(-1.03), Inches(16.21), Inches(8.78))
        globe2.name = "!!Space_Globe"
        _apply_solid_fill(globe2, RGBColor(2, 132, 199))
        _apply_hairline_border(globe2, RGBColor(56, 189, 248), width_pt=2.0)

        # Moon node
        moon = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.50), Inches(2.60), Inches(1.20), Inches(1.20))
        moon.name = "!!Space_Moon"
        _apply_solid_fill(moon, RGBColor(220, 225, 235))
        moon.line.fill.background()

        t2 = s2.shapes.add_textbox(Inches(3.74), Inches(0.80), Inches(8.66), Inches(1.5))
        t2.name = "!!Space_HeroTitle"
        tf2 = t2.text_frame
        p2 = tf2.paragraphs[0]
        p2.text = f"THIS IS {self.topic_title}"
        p2.font.size = Pt(36)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(255, 255, 255)
        p2.alignment = PP_ALIGN.CENTER

        # Staged targeting reticle corner brackets (waiting off-screen)
        self._render_viewfinder_reticle(s2, center_x=17.0, center_y=12.0, size_w=2.0, size_h=2.0)

        # -------------------------------------------------------------------
        # SLIDE 3: Targeting Reticle / Viewfinder Zoom
        # -------------------------------------------------------------------
        s3 = deck.add_slide(transition="morph", duration_ms=2000)
        _set_slide_background(s3, self.stars_bg_hex)
        self._add_cosmic_particles(s3)

        # Globe zooms in closer
        globe3 = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-1.87), Inches(-9.19), Inches(44.88), Inches(20.97))
        globe3.name = "!!Space_Globe"
        _apply_solid_fill(globe3, RGBColor(2, 132, 199))
        globe3.line.fill.background()

        # Viewfinder reticle locks into center-right coordinates [ ]
        self._render_viewfinder_reticle(s3, center_x=9.50, center_y=3.50, size_w=2.40, size_h=2.40)

        t3 = s3.shapes.add_textbox(Inches(1.5), Inches(0.80), Inches(8.0), Inches(1.0))
        t3.name = "!!Space_HeroTitle"
        tf3 = t3.text_frame
        p3 = tf3.paragraphs[0]
        p3.text = "TARGET COORDINATES LOCKED"
        p3.font.size = Pt(28)
        p3.font.bold = True
        p3.font.color.rgb = self.accent_rgb

        # -------------------------------------------------------------------
        # SLIDE 4: Viewfinder Expansion to 3 Framed Highlight Cards
        # -------------------------------------------------------------------
        s4 = deck.add_slide(transition="morph", duration_ms=2000)
        _set_slide_background(s4, self.stars_bg_hex)
        self._add_cosmic_particles(s4)

        # 3 Framed cards (e.g. Ha Long Bay, Street Food, Transport / Technology)
        card_data = [
            ("LANDMARK", "Natural & Cultural Wonders", 1.36, 1.20, 3.40, 4.80),
            ("INNOVATION", "Modern Infrastructure", 5.00, 1.20, 3.40, 4.80),
            ("COMMUNITY", "Connected Populations", 8.64, 1.20, 3.40, 4.80),
        ]
        for c_idx, (c_label, c_sub, c_l, c_t, c_w, c_h) in enumerate(card_data):
            # Outer reticle frame
            self._render_viewfinder_reticle(s4, center_x=c_l + c_w / 2, center_y=c_t + c_h / 2, size_w=c_w, size_h=c_h, name_suffix=f"_{c_idx}")

            card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(c_l + 0.15), Inches(c_t + 0.15), Inches(c_w - 0.3), Inches(c_h - 0.3))
            card.name = f"!!Space_Card_{c_idx}"
            _apply_solid_fill(card, RGBColor(20, 25, 35), alpha=0.88)
            _apply_hairline_border(card, self.accent_rgb, width_pt=1.0)
            tf_c = card.text_frame
            tf_c.word_wrap = True
            p_ch = tf_c.paragraphs[0]
            p_ch.text = c_label
            p_ch.font.size = Pt(16)
            p_ch.font.bold = True
            p_ch.font.color.rgb = self.accent_rgb
            p_cb = tf_c.add_paragraph()
            p_cb.text = c_sub
            p_cb.font.size = Pt(12)
            p_cb.font.color.rgb = RGBColor(220, 220, 220)

        # -------------------------------------------------------------------
        # SLIDE 5: Focal Mystery / Transition Node ("BUT SOMETHING IS MISSING")
        # -------------------------------------------------------------------
        s5 = deck.add_slide(transition="morph", duration_ms=2000)
        _set_slide_background(s5, self.stars_bg_hex)
        self._add_cosmic_particles(s5)

        focal_top_txt = s5.shapes.add_textbox(Inches(3.0), Inches(0.80), Inches(7.33), Inches(1.0))
        focal_top_txt.name = "!!Space_FocalTop"
        p_ft = focal_top_txt.text_frame.paragraphs[0]
        p_ft.text = "BUT SOMETHING"
        p_ft.font.size = Pt(40)
        p_ft.font.bold = True
        p_ft.font.color.rgb = RGBColor(255, 255, 255)
        p_ft.alignment = PP_ALIGN.CENTER

        # Central focal dark disk with question mark
        focal_orb = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4.40), Inches(1.60), Inches(4.50), Inches(4.50))
        focal_orb.name = "!!Space_FocalNode"
        _apply_solid_fill(focal_orb, RGBColor(5, 5, 10))
        _apply_hairline_border(focal_orb, RGBColor(120, 120, 120), width_pt=1.5)
        tf_fo = focal_orb.text_frame
        p_fo = tf_fo.paragraphs[0]
        p_fo.text = "?"
        p_fo.font.size = Pt(88)
        p_fo.font.bold = True
        p_fo.font.color.rgb = RGBColor(255, 255, 255)
        p_fo.alignment = PP_ALIGN.CENTER

        focal_bot_txt = s5.shapes.add_textbox(Inches(3.0), Inches(6.20), Inches(7.33), Inches(1.0))
        focal_bot_txt.name = "!!Space_FocalBot"
        p_fb = focal_bot_txt.text_frame.paragraphs[0]
        p_fb.text = "IS MISSING"
        p_fb.font.size = Pt(40)
        p_fb.font.bold = True
        p_fb.font.color.rgb = RGBColor(255, 255, 255)
        p_fb.alignment = PP_ALIGN.CENTER

        # -------------------------------------------------------------------
        # SLIDES 6-8: 4-Quadrant Segmented Radar Wheel Navigator
        # -------------------------------------------------------------------
        if num_slides >= 6:
            quad_rotations = [0.0, 90.0, 180.0, 270.0]
            pillars = [
                ("PILLAR 01: DISCOVERY", "Exploring unknown planetary bodies and deep telemetry signals."),
                ("PILLAR 02: ANALYSIS", "Spectral classification and atmospheric composition sampling."),
                ("PILLAR 03: EXPANSION", "Long-range robotic probes establishing lunar relays."),
            ]
            for q_idx in range(min(len(pillars), num_slides - 5)):
                s_quad = deck.add_slide(transition="morph", duration_ms=2000)
                _set_slide_background(s_quad, self.stars_bg_hex)
                self._add_cosmic_particles(s_quad)

                # Rotating 4-Quadrant wheel on the right
                rot_deg = quad_rotations[q_idx]
                self._render_quadrant_radar_wheel(s_quad, center_x=10.0, center_y=3.75, radius=2.6, rot_deg=rot_deg, active_quad=q_idx)

                # Pillar title & description card on the left
                p_title, p_desc = pillars[q_idx]
                p_card = s_quad.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(2.0), Inches(6.0), Inches(3.5))
                p_card.name = f"!!Space_PillarCard_{q_idx}"
                _apply_solid_fill(p_card, RGBColor(20, 22, 32), alpha=0.90)
                _apply_hairline_border(p_card, self.accent_rgb, width_pt=1.2)
                tf_p = p_card.text_frame
                tf_p.word_wrap = True
                p_head = tf_p.paragraphs[0]
                p_head.text = p_title
                p_head.font.size = Pt(26)
                p_head.font.bold = True
                p_head.font.color.rgb = self.accent_rgb
                p_body = tf_p.add_paragraph()
                p_body.text = p_desc
                p_body.font.size = Pt(14)
                p_body.font.color.rgb = RGBColor(220, 220, 220)

        deck.save(output_path)
        return output_path

    def _add_cosmic_particles(self, slide) -> None:
        """Adds subtle star dust points."""
        coords = [(1.2, 0.8), (3.4, 2.1), (7.8, 0.5), (11.2, 1.4), (2.0, 6.2), (8.5, 6.8), (12.4, 5.5)]
        for i, (x, y) in enumerate(coords):
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.04), Inches(0.04))
            dot.name = f"!!Space_Star_{i}"
            _apply_solid_fill(dot, RGBColor(255, 255, 255), alpha=0.7)
            dot.line.fill.background()

    def _render_viewfinder_reticle(
        self, slide, center_x: float, center_y: float, size_w: float, size_h: float, name_suffix: str = ""
    ) -> None:
        """
        Renders 4 L-shaped corner brackets [ ] forming a cinematic targeting reticle.
        """
        half_w = size_w / 2.0
        half_h = size_h / 2.0
        arm_len = min(size_w, size_h) * 0.28
        thickness = 0.05

        corners = [
            ("TL", center_x - half_w, center_y - half_h, 0),
            ("TR", center_x + half_w - arm_len, center_y - half_h, 90),
            ("BR", center_x + half_w - arm_len, center_y + half_h - arm_len, 180),
            ("BL", center_x - half_w, center_y + half_h - arm_len, 270),
        ]
        for c_tag, x, y, rot in corners:
            bracket = slide.shapes.add_shape(MSO_SHAPE.CORNER, Inches(x), Inches(y), Inches(arm_len), Inches(arm_len))
            bracket.name = f"!!Reticle_{c_tag}{name_suffix}"
            _apply_solid_fill(bracket, RGBColor(220, 220, 230))
            bracket.line.fill.background()
            _set_shape_rotation(bracket, rot)

    def _render_quadrant_radar_wheel(
        self, slide, center_x: float, center_y: float, radius: float, rot_deg: float, active_quad: int = 0
    ) -> None:
        """
        Renders a 4-quadrant segmented radar wheel rotating per slide.
        """
        angles = [0, 90, 180, 270]
        for idx, base_ang in enumerate(angles):
            seg = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, Inches(center_x - radius), Inches(center_y - radius), Inches(radius * 2), Inches(radius * 2))
            seg.name = f"!!Radar_Quadrant_{idx}"
            is_active = (idx == active_quad)
            fill_clr = self.accent_rgb if is_active else RGBColor(50, 55, 70)
            _apply_solid_fill(seg, fill_clr, alpha=0.95 if is_active else 0.50)
            _apply_hairline_border(seg, RGBColor(255, 255, 255), width_pt=1.0)
            _set_shape_rotation(seg, base_ang + rot_deg)


# ===========================================================================
# 3. UNIVERSITY CONCENTRIC BLOOM & 8-PETAL MORPH BUILDER (from PROMPT.pdf)
# ===========================================================================

class UniversityBloomBuilder:
    """
    Constructs the master 4-slide university morph presentation based on PROMPT.pdf:
    - Slide 1: Centered concentric circles & diamond corner accents
    - Slide 2: Morph expansion into massive concentric rings with cover title and instructor badges
    - Slide 3: 8 curved bezier petals blooming radially around central hub with 8 icons
    - Slide 4: Morph into the 8-item two-column Table of Contents (Mục Lục)
    """

    def __init__(
        self,
        presentation_title: str = "TIÊU ĐỀ BÀI THUYẾT TRÌNH",
        subtitle: str = "KHẢO SÁT Ý KIẾN VÀ ĐỀ XUẤT",
        student_name: str = "Nguyễn Văn A",
        student_id: str = "112233",
        advisor_name: str = "TS. Nguyễn Văn B",
        theme_blue_hex: str = "03A9F4",
    ):
        self.title = presentation_title
        self.subtitle = subtitle
        self.student_name = student_name
        self.student_id = student_id
        self.advisor_name = advisor_name
        self.accent_hex = theme_blue_hex
        self.accent_rgb = RGBColor.from_string(theme_blue_hex)

    def build_deck(self, output_path: str) -> str:
        deck = MotionDeck(theme=Themes.LIGHT_ACADEMIC_CYAN)

        # -------------------------------------------------------------------
        # SLIDE 1: Centered Concentric Circles & Diamond Corners
        # -------------------------------------------------------------------
        s1 = deck.add_slide(transition="none")

        # Concentric center circles
        c1 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.23), Inches(3.31), Inches(0.87), Inches(0.87))
        c1.name = "!!Oval 6"
        _apply_hairline_border(c1, self.accent_rgb, width_pt=1.5)
        c1.fill.background()

        c2 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.11), Inches(3.19), Inches(1.11), Inches(1.11))
        c2.name = "!!Oval 1"
        _apply_solid_fill(c2, self.accent_rgb, alpha=0.3)
        _apply_hairline_border(c2, self.accent_rgb, width_pt=1.0)

        # Diamond corner accents
        d1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.78), Inches(3.35), Inches(0.91), Inches(0.91))
        d1.name = "!!Rectangle: Rounded Corners 10"
        _apply_solid_fill(d1, RGBColor(220, 220, 225))
        d1.line.fill.background()
        _set_shape_rotation(d1, 135)

        d2 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.64), Inches(3.35), Inches(0.91), Inches(0.91))
        d2.name = "!!Rectangle: Rounded Corners 11"
        _apply_solid_fill(d2, RGBColor(220, 220, 225))
        d2.line.fill.background()
        _set_shape_rotation(d2, 135)

        # -------------------------------------------------------------------
        # SLIDE 2: Expanding Concentric Bloom & Cover Title
        # -------------------------------------------------------------------
        s2 = deck.add_slide(transition="morph", duration_ms=1600)

        # Big expanded rings
        r1 = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.17), Inches(-1.73), Inches(10.97), Inches(10.97))
        r1.name = "!!Oval 6"
        _apply_hairline_border(r1, self.accent_rgb, width_pt=2.0)
        r1.fill.background()

        r2 = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.89), Inches(0.97), Inches(5.54), Inches(5.54))
        r2.name = "!!Oval 1"
        _apply_solid_fill(r2, self.accent_rgb, alpha=0.25)
        _apply_hairline_border(r2, self.accent_rgb, width_pt=1.5)

        # Title card
        tc = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.90), Inches(5.50), Inches(11.53), Inches(1.80))
        tc.name = "!!TitleCard"
        _apply_solid_fill(tc, self.accent_rgb)
        tc.line.fill.background()
        tf_tc = tc.text_frame
        tf_tc.word_wrap = True
        p_tc = tf_tc.paragraphs[0]
        p_tc.text = self.title
        p_tc.font.size = Pt(28)
        p_tc.font.bold = True
        p_tc.font.color.rgb = RGBColor(255, 255, 255)
        p_tc.alignment = PP_ALIGN.CENTER

        # Student badge
        sb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.67), Inches(6.55), Inches(3.98), Inches(0.70))
        sb.name = "!!StudentBadge"
        _apply_solid_fill(sb, RGBColor(255, 255, 255))
        sb.line.fill.background()
        tf_sb = sb.text_frame
        p_sb = tf_sb.paragraphs[0]
        p_sb.text = f"{self.student_name} — MSSV: {self.student_id}"
        p_sb.font.size = Pt(13)
        p_sb.font.bold = True
        p_sb.font.color.rgb = self.accent_rgb
        p_sb.alignment = PP_ALIGN.CENTER

        # -------------------------------------------------------------------
        # SLIDE 3: 8 Curved Petals Blooming Radially Around Center Hub
        # -------------------------------------------------------------------
        s3 = deck.add_slide(transition="morph", duration_ms=1600)

        # Title card morphs to top
        tc3 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.90), Inches(0.55), Inches(11.53), Inches(6.40))
        tc3.name = "!!TitleCard"
        _apply_solid_fill(tc3, self.accent_rgb)
        tc3.line.fill.background()

        # Center circular white hub
        hub = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.81), Inches(1.10), Inches(1.70), Inches(1.70))
        hub.name = "!!BloomHub"
        _apply_solid_fill(hub, RGBColor(255, 255, 255))
        hub.line.fill.background()

        # 8 Blooming petals around the hub
        angles = [0, 45, 90, 135, 180, 225, 270, 315]
        for p_idx, ang in enumerate(angles):
            rad = math.radians(ang)
            dist = 1.35
            px = 6.66 + dist * math.cos(rad) - 0.40
            py = 1.95 + dist * math.sin(rad) - 0.40
            petal = s3.shapes.add_shape(MSO_SHAPE.OVAL, Inches(px), Inches(py), Inches(0.80), Inches(0.80))
            petal.name = f"!!Petal_{p_idx}"
            _apply_solid_fill(petal, RGBColor(255, 255, 255), alpha=0.9)
            _apply_hairline_border(petal, self.accent_rgb, width_pt=1.0)
            tf_pt = petal.text_frame
            tf_pt.paragraphs[0].text = str(p_idx + 1)
            tf_pt.paragraphs[0].font.size = Pt(14)
            tf_pt.paragraphs[0].font.bold = True
            tf_pt.paragraphs[0].font.color.rgb = self.accent_rgb
            tf_pt.paragraphs[0].alignment = PP_ALIGN.CENTER

        # -------------------------------------------------------------------
        # SLIDE 4: Morph into 8-Section Table of Contents (Mục Lục)
        # -------------------------------------------------------------------
        s4 = deck.add_slide(transition="morph", duration_ms=1600)
        _set_slide_background(s4, "0288D1")

        # Top Title "MỤC LỤC"
        toc_title = s4.shapes.add_textbox(Inches(4.5), Inches(0.30), Inches(4.33), Inches(0.8))
        toc_title.name = "!!TocTitle"
        p_tt = toc_title.text_frame.paragraphs[0]
        p_tt.text = "MỤC LỤC"
        p_tt.font.size = Pt(36)
        p_tt.font.bold = True
        p_tt.font.color.rgb = RGBColor(255, 255, 255)
        p_tt.alignment = PP_ALIGN.CENTER

        # 8 Sections in 2 Columns (1-4 on left, 5-8 on right)
        section_titles = [
            "Tổng quan khảo sát", "Đặc điểm mẫu", "Thực trạng vấn đề", "Nguyên nhân cốt lõi",
            "Phỏng vấn chuyên gia", "Đánh giá mức độ", "Giải pháp đề xuất", "Kết luận & Hành động",
        ]
        col_x = [1.0, 7.0]
        base_y = 1.30
        gap_y = 1.35
        for s_idx, s_name in enumerate(section_titles):
            c_col = 0 if s_idx < 4 else 1
            row = s_idx % 4
            x = col_x[c_col]
            y = base_y + row * gap_y

            # Number badge (morphed from !!Petal_X)
            p_badge = s4.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.90), Inches(0.90))
            p_badge.name = f"!!Petal_{s_idx}"
            _apply_solid_fill(p_badge, RGBColor(255, 255, 255))
            p_badge.line.fill.background()
            tf_b = p_badge.text_frame
            tf_b.paragraphs[0].text = f"0{s_idx + 1}"
            tf_b.paragraphs[0].font.size = Pt(16)
            tf_b.paragraphs[0].font.bold = True
            tf_b.paragraphs[0].font.color.rgb = self.accent_rgb
            tf_b.paragraphs[0].alignment = PP_ALIGN.CENTER

            # Text content card
            card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 1.10), Inches(y), Inches(4.20), Inches(0.90))
            card.name = f"!!TocCard_{s_idx}"
            _apply_solid_fill(card, RGBColor(255, 255, 255), alpha=0.15)
            _apply_hairline_border(card, RGBColor(255, 255, 255), width_pt=0.75)
            tf_c = card.text_frame
            tf_c.word_wrap = True
            p_c = tf_c.paragraphs[0]
            p_c.text = s_name
            p_c.font.size = Pt(15)
            p_c.font.bold = True
            p_c.font.color.rgb = RGBColor(255, 255, 255)

        deck.save(output_path)
        return output_path
