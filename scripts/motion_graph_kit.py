"""
motion_graph_kit.py (v4.0 - 3D Macro Dolly, Spotify UI Shell & Thematic Motion Edition)
========================================================================================
Production-grade Python engine for designing broadcast-grade Motion Graphics presentations:
- Modern Bright/Light Themes as Primary Defaults (Clean White, Academic Cyan, Eco Mint, Editorial Ivory)
- Spotify-Style UI Shell & Radial Morph Unlock (decoded from MAC-LENIN slides 1-4)
- 3D Macro Dolly Zoom & Orbital Pan (decoded from Earth & Space slides 1-21)
- Dynamic Thematic Traveling Entities (Space, Music, Tech, Nature, Business, Culture)
- Hardware-Accelerated 60 FPS Morph & Focal Spotlight Transitions
- Automated 60 FPS MP4 Video Export via PowerPoint Engine
"""

import os
import sys
import time
from typing import Optional, Tuple, List, Union
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml


# ==========================================
# 1. CURATED VIBRANT PALETTES ("GU" THẨM MỸ)
# ==========================================

class Themes:
    """
    Curated color systems designed for maximum legibility and elegance.
    Defaults to Bright, Crisp, Modern Light palettes (academic, corporate, research),
    with selective Dark Obsidian options for late-night / deep tech keynotes.
    """
    
    # --- 1. LIGHT ACADEMIC CYAN (MẶC ĐỊNH SỐ 1 - Sáng sủa, Hiện đại, Đại học, Nghiên cứu) ---
    LIGHT_ACADEMIC_CYAN = {
        "name": "Light Academic Cyan",
        "is_dark": False,
        "bg_hex": "FFFFFF",                   # Pure crisp snow white
        "card_bg_hex": "FFFFFF",              # Frosted white acrylic
        "card_alpha": "92000",                # 92% opacity
        "border_hex": "0284C7",               # Vibrant sky cyan (#0284C7)
        "border_alpha": "55000",
        "accent": RGBColor(2, 132, 199),      # #0284C7 Sky Cyan
        "accent_secondary": RGBColor(37, 99, 235), # #2563EB Royal Blue
        "accent_tertiary": RGBColor(16, 185, 129), # #10B981 Emerald
        "text_primary": RGBColor(15, 23, 42), # #0F172A Deep Slate (Super sharp)
        "text_secondary": RGBColor(71, 85, 105), # #475569 Slate Subtitle
        "text_muted": RGBColor(148, 163, 184), # #94A3B8 Subtle Muted
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI"
    }

    # --- 2. LIGHT MODERN STUDIO (Doanh nghiệp, Startup, Công nghệ thanh lịch) ---
    LIGHT_MODERN_STUDIO = {
        "name": "Light Modern Studio",
        "is_dark": False,
        "bg_hex": "F8FAFC",                   # Soft studio ice white
        "card_bg_hex": "FFFFFF",              # Pure white card surface
        "card_alpha": "90000",
        "border_hex": "4F46E5",               # Indigo electric glow (#4F46E5)
        "border_alpha": "50000",
        "accent": RGBColor(79, 70, 229),      # #4F46E5 Indigo
        "accent_secondary": RGBColor(16, 185, 129), # #10B981 Emerald Mint
        "accent_tertiary": RGBColor(245, 158, 11),  # #F59E0B Amber
        "text_primary": RGBColor(30, 41, 59),  # #1E293B Charcoal Slate
        "text_secondary": RGBColor(100, 116, 139), # #64748B
        "text_muted": RGBColor(160, 174, 192),
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI"
    }

    # --- 3. LIGHT ECO EMERALD (Môi trường, Tái chế rác, Sinh thái, Y tế) ---
    LIGHT_ECO_EMERALD = {
        "name": "Light Eco Emerald",
        "is_dark": False,
        "bg_hex": "F6FBF7",                   # Clean mint tinted snow
        "card_bg_hex": "FFFFFF",
        "card_alpha": "92000",
        "border_hex": "059669",               # Fresh emerald border
        "border_alpha": "50000",
        "accent": RGBColor(5, 150, 105),      # #059669 Forest Emerald
        "accent_secondary": RGBColor(14, 165, 233), # #0EA5E9 Sky Water
        "accent_tertiary": RGBColor(217, 119, 6),   # #D97706 Warm Earth
        "text_primary": RGBColor(6, 78, 59),  # #064E3B Deep Forest Charcoal
        "text_secondary": RGBColor(52, 98, 81),
        "text_muted": RGBColor(120, 155, 140),
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI"
    }

    # --- 4. LIGHT EDITORIAL WARM (Văn học, Lịch sử, Xã hội) ---
    LIGHT_EDITORIAL_WARM = {
        "name": "Light Editorial Warm",
        "is_dark": False,
        "bg_hex": "FAF7F2",                   # Warm ivory book paper
        "card_bg_hex": "FFFFFF",
        "card_alpha": "90000",
        "border_hex": "C2410C",               # Terracotta red (#C2410C)
        "border_alpha": "50000",
        "accent": RGBColor(194, 65, 12),      # #C2410C Terracotta
        "accent_secondary": RGBColor(217, 119, 6), # #D97706 Amber Ochre
        "accent_tertiary": RGBColor(79, 70, 229),
        "text_primary": RGBColor(41, 37, 36), # #292524 Deep Espresso Stone
        "text_secondary": RGBColor(87, 83, 78),
        "text_muted": RGBColor(168, 162, 158),
        "font_heading": "Georgia",
        "font_body": "Segoe UI"
    }

    # --- 5. SPOTIFY CYBER (Phong cách giao diện đen tuyền âm nhạc & công nghệ) ---
    SPOTIFY_CYBER = {
        "name": "Spotify Cyber Dark",
        "is_dark": True,
        "bg_hex": "000000",                   # Pitch black
        "card_bg_hex": "121212",              # Spotify card surface
        "card_alpha": "88000",
        "border_hex": "1DB954",               # Spotify neon green
        "border_alpha": "75000",
        "accent": RGBColor(29, 185, 84),      # #1DB954 Green
        "accent_secondary": RGBColor(140, 92, 246),
        "accent_tertiary": RGBColor(0, 212, 255),
        "text_primary": RGBColor(255, 255, 255),
        "text_secondary": RGBColor(179, 179, 179),
        "text_muted": RGBColor(115, 115, 115),
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI"
    }

    # --- 6. DEEP COSMOS OBSIDIAN (Vũ trụ, Khoa học không gian, Thiên văn) ---
    COSMOS_DEEP = {
        "name": "Deep Cosmos Dark",
        "is_dark": True,
        "bg_hex": "02040A",
        "card_bg_hex": "080F1E",
        "card_alpha": "80000",
        "border_hex": "00E5FF",               # Electric cyan
        "border_alpha": "75000",
        "accent": RGBColor(0, 229, 255),
        "accent_secondary": RGBColor(255, 184, 0),
        "accent_tertiary": RGBColor(168, 85, 247),
        "text_primary": RGBColor(250, 252, 255),
        "text_secondary": RGBColor(160, 180, 210),
        "text_muted": RGBColor(100, 120, 150),
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI"
    }


# ==========================================
# 2. TRANSITION SUITE INJECTION ENGINE
# ==========================================

def apply_transition(slide, trans_type: str = "morph", duration_ms: int = 1600, **kwargs):
    """
    Injects high-impact native PowerPoint transitions into a slide:
    - 'morph': Smooth 60 FPS shape/element interpolation (options: byObject, byWord, byChar)
    - 'fade': Smooth cinematic dissolve
    - 'circle': Iris spotlight circle reveal
    - 'split': Theater curtain split
    - 'push': Spatial camera pan (dir='u', 'd', 'l', 'r')
    - 'wipe': Crisp progressive wipe
    """
    sld_elem = slide.element
    for child in list(sld_elem):
        if child.tag.endswith('transition') or child.tag.endswith('AlternateContent'):
            sld_elem.remove(child)

    trans_type = trans_type.lower()
    
    if trans_type == "fade":
        inner_xml = '<p:fade/>'
    elif trans_type == "circle":
        inner_xml = '<p:circle/>'
    elif trans_type == "split":
        orient = kwargs.get("orient", "vert")
        inner_xml = f'<p:split orient="{orient}"/>'
    elif trans_type == "push":
        direction = kwargs.get("dir", "u")
        inner_xml = f'<p:push dir="{direction}"/>'
    elif trans_type == "wipe":
        direction = kwargs.get("dir", "l")
        inner_xml = f'<p:wipe dir="{direction}"/>'
    else:
        # Default to Morph (The universal 60 FPS engine)
        mode = kwargs.get("mode", "byObject")
        inner_xml = f'<p159:morph option="{mode}"/>'

    trans_xml = (
        f'<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
        f'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
        f'spd="slow" p14:dur="{duration_ms}">'
        f'{inner_xml}'
        f'</p:transition>'
    )
    trans_node = parse_xml(trans_xml)
    
    clrMapOvr = None
    for child in sld_elem:
        if child.tag.endswith('clrMapOvr'):
            clrMapOvr = child
            break
            
    if clrMapOvr is not None:
        clrMapOvr.addnext(trans_node)
    else:
        sld_elem.append(trans_node)


def apply_morph(slide, duration_ms: int = 1600, option: str = "byObject"):
    """Helper alias for apply_transition(slide, 'morph', duration_ms, mode=option)."""
    apply_transition(slide, trans_type="morph", duration_ms=duration_ms, mode=option)


def link_morph(shape_a, shape_b, name_tag: str):
    """Assigns identical !!<tag> to guarantee 60 FPS Morph interpolation across slides."""
    tag = f"!!{name_tag}" if not name_tag.startswith("!!") else name_tag
    shape_a.name = tag
    shape_b.name = tag


# ==========================================
# 3. MOTION DECK BUILDER (API V4.0)
# ==========================================

class MotionDeck:
    """
    Widescreen 16:9 Presentation Builder with Bright Themes, 
    Frosted Glassmorphism, 3D Macro Dolly, Spotify UI Shells,
    Thematic Traveling Entities, and 60 FPS Video Export.
    """
    
    def __init__(self, theme: dict = Themes.LIGHT_ACADEMIC_CYAN):
        self.prs = pptx.Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]
        self.theme = theme
        self.slides = []
        
    def add_slide(self, transition: str = "morph", duration_ms: int = 1600, **kwargs):
        """Creates a new slide with clean theme background and attaches native transition."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        bg = slide.background
        fill = bg.fill
        fill.solid()
        bg_hex = self.theme.get("bg_hex", "FFFFFF")
        fill.fore_color.rgb = RGBColor(int(bg_hex[0:2], 16), int(bg_hex[2:4], 16), int(bg_hex[4:6], 16))
        
        # Apply transition (skip on slide 1 if desired, or apply starting from slide 2)
        if len(self.slides) > 0 and transition != "none":
            apply_transition(slide, trans_type=transition, duration_ms=duration_ms, **kwargs)
            
        self.slides.append(slide)
        return slide

    def add_thematic_traveling_entity(self, slide, asset_path: Optional[str] = None, 
                                     shape_type: Optional[MSO_SHAPE] = None,
                                     left: float = 1.0, top: float = 1.0, 
                                     width: float = 1.5, height: float = 1.5, 
                                     rot_deg: float = 0, name: str = "!!Traveling_Hero",
                                     fill_color: Optional[RGBColor] = None):
        """
        Inserts a persistent Traveling Hero Entity that automatically adapts to the topic:
        - Image mode: Loads any PNG/SVG asset (satellite, vinyl, leaf, arrow, molecule, logo).
        - Vector mode: Creates a crisp geometric vector shape if asset_path is not specified.
        Strictly preserves '!!' Morph locking across slides with dynamic rotation.
        """
        if asset_path and os.path.exists(asset_path):
            pic = slide.shapes.add_picture(asset_path, Inches(left), Inches(top), Inches(width), Inches(height))
            pic.name = name
            if rot_deg != 0:
                spPr = pic._element.spPr
                xfrm = spPr.xfrm
                xfrm.set('rot', str(int(rot_deg * 60000)))
            return pic
        else:
            # Fallback to topic vector shape (e.g. Star, Oval, Arrow, Octagon)
            st = shape_type or MSO_SHAPE.OVAL
            shp = slide.shapes.add_shape(st, Inches(left), Inches(top), Inches(width), Inches(height))
            shp.name = name
            shp.fill.solid()
            shp.fill.fore_color.rgb = fill_color or self.theme["accent"]
            shp.line.color.rgb = RGBColor(255, 255, 255)
            shp.line.width = Pt(1.5)
            if rot_deg != 0:
                spPr = shp._element.spPr
                xfrm = spPr.xfrm
                xfrm.set('rot', str(int(rot_deg * 60000)))
            return shp

    # -------------------------------------------------------------
    # SPOTIFY RADIAL MORPH & APP SHELL UNLOCK (from MAC-LENIN)
    # -------------------------------------------------------------
    def create_spotify_style_unlock(self, brand_title: str = "SPOTIFY", 
                                    subtitle: str = "MUSIC FOR EVERYONE", 
                                    presenter: str = "EXECUTIVE LABS",
                                    emblem_asset: Optional[str] = None):
        """
        Creates the iconic 3-slide Spotify Radial Morph & App Shell opening:
        - Slide 1: Giant ambient orb center, emblem hero, text staged off-screen.
        - Slide 2: Morph! Orb contracts to corner, titles swoop in from margin to lock up.
        - Slide 3: Morph! Expands into interactive Application Shell (Sidebar + Stage).
        """
        # --- SLIDE 1: THE BRAND MONOLITH ---
        s1 = self.add_slide(transition="none")
        # Massive background radial orb
        orb1 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-3.13), Inches(-2.55), Inches(19.59), Inches(11.97))
        orb1.name = "!!ambient_orb_1"
        orb1.fill.solid()
        orb_color = self.theme["accent"]
        orb1.fill.fore_color.rgb = orb_color
        orb1.line.fill.background()
        
        # Center brand emblem / icon
        if emblem_asset and os.path.exists(emblem_asset):
            s1.shapes.add_picture(emblem_asset, Inches(5.83), Inches(2.61), Inches(1.67), Inches(1.66)).name = "!!brand_logo"
        else:
            emb = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.83), Inches(2.61), Inches(1.67), Inches(1.66))
            emb.name = "!!brand_logo"
            emb.fill.solid()
            emb.fill.fore_color.rgb = RGBColor(255, 255, 255)
            emb.line.fill.background()
            
        # Titles staged OFF-SCREEN RIGHT
        tx_title1 = s1.shapes.add_textbox(Inches(14.96), Inches(2.52), Inches(5.0), Inches(1.11))
        tx_title1.name = "!!hero_title"
        tx_title1.text_frame.text = brand_title
        
        tx_sub1 = s1.shapes.add_textbox(Inches(14.96), Inches(3.63), Inches(5.0), Inches(0.64))
        tx_sub1.name = "!!hero_sub"
        tx_sub1.text_frame.text = presenter
        
        # --- SLIDE 2: THE CONVERGENCE LOCKUP ---
        s2 = self.add_slide(transition="morph", duration_ms=1600)
        # Orb contracts and moves to top-left
        orb2 = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-3.27), Inches(-2.82), Inches(6.18), Inches(6.26))
        orb2.name = "!!ambient_orb_1"
        orb2.fill.solid()
        orb2.fill.fore_color.rgb = orb_color
        orb2.line.fill.background()
        
        # Second ambient orb enters bottom-right
        orb_br = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.70), Inches(4.55), Inches(6.18), Inches(6.26))
        orb_br.name = "!!ambient_orb_2"
        orb_br.fill.solid()
        orb_br.fill.fore_color.rgb = self.theme["accent_secondary"]
        orb_br.line.fill.background()
        
        # Center emblem
        if emblem_asset and os.path.exists(emblem_asset):
            s2.shapes.add_picture(emblem_asset, Inches(5.83), Inches(2.61), Inches(1.67), Inches(1.66)).name = "!!brand_logo"
        else:
            emb2 = s2.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.83), Inches(2.61), Inches(1.67), Inches(1.66))
            emb2.name = "!!brand_logo"
            emb2.fill.solid()
            emb2.fill.fore_color.rgb = RGBColor(255, 255, 255)
            emb2.line.fill.background()
            
        # Titles swoop into place from offscreen right!
        tx_title2 = s2.shapes.add_textbox(Inches(7.73), Inches(2.52), Inches(4.5), Inches(1.11))
        tx_title2.name = "!!hero_title"
        p = tx_title2.text_frame.paragraphs[0]
        p.text = brand_title
        p.font.name = self.theme["font_heading"]
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = self.theme["text_primary"]
        
        tx_sub2 = s2.shapes.add_textbox(Inches(7.73), Inches(3.63), Inches(4.5), Inches(0.64))
        tx_sub2.name = "!!hero_sub"
        p_sub = tx_sub2.text_frame.paragraphs[0]
        p_sub.text = presenter.upper()
        p_sub.font.name = self.theme["font_body"]
        p_sub.font.size = Pt(13)
        p_sub.font.bold = True
        p_sub.font.color.rgb = self.theme["accent"]
        
        # --- SLIDE 3: THE APP SHELL / DASHBOARD UNLOCK ---
        s3 = self.add_slide(transition="morph", duration_ms=1600)
        # Left Sidebar Navigation Bar
        sidebar = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.0), Inches(0.0), Inches(3.29), Inches(7.50))
        sidebar.name = "!!app_sidebar"
        sidebar.fill.solid()
        sidebar_bg = RGBColor(18, 18, 18) if self.theme.get("is_dark", False) else RGBColor(241, 245, 249)
        sidebar.fill.fore_color.rgb = sidebar_bg
        sidebar.line.fill.background()
        
        # Sidebar branding
        tx_side = s3.shapes.add_textbox(Inches(0.4), Inches(0.6), Inches(2.5), Inches(0.8))
        ps = tx_side.text_frame.paragraphs[0]
        ps.text = brand_title
        ps.font.name = self.theme["font_heading"]
        ps.font.size = Pt(18)
        ps.font.bold = True
        ps.font.color.rgb = self.theme["accent"]
        
        # Top Header Bar User Badge
        badge = self.add_badge(s3, "PRO // VERIFIED", left=10.8, top=0.3, width=2.1, height=0.38,
                               bg_color=self.theme["accent"], name="!!top_badge")
                               
        # Main Viewport Stage Card
        stage = self.add_glass_card(s3, left=3.6, top=1.2, width=9.2, height=5.8, name="!!app_stage",
                                    border_width_pt=1.5, corner_radius=6000)
                                    
        tx_stage = s3.shapes.add_textbox(Inches(4.0), Inches(1.6), Inches(8.4), Inches(1.2))
        tx_stage.name = "!!stage_title"
        pt = tx_stage.text_frame.paragraphs[0]
        pt.text = "TỔNG QUAN HỆ THỐNG"
        pt.font.name = self.theme["font_heading"]
        pt.font.size = Pt(24)
        pt.font.bold = True
        pt.font.color.rgb = self.theme["accent"]
        
        pt2 = tx_stage.text_frame.add_paragraph()
        pt2.space_before = Pt(4)
        pt2.text = subtitle
        pt2.font.name = self.theme["font_body"]
        pt2.font.size = Pt(12)
        pt2.font.color.rgb = self.theme["text_secondary"]
        
        return s1, s2, s3

    # -------------------------------------------------------------
    # 3D MACRO DOLLY ZOOM & ORBITAL PAN (from Earth & Space)
    # -------------------------------------------------------------
    def create_3d_macro_dolly_sequence(self, hero_asset: str, title: str, subtitle: str,
                                       category: str = "KHÁM PHÁ CHI TIẾT"):
        """
        Creates the Earth & Space 3D camera dolly zoom (100% -> 180% -> 350%):
        - Slide 1: Wide establishing shot (hero centered at normal scale).
        - Slide 2: Morph approach & rotation (hero scales up 1.8x, draws closer).
        - Slide 3: Morph touchdown macro push-in (hero scales up 3.5x, fills viewport).
        """
        # --- SLIDE 1: WIDE ESTABLISHING SHOT (100%) ---
        s1 = self.add_slide(transition="none")
        self.add_badge(s1, category, left=1.0, top=1.0, width=3.0, height=0.38, name="!!dolly_badge")
        self.add_hero_title(s1, title, subtitle=subtitle, left=1.0, top=1.55, width=6.5, height=1.6, name="!!dolly_title")
        
        # Hero asset at 100% scale
        if os.path.exists(hero_asset):
            s1.shapes.add_picture(hero_asset, Inches(7.5), Inches(1.8), Inches(5.0), Inches(5.0)).name = "!!3d_dolly_hero"
            
        # --- SLIDE 2: APPROACH & ROTATION (180%) ---
        s2 = self.add_slide(transition="morph", duration_ms=1600)
        self.add_badge(s2, category, left=1.0, top=1.0, width=3.0, height=0.38, name="!!dolly_badge")
        self.add_hero_title(s2, title, subtitle=subtitle, left=1.0, top=1.55, width=6.5, height=1.6, name="!!dolly_title")
        
        # Hero draws closer, scales to 8.5" and moves to center-right
        if os.path.exists(hero_asset):
            pic2 = s2.shapes.add_picture(hero_asset, Inches(6.0), Inches(0.8), Inches(8.5), Inches(8.5))
            pic2.name = "!!3d_dolly_hero"
            # Slight dynamic rotation
            pic2._element.spPr.xfrm.set('rot', str(int(-15 * 60000)))
            
        # --- SLIDE 3: TOUCHDOWN MACRO PUSH-IN (350%) ---
        s3 = self.add_slide(transition="morph", duration_ms=1800)
        # Hero plunges forward, expanding to 18.0 inches!
        if os.path.exists(hero_asset):
            pic3 = s3.shapes.add_picture(hero_asset, Inches(2.0), Inches(-3.0), Inches(18.0), Inches(18.0))
            pic3.name = "!!3d_dolly_hero"
            pic3._element.spPr.xfrm.set('rot', str(int(-30 * 60000)))
            
        # Foreground Telemetry HUD Card
        hud = self.add_glass_card(s3, left=1.0, top=4.0, width=5.5, height=2.8, name="!!hud_macro_card",
                                  border_width_pt=1.5, corner_radius=6000)
        tx_hud = s3.shapes.add_textbox(Inches(1.3), Inches(4.3), Inches(4.9), Inches(2.2))
        tf_h = tx_hud.text_frame
        tf_h.word_wrap = True
        ph1 = tf_h.paragraphs[0]
        ph1.text = "CẬN CẢNH BỀ MẶT & ĐẶC TÍNH"
        ph1.font.name = self.theme["font_heading"]
        ph1.font.size = Pt(14)
        ph1.font.bold = True
        ph1.font.color.rgb = self.theme["accent"]
        
        ph2 = tf_h.add_paragraph()
        ph2.space_before = Pt(4)
        ph2.text = "Góc nhìn cận cảnh mô tả chi tiết vi cấu trúc, vật liệu hoặc địa tầng thực địa với độ phân giải siêu cao."
        ph2.font.name = self.theme["font_body"]
        ph2.font.size = Pt(11)
        ph2.font.color.rgb = self.theme["text_primary"]
        
        return s1, s2, s3

    def add_glass_card(self, slide, left: float, top: float, width: float, height: float, 
                       name: Optional[str] = None, bg_hex: Optional[str] = None, 
                       alpha: Optional[str] = None, border_hex: Optional[str] = None, 
                       border_alpha: Optional[str] = None, border_width_pt: float = 1.25,
                       corner_radius: int = 5000):
        """Creates an authentic Frosted Glassmorphism container with OpenXML alpha transparency."""
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        if name:
            card.name = name
            
        b_hex = bg_hex or self.theme.get("card_bg_hex", "FFFFFF")
        a_val = alpha or self.theme.get("card_alpha", "92000")
        bo_hex = border_hex or self.theme.get("border_hex", "0284C7")
        bo_a_val = border_alpha or self.theme.get("border_alpha", "55000")
        ln_w = int(border_width_pt * 12700)
        
        sp_xml = (
            f'<p:spPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:xfrm><a:off x="{int(left*914400)}" y="{int(top*914400)}"/>'
            f'<a:ext cx="{int(width*914400)}" cy="{int(height*914400)}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {corner_radius}"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{b_hex}"><a:alpha val="{a_val}"/></a:srgbClr></a:solidFill>'
            f'<a:ln w="{ln_w}"><a:solidFill><a:srgbClr val="{bo_hex}"><a:alpha val="{bo_a_val}"/></a:srgbClr></a:solidFill></a:ln>'
            f'</p:spPr>'
        )
        card._element.replace(card._element.spPr, parse_xml(sp_xml))
        return card

    def add_badge(self, slide, text: str, left: float, top: float, width: float = 2.4, height: float = 0.38,
                  bg_color: Optional[RGBColor] = None, text_color: Optional[RGBColor] = None, name: Optional[str] = None):
        """Creates an illuminated pill badge with crisp typography."""
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        if name:
            pill.name = name
        pill.fill.solid()
        pill.fill.fore_color.rgb = bg_color or self.theme["accent"]
        pill.line.fill.background()
        
        tf = pill.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = text.upper()
        p.font.name = self.theme["font_heading"]
        p.font.size = Pt(10)
        p.font.bold = True
        pill_text_clr = text_color or (RGBColor(255, 255, 255) if not self.theme.get("is_dark", False) else RGBColor(10, 10, 15))
        p.font.color.rgb = pill_text_clr
        return pill

    def add_hero_title(self, slide, title: str, subtitle: Optional[str] = None, 
                       left: float = 1.0, top: float = 1.0, width: float = 10.0, height: float = 1.8,
                       title_size: int = 32, subtitle_size: int = 13, 
                       title_color: Optional[RGBColor] = None, name: Optional[str] = None):
        """Adds crisp non-overlapping headline typography."""
        tx_box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        if name:
            tx_box.name = name
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = self.theme["font_heading"]
        p.font.size = Pt(title_size)
        p.font.bold = True
        p.font.color.rgb = title_color or self.theme["accent"]
        p.space_after = Pt(8)
        
        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.name = self.theme["font_body"]
            p2.font.size = Pt(subtitle_size)
            p2.font.color.rgb = self.theme["text_secondary"]
            
        return tx_box

    def add_metric_card(self, slide, title: str, value: str, desc: str, 
                        left: float, top: float, width: float, height: float, 
                        accent_color: Optional[RGBColor] = None, name: Optional[str] = None):
        """Creates a well-padded metric card with clean internal hierarchy."""
        c = self.add_glass_card(slide, left, top, width, height, name=name)
        
        tx = slide.shapes.add_textbox(Inches(left + 0.25), Inches(top + 0.16), Inches(width - 0.5), Inches(height - 0.25))
        tf = tx.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p1 = tf.paragraphs[0]
        p1.text = title.upper()
        p1.font.name = self.theme["font_body"]
        p1.font.size = Pt(9)
        p1.font.bold = True
        p1.font.color.rgb = self.theme["text_muted"]
        
        p2 = tf.add_paragraph()
        p2.space_before = Pt(3)
        p2.text = value
        p2.font.name = self.theme["font_heading"]
        p2.font.size = Pt(20)
        p2.font.bold = True
        p2.font.color.rgb = accent_color or self.theme["accent"]
        
        p3 = tf.add_paragraph()
        p3.space_before = Pt(4)
        p3.text = desc
        p3.font.name = self.theme["font_body"]
        p3.font.size = Pt(10.5)
        p3.font.color.rgb = self.theme["text_secondary"]
        return c, tx

    def add_progress_bar(self, slide, current_step: int, total_steps: int, 
                         left: float = 1.0, top: float = 6.9, width: float = 11.333, height: float = 0.08,
                         name_track: str = "!!prog_track", name_fill: str = "!!prog_fill"):
        """Bottom progressive scrubber bar."""
        track = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        track.name = name_track
        track.fill.solid()
        track_color = RGBColor(226, 232, 240) if not self.theme.get("is_dark", False) else RGBColor(30, 30, 40)
        track.fill.fore_color.rgb = track_color
        track.line.fill.background()
        
        fill_pct = max(0.05, min(1.0, current_step / total_steps))
        fill_width = width * fill_pct
        fill_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(fill_width), Inches(height))
        fill_bar.name = name_fill
        fill_bar.fill.solid()
        fill_bar.fill.fore_color.rgb = self.theme["accent"]
        fill_bar.line.fill.background()
        return track, fill_bar

    def save(self, output_path: str):
        self.prs.save(output_path)
        print(f"[OK] Saved presentation: {output_path}")


# ==========================================
# 4. HIGH-SPEED VIDEO EXPORT ENGINE
# ==========================================

def export_deck_to_video(pptx_path: str, video_path: str, slide_duration: float = 3.5, 
                         fps: int = 60, resolution: int = 1080) -> str:
    """
    Renders the PowerPoint deck into a high-definition 60 FPS MP4 video
    utilizing native PowerPoint hardware acceleration and transition interpolation.
    """
    import win32com.client
    
    abs_pptx = os.path.abspath(pptx_path)
    abs_video = os.path.abspath(video_path)
    os.makedirs(os.path.dirname(abs_video), exist_ok=True)
    
    if os.path.exists(abs_video):
        try:
            os.remove(abs_video)
        except Exception:
            pass

    print(f"[*] Launching PowerPoint COM engine for video render...")
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    try:
        ppt.Visible = True
    except Exception:
        pass
        
    deck = ppt.Presentations.Open(abs_pptx)
    
    print(f"[*] Rendering '{os.path.basename(abs_pptx)}' -> '{os.path.basename(abs_video)}' ({fps} FPS, {resolution}p)...")
    deck.CreateVideo(abs_video, -1, int(slide_duration), resolution, fps, 90)
    
    time.sleep(1.5)
    start_time = time.time()
    while True:
        status = deck.CreateVideoStatus
        if status != 1:
            break
        time.sleep(0.8)
        if time.time() - start_time > 180:
            break
            
    time.sleep(2.0)
    try:
        deck.Close()
    except Exception:
        pass
    try:
        ppt.Quit()
    except Exception:
        pass
    
    if os.path.exists(abs_video) and os.path.getsize(abs_video) > 50000:
        v_size = os.path.getsize(abs_video)
        print(f"[+] Video render complete! Saved: {abs_video} ({v_size:,} bytes)")
        return abs_video
    else:
        raise RuntimeError(f"Video file was not successfully generated at: {abs_video}")
