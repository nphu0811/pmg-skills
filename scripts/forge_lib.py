"""
forge_lib.py (v3.0 - 3D Cinematic Visual-First Motion Engine)
============================================================
The core orchestrator library for building 3D Industrial Explainer and 
Cinematic Scientific presentations in PowerPoint (.pptx).

Built strictly according to the Visual-First philosophy:
- 3D Imagery & Animations occupy 85% - 90% of screen area
- Minimalist text (< 25 words/slide, < 10% screen area)
- 3D HUD Callout Leader Lines with Target Dots
- Multi-Plane Parallax Camera Coordinate System
- Energy Step Tracker
- Shape-First -> Content-Second timing discipline
"""

import os
import sys
from typing import Optional, Tuple, List, Union
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml


# ==========================================
# 1. 3D THERMAL & INDUSTRIAL COLOR PALETTE
# ==========================================

class IndustrialPalette:
    """Accurate physical energy and industrial thermal colors."""
    
    # Background & Structure
    BG_CARBON = RGBColor(11, 17, 30)         # #0B111E Dark Carbon Blue
    BG_STUDIO_HEX = "0B111E"
    STEEL_TITANIUM = RGBColor(142, 154, 175)  # #8E9AAF Steel gray
    SILVER_METALLIC = RGBColor(220, 225, 233) # #DCE1E9 Metallic silver
    
    # Physical Energy States
    FIRE_RED = RGBColor(255, 59, 48)          # #FF3B30 Combustion flame
    FIRE_RED_HEX = "FF3B30"
    THERMAL_ORANGE = RGBColor(255, 149, 0)    # #FF9500 High-heat steam
    THERMAL_ORANGE_HEX = "FF9500"
    ELECTRIC_GOLD = RGBColor(255, 214, 10)    # #FFD60A Rotation & Electricity
    ELECTRIC_GOLD_HEX = "FFD60A"
    COOLING_CYAN = RGBColor(0, 245, 212)      # #00F5D4 Cooling water / Condenser
    COOLING_CYAN_HEX = "00F5D4"
    DEEP_BLUE = RGBColor(0, 168, 232)         # #00A8E8 Water reservoir
    
    # Text
    TEXT_WHITE = RGBColor(255, 255, 255)
    TEXT_SILVER = RGBColor(210, 218, 230)
    TEXT_MUTED = RGBColor(120, 135, 160)
    
    FONT_HEADING = "Segoe UI"
    FONT_BODY = "Segoe UI"


# ==========================================
# 2. MORPH TRANSITION INJECTOR
# ==========================================

def apply_morph(slide, duration_ms: int = 1200, option: str = "byObject"):
    """Injects high-precision Morph transition for mechanical cameras."""
    sld_elem = slide.element
    for child in list(sld_elem):
        if child.tag.endswith('transition'):
            sld_elem.remove(child)
            
    morph_xml = (
        f'<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
        f'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
        f'spd="slow" p14:dur="{duration_ms}">'
        f'<p159:morph option="{option}"/>'
        f'</p:transition>'
    )
    trans_node = parse_xml(morph_xml)
    
    clrMapOvr = None
    for child in sld_elem:
        if child.tag.endswith('clrMapOvr'):
            clrMapOvr = child
            break
            
    if clrMapOvr is not None:
        clrMapOvr.addnext(trans_node)
    else:
        sld_elem.append(trans_node)


# ==========================================
# 3. 3D FORGE ENGINE
# ==========================================

class ForgeDeck:
    """
    State-of-the-art 3D Motion Graphics Presentation Orchestrator.
    Configured for 16:9 Widescreen (13.333 x 7.5 inches).
    """
    
    def __init__(self, bg_hex: str = IndustrialPalette.BG_STUDIO_HEX):
        self.prs = pptx.Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]
        self.bg_hex = bg_hex
        self.slides = []
        
    def add_slide(self, enable_morph: bool = True, duration_ms: int = 1200):
        """Adds a slide with pure carbon dark background and morph transition."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        bg = slide.background
        fill = bg.fill
        fill.solid()
        r = int(self.bg_hex[0:2], 16)
        g = int(self.bg_hex[2:4], 16)
        b = int(self.bg_hex[4:6], 16)
        fill.fore_color.rgb = RGBColor(r, g, b)
        
        if enable_morph and len(self.slides) > 0:
            apply_morph(slide, duration_ms=duration_ms)
            
        self.slides.append(slide)
        return slide

    # -------------------------------------------------------------
    # CORNER HEADER BADGE (< 5% OF SCREEN SPACE)
    # -------------------------------------------------------------
    def add_corner_badge_header(self, slide, scene_num: str, title: str, category: str, 
                                accent_color: RGBColor = IndustrialPalette.THERMAL_ORANGE,
                                accent_hex: str = IndustrialPalette.THERMAL_ORANGE_HEX,
                                name_prefix: str = "hdr"):
        """
        Creates a minimalist, compact HUD corner header in the top-left corner.
        Strictly limits text to 3-5 words to preserve 90% screen space for 3D visuals.
        """
        # Outer Glass Container
        c_w, c_h = 4.2, 0.95
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.6), Inches(c_w), Inches(c_h))
        card.name = f"!!{name_prefix}_card"
        
        sp_xml = (
            f'<p:spPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:xfrm><a:off x="{int(0.8*914400)}" y="{int(0.6*914400)}"/>'
            f'<a:ext cx="{int(c_w*914400)}" cy="{int(c_h*914400)}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 4000"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="060A14"><a:alpha val="85000"/></a:srgbClr></a:solidFill>'
            f'<a:ln w="15875"><a:solidFill><a:srgbClr val="{accent_hex}"><a:alpha val="80000"/></a:srgbClr></a:solidFill></a:ln>'
            f'</p:spPr>'
        )
        card._element.replace(card._element.spPr, parse_xml(sp_xml))
        
        # Text inside header
        tx = slide.shapes.add_textbox(Inches(0.95), Inches(0.68), Inches(3.9), Inches(0.8))
        tx.name = f"!!{name_prefix}_txt"
        tf = tx.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        # Category / Scene pill
        p1 = tf.paragraphs[0]
        p1.text = f"{scene_num}  •  {category}".upper()
        p1.font.name = IndustrialPalette.FONT_HEADING
        p1.font.size = Pt(8.5)
        p1.font.bold = True
        p1.font.color.rgb = accent_color
        
        # Main short title
        p2 = tf.add_paragraph()
        p2.space_before = Pt(2)
        p2.text = title
        p2.font.name = IndustrialPalette.FONT_HEADING
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = IndustrialPalette.TEXT_WHITE
        return card, tx

    # -------------------------------------------------------------
    # 3D HUD CALLOUT WITH TARGET DOT & LEADER LINE
    # -------------------------------------------------------------
    def add_3d_callout(self, slide, target_x: float, target_y: float, 
                       callout_x: float, callout_y: float, 
                       part_name: str, verb_desc: str, 
                       accent_color: RGBColor = IndustrialPalette.THERMAL_ORANGE,
                       accent_hex: str = IndustrialPalette.THERMAL_ORANGE_HEX,
                       name_id: str = "callout_1"):
        """
        Draws a professional 3D technical leader line pointing directly to a machine part:
        1. Glowing Target Dot at (target_x, target_y)
        2. Angled Leader Line from target to callout box
        3. Frosted Glass HUD card with Part Name (Bold) + Short Verb Action (< 8 words)
        """
        # 1. Outer Target Dot Pulse (0.24" circle)
        dot_pulse = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, 
            Inches(target_x - 0.12), Inches(target_y - 0.12), Inches(0.24), Inches(0.24)
        )
        dot_pulse.name = f"!!{name_id}_dot_pulse"
        dot_pulse.fill.solid()
        dot_pulse.fill.fore_color.rgb = accent_color
        dot_pulse.line.color.rgb = IndustrialPalette.TEXT_WHITE
        dot_pulse.line.width = Pt(1.0)
        
        # 2. Inner Glowing Core (0.10" circle)
        dot_core = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, 
            Inches(target_x - 0.05), Inches(target_y - 0.05), Inches(0.10), Inches(0.10)
        )
        dot_core.name = f"!!{name_id}_dot_core"
        dot_core.fill.solid()
        dot_core.fill.fore_color.rgb = IndustrialPalette.TEXT_WHITE
        dot_core.line.fill.background()
        
        # 3. Leader Line (from target_x, target_y to callout_x, callout_y + 0.3)
        # Using connector line
        line = slide.shapes.add_connector(
            pptx.enum.shapes.MSO_CONNECTOR.STRAIGHT,
            Inches(target_x), Inches(target_y),
            Inches(callout_x), Inches(callout_y + 0.35)
        )
        line.name = f"!!{name_id}_line"
        line.line.color.rgb = accent_color
        line.line.width = Pt(1.5)
        
        # 4. Glass HUD Card
        card_w, card_h = 2.8, 0.75
        hud_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, 
            Inches(callout_x), Inches(callout_y), Inches(card_w), Inches(card_h)
        )
        hud_card.name = f"!!{name_id}_box"
        
        sp_xml = (
            f'<p:spPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:xfrm><a:off x="{int(callout_x*914400)}" y="{int(callout_y*914400)}"/>'
            f'<a:ext cx="{int(card_w*914400)}" cy="{int(card_h*914400)}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 3500"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="081020"><a:alpha val="88000"/></a:srgbClr></a:solidFill>'
            f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{accent_hex}"><a:alpha val="75000"/></a:srgbClr></a:solidFill></a:ln>'
            f'</p:spPr>'
        )
        hud_card._element.replace(hud_card._element.spPr, parse_xml(sp_xml))
        
        # 5. Callout Text (Shape-First -> Content-Second)
        tx = slide.shapes.add_textbox(Inches(callout_x + 0.15), Inches(callout_y + 0.1), Inches(card_w - 0.3), Inches(0.55))
        tx.name = f"!!{name_id}_txt"
        tf = tx.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        # Part Name
        p1 = tf.paragraphs[0]
        p1.text = part_name.upper()
        p1.font.name = IndustrialPalette.FONT_HEADING
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = accent_color
        
        # Short Action Verb (< 8 words)
        p2 = tf.add_paragraph()
        p2.space_before = Pt(2)
        p2.text = verb_desc
        p2.font.name = IndustrialPalette.FONT_BODY
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = IndustrialPalette.TEXT_WHITE
        
        return dot_pulse, dot_core, line, hud_card, tx

    # -------------------------------------------------------------
    # BOTTOM ENERGY STEP TRACKER
    # -------------------------------------------------------------
    def add_energy_tracker(self, slide, steps: List[Tuple[str, str]], active_idx: int,
                           left: float = 2.0, top: float = 6.85, total_width: float = 9.333):
        """
        Renders a streamlined technical process tracker along the bottom edge:
        [ HÓA NĂNG ] -> [ NHIỆT NĂNG ] -> [ CƠ NĂNG ] -> [ ĐIỆN NĂNG ] -> [ LƯỚI ĐIỆN ]
        Active step is brightly illuminated; past steps are subtly lit; future steps are muted.
        """
        n = len(steps)
        gap = 0.2
        item_w = (total_width - (n - 1) * gap) / n
        item_h = 0.32
        
        shapes = []
        for i, (code, label) in enumerate(steps):
            x = left + i * (item_w + gap)
            is_active = (i == active_idx)
            is_past = (i < active_idx)
            
            pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(top), Inches(item_w), Inches(item_h))
            pill.name = f"!!trk_step_{i}"
            
            if is_active:
                fill_hex = "FF9500"
                line_hex = "FFFFFF"
                txt_col = RGBColor(10, 10, 15)
                alpha_val = "95000"
            elif is_past:
                fill_hex = "00F5D4"
                line_hex = "00F5D4"
                txt_col = RGBColor(0, 245, 212)
                alpha_val = "40000"
            else:
                fill_hex = "101826"
                line_hex = "2A3850"
                txt_col = IndustrialPalette.TEXT_MUTED
                alpha_val = "60000"
                
            sp_xml = (
                f'<p:spPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                f'<a:xfrm><a:off x="{int(x*914400)}" y="{int(top*914400)}"/>'
                f'<a:ext cx="{int(item_w*914400)}" cy="{int(item_h*914400)}"/></a:xfrm>'
                f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 4500"/></a:avLst></a:prstGeom>'
                f'<a:solidFill><a:srgbClr val="{fill_hex}"><a:alpha val="{alpha_val}"/></a:srgbClr></a:solidFill>'
                f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line_hex}"><a:alpha val="70000"/></a:srgbClr></a:solidFill></a:ln>'
                f'</p:spPr>'
            )
            pill._element.replace(pill._element.spPr, parse_xml(sp_xml))
            
            tf = pill.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = label.upper()
            p.font.name = IndustrialPalette.FONT_HEADING
            p.font.size = Pt(8.5)
            p.font.bold = True
            p.font.color.rgb = txt_col
            shapes.append(pill)
            
        return shapes

    def save(self, output_path: str):
        self.prs.save(output_path)
        print(f"[FORGE_LIB] Successfully generated 3D cinematic presentation: {output_path}")
