"""
pmg.compiler
============
Consolidated Shared Native OOXML Compiler for PMG Skills (v2.1.0).

Produces 100% native, editable PowerPoint presentations without requiring Microsoft
PowerPoint Desktop. Reuses canonical PMG modules (`SlideAnimationTimeline`,
`apply_transition`, `Themes`, etc.) behind a clean, stable compiler interface.

Supports:
- 16:9 Widescreen canvas (13.333" x 7.5")
- Modern Bright/Light Themes as primary default (Academic Cyan, Modern Studio, Eco Emerald)
- Frosted glassmorphism acrylic cards (translucent alpha with hairline borders)
- 60 FPS Morph transitions via OpenXML (<p159:morph>)
- Persistent '!!' traveling entities and object identity matching
- Intra-slide native Animation Pane effects:
  * Entrances: appear, fade, fly_in, wipe, zoom
  * Emphasis: scale (grow_shrink), rotate (spin), transparency (opacity), pulse
  * Motion paths: linear, cubic Bezier (M ... C ... E)
  * Exits: disappear, fade_out, fly_out, wipe_out
  * Timing: with_previous, after_previous, on_click, delay_ms, duration_ms
- Zero picture-flattening or screenshot backgrounds
- Engineering honesty reporting for unsupported or unverified effects
"""

import os
import json
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Tuple, Union

import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml

from pmg.animations import (
    SlideAnimationTimeline,
    AnimationStep,
    normalize_effect,
    SUPPORTED_EFFECTS,
    P_NS,
    A_NS,
)
from pmg.animation_ir import (
    AnimationIR,
    SceneIR,
    ObjectIR,
    AnimationTrackIR,
    TimelineCompiler,
    MorphReliabilityValidator,
    AnimationCapabilityStatus,
    get_effect_capability_status,
)
from scripts.motion_graph_kit import (
    Themes,
    apply_transition,
    apply_morph,
    link_morph,
)


@dataclass
class CompilerReport:
    """Compilation metadata and engineering honesty record."""
    output_path: str
    total_slides: int = 0
    total_shapes: int = 0
    total_animations: int = 0
    morph_transition_count: int = 0
    named_anchors: List[str] = field(default_factory=list)
    unsupported_effects: List[str] = field(default_factory=list)
    approximate_effects: List[str] = field(default_factory=list)
    unverified_playback_effects: List[str] = field(default_factory=list)
    editable_shapes_count: int = 0
    picture_count: int = 0
    is_fully_native: bool = True
    timing_trees_attached: int = 0
    capability_breakdown: Dict[str, int] = field(default_factory=dict)
    morph_validation_report: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "output_path": self.output_path,
            "total_slides": self.total_slides,
            "total_shapes": self.total_shapes,
            "total_animations": self.total_animations,
            "morph_transition_count": self.morph_transition_count,
            "named_anchors": list(self.named_anchors),
            "unsupported_effects": list(self.unsupported_effects),
            "approximate_effects": list(self.approximate_effects),
            "unverified_playback_effects": list(self.unverified_playback_effects),
            "editable_shapes_count": self.editable_shapes_count,
            "picture_count": self.picture_count,
            "is_fully_native": self.is_fully_native,
            "timing_trees_attached": self.timing_trees_attached,
            "capability_breakdown": dict(self.capability_breakdown),
            "morph_validation_report": self.morph_validation_report,
        }


class NativePPTXCompiler:
    """
    Consolidated Native PowerPoint OOXML Compiler.
    Compiles declarative presentation specifications or presets into native PPTX files.
    """

    def __init__(self, theme: Optional[Dict[str, Any]] = None):
        self.theme = theme or Themes.LIGHT_ACADEMIC_CYAN
        self.prs: Optional[pptx.Presentation] = None
        self.blank_layout = None
        self.slides: List[Any] = []
        self.timelines: Dict[int, SlideAnimationTimeline] = {}
        self.report = CompilerReport(output_path="")

    def new_presentation(self, width_in: float = 13.333, height_in: float = 7.5) -> None:
        """Initializes a clean 16:9 widescreen presentation."""
        self.prs = pptx.Presentation()
        self.prs.slide_width = Inches(width_in)
        self.prs.slide_height = Inches(height_in)
        self.blank_layout = self.prs.slide_layouts[6]
        self.slides = []
        self.timelines = {}

    def add_slide(
        self,
        transition: str = "morph",
        duration_ms: int = 1400,
        bg_color_hex: Optional[str] = None,
    ) -> Any:
        """Adds a slide with solid background and native transition."""
        if self.prs is None:
            self.new_presentation()

        slide = self.prs.slides.add_slide(self.blank_layout)
        bg = slide.background
        fill = bg.fill
        fill.solid()

        hex_c = bg_color_hex or self.theme.get("bg_hex", "FFFFFF")
        try:
            r = int(hex_c[0:2], 16)
            g = int(hex_c[2:4], 16)
            b = int(hex_c[4:6], 16)
            fill.fore_color.rgb = RGBColor(r, g, b)
        except Exception:
            fill.fore_color.rgb = RGBColor(255, 255, 255)

        apply_transition(slide, trans_type=transition, duration_ms=duration_ms)
        if transition.lower() == "morph":
            self.report.morph_transition_count += 1

        self.slides.append(slide)
        slide_idx = len(self.slides)
        self.timelines[slide_idx] = SlideAnimationTimeline(slide_number=slide_idx)
        return slide

    def add_textbox(
        self,
        slide: Any,
        left_in: float,
        top_in: float,
        width_in: float,
        height_in: float,
        text: str,
        name: Optional[str] = None,
        font_size_pt: int = 18,
        font_bold: bool = False,
        font_color: Optional[RGBColor] = None,
        align: PP_ALIGN = PP_ALIGN.LEFT,
    ) -> Any:
        """Adds a native editable text box."""
        tb = slide.shapes.add_textbox(Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in))
        if name:
            tb.name = name
            if name.startswith("!!"):
                self.report.named_anchors.append(name)
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.alignment = align
        p.font.size = Pt(font_size_pt)
        p.font.bold = font_bold
        p.font.color.rgb = font_color or self.theme.get("text_primary", RGBColor(15, 23, 42))
        p.font.name = self.theme.get("font_body", "Segoe UI")
        self.report.total_shapes += 1
        self.report.editable_shapes_count += 1
        return tb

    def add_glass_card(
        self,
        slide: Any,
        left_in: float,
        top_in: float,
        width_in: float,
        height_in: float,
        name: Optional[str] = None,
        title: Optional[str] = None,
        body: Optional[str] = None,
        bg_hex: Optional[str] = None,
        border_hex: Optional[str] = None,
        alpha_val: str = "92000",
    ) -> Any:
        """
        Adds an authentic frosted glassmorphism card using OpenXML alpha scrim and hairline border.
        """
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left_in),
            Inches(top_in),
            Inches(width_in),
            Inches(height_in),
        )
        if name:
            card.name = name
            if name.startswith("!!"):
                self.report.named_anchors.append(name)

        card_bg = bg_hex or self.theme.get("card_bg_hex", "FFFFFF")
        card_border = border_hex or self.theme.get("border_hex", "0284C7")

        x_emu = int(left_in * 914400)
        y_emu = int(top_in * 914400)
        cx_emu = int(width_in * 914400)
        cy_emu = int(height_in * 914400)

        spPr_xml = (
            f'<p:spPr xmlns:p="{P_NS}" xmlns:a="{A_NS}">'
            f'<a:xfrm><a:off x="{x_emu}" y="{y_emu}"/><a:ext cx="{cx_emu}" cy="{cy_emu}"/></a:xfrm>'
            f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8000"/></a:avLst></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{card_bg}"><a:alpha val="{alpha_val}"/></a:srgbClr></a:solidFill>'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{card_border}"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:ln>'
            f'</p:spPr>'
        )
        card._element.replace(card._element.spPr, parse_xml(spPr_xml))

        if title or body:
            tf = card.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.TOP
            if title:
                p0 = tf.paragraphs[0]
                p0.text = title
                p0.font.bold = True
                p0.font.size = Pt(16)
                p0.font.color.rgb = self.theme.get("text_primary", RGBColor(15, 23, 42))
                p0.font.name = self.theme.get("font_heading", "Segoe UI")
            if body:
                p1 = tf.add_paragraph() if title else tf.paragraphs[0]
                p1.text = body
                p1.font.size = Pt(12)
                p1.font.color.rgb = self.theme.get("text_secondary", RGBColor(71, 85, 105))
                p1.font.name = self.theme.get("font_body", "Segoe UI")

        self.report.total_shapes += 1
        self.report.editable_shapes_count += 1
        return card

    def add_shape(
        self,
        slide: Any,
        shape_type: MSO_SHAPE,
        left_in: float,
        top_in: float,
        width_in: float,
        height_in: float,
        name: Optional[str] = None,
        fill_rgb: Optional[RGBColor] = None,
        border_rgb: Optional[RGBColor] = None,
        text: Optional[str] = None,
    ) -> Any:
        """Adds a native geometric shape."""
        sh = slide.shapes.add_shape(shape_type, Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in))
        if name:
            sh.name = name
            if name.startswith("!!"):
                self.report.named_anchors.append(name)
        if fill_rgb:
            sh.fill.solid()
            sh.fill.fore_color.rgb = fill_rgb
        else:
            sh.fill.background()

        if border_rgb:
            sh.line.color.rgb = border_rgb
            sh.line.width = Pt(1.0)
        else:
            sh.line.fill.background()

        if text:
            tf = sh.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = text
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(14)
            p.font.color.rgb = self.theme.get("text_primary", RGBColor(15, 23, 42))

        self.report.total_shapes += 1
        self.report.editable_shapes_count += 1
        return sh

    def add_animation(
        self,
        slide: Any,
        shape: Any,
        effect_type: str,
        trigger: str = "with_previous",
        duration_ms: int = 800,
        delay_ms: int = 0,
        direction: str = "up",
        motion_path: Optional[str] = None,
        rotation_deg: float = 0.0,
        scale_factor: Tuple[float, float] = (1.0, 1.0),
        opacity_from: float = 1.0,
        opacity_to: float = 1.0,
    ) -> AnimationStep:
        """Registers a native Animation Pane effect on a shape."""
        slide_idx = self.slides.index(slide) + 1
        timeline = self.timelines[slide_idx]

        cap_status, norm_eff = get_effect_capability_status(effect_type)
        self.report.capability_breakdown[cap_status.value] = (
            self.report.capability_breakdown.get(cap_status.value, 0) + 1
        )
        if cap_status == AnimationCapabilityStatus.UNSUPPORTED:
            self.report.unsupported_effects.append(effect_type)
        elif cap_status == AnimationCapabilityStatus.SUPPORTED_UNVERIFIED_PLAYBACK:
            self.report.unverified_playback_effects.append(effect_type)

        step = timeline.add_effect(
            shape=shape,
            effect_type=norm_eff,
            trigger=trigger,
            duration_ms=duration_ms,
            delay_ms=delay_ms,
            direction=direction,
            motion_path=motion_path,
            rotation_deg=rotation_deg,
            scale_factor=scale_factor,
            opacity_from=opacity_from,
            opacity_to=opacity_to,
        )
        self.report.total_animations += 1
        return step

    def compile_all_timelines(self) -> None:
        """Attaches compiled <p:timing> OpenXML trees to all slides."""
        for idx, slide in enumerate(self.slides, start=1):
            tl = self.timelines.get(idx)
            if tl and tl.steps:
                tl.compile_to_openxml(slide)
                self.report.timing_trees_attached += 1

    def save(self, output_pptx_path: str) -> CompilerReport:
        """Compiles timelines and writes the native PPTX to disk."""
        if self.prs is None:
            raise RuntimeError("No presentation initialized. Call new_presentation() first.")

        self.compile_all_timelines()

        abs_out = os.path.abspath(output_pptx_path)
        os.makedirs(os.path.dirname(abs_out), exist_ok=True)
        self.prs.save(abs_out)
        self.report.output_path = abs_out
        self.report.total_slides = len(self.slides)
        return self.report

    def compile_from_ir(self, ir: AnimationIR, output_path: str) -> CompilerReport:
        """
        Compiles a validated AnimationIR into a native PowerPoint presentation.
        """
        self.report.morph_validation_report = MorphReliabilityValidator.validate_morph_flow(ir)

        theme_map = {
            "LIGHT_ACADEMIC_CYAN": Themes.LIGHT_ACADEMIC_CYAN,
            "LIGHT_MODERN_STUDIO": Themes.LIGHT_MODERN_STUDIO,
            "LIGHT_ECO_EMERALD": getattr(Themes, "LIGHT_ECO_EMERALD", Themes.LIGHT_ACADEMIC_CYAN),
            "DARK_OBSIDIAN": getattr(Themes, "DARK_OBSIDIAN", Themes.LIGHT_ACADEMIC_CYAN),
        }
        self.theme = theme_map.get(ir.palette_name, Themes.LIGHT_ACADEMIC_CYAN)
        self.new_presentation(width_in=ir.canvas_width_in, height_in=ir.canvas_height_in)

        for scene in ir.scenes:
            trans = scene.transition_type if scene.transition_type != "none" else "fade"
            dur_ms = scene.transition_duration_ms
            slide = self.add_slide(transition=trans, duration_ms=dur_ms)

            # Titles
            if scene.title:
                self.add_textbox(
                    slide=slide,
                    left_in=0.8,
                    top_in=0.8,
                    width_in=8.0,
                    height_in=1.0,
                    text=scene.title,
                    name=f"!!Title_{scene.scene_id}",
                    font_size_pt=32,
                    font_bold=True,
                )
            if scene.subtitle:
                self.add_textbox(
                    slide=slide,
                    left_in=0.8,
                    top_in=1.8,
                    width_in=8.0,
                    height_in=0.6,
                    text=scene.subtitle,
                    name=f"!!Subtitle_{scene.scene_id}",
                    font_size_pt=18,
                    font_color=self.theme.get("text_secondary", RGBColor(71, 85, 105)),
                )

            # Objects
            shape_map: Dict[str, Any] = {}
            for obj in scene.objects:
                x, y, w, h = obj.transform.x, obj.transform.y, obj.transform.width, obj.transform.height
                if obj.kind == "card":
                    shp = self.add_glass_card(
                        slide=slide,
                        left_in=x,
                        top_in=y,
                        width_in=w,
                        height_in=h,
                        name=obj.name or obj.id,
                        title=obj.title,
                        body=obj.body,
                    )
                elif obj.kind == "textbox":
                    shp = self.add_textbox(
                        slide=slide,
                        left_in=x,
                        top_in=y,
                        width_in=w,
                        height_in=h,
                        text=obj.body or obj.title,
                        name=obj.name or obj.id,
                    )
                elif obj.kind == "oval":
                    shp = self.add_shape(
                        slide=slide,
                        shape_type=MSO_SHAPE.OVAL,
                        left_in=x,
                        top_in=y,
                        width_in=w,
                        height_in=h,
                        name=obj.name or obj.id,
                        fill_rgb=self.theme.get("accent"),
                    )
                else:
                    shp = self.add_shape(
                        slide=slide,
                        shape_type=MSO_SHAPE.ROUNDED_RECTANGLE,
                        left_in=x,
                        top_in=y,
                        width_in=w,
                        height_in=h,
                        name=obj.name or obj.id,
                        fill_rgb=self.theme.get("accent"),
                        text=obj.title,
                    )
                shape_map[obj.id] = shp
                shape_map[obj.name] = shp

            # Timeline compile
            timeline_compiler = TimelineCompiler(scene)
            scheduled_tracks = timeline_compiler.compile()

            for track in scheduled_tracks:
                shp = shape_map.get(track.target_id)
                if shp:
                    self.add_animation(
                        slide=slide,
                        shape=shp,
                        effect_type=track.effect,
                        trigger=track.trigger,
                        duration_ms=track.duration_ms,
                        delay_ms=track.delay_ms,
                        direction=track.direction or "up",
                        motion_path=track.motion_path,
                        rotation_deg=track.rotation_deg,
                        scale_factor=track.scale_factor,
                        opacity_from=track.opacity_from,
                        opacity_to=track.opacity_to,
                    )

        return self.save(output_path)

    def compile_from_spec(self, spec: Union[str, Dict[str, Any], AnimationIR], output_path: str) -> CompilerReport:
        """
        Compiles a complete native presentation from a JSON string, file path, dictionary, or AnimationIR.
        """
        if isinstance(spec, AnimationIR):
            return self.compile_from_ir(spec, output_path)

        if isinstance(spec, str):
            if os.path.isfile(spec):
                with open(spec, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = json.loads(spec)
        else:
            data = spec

        if isinstance(data, dict) and "scenes" in data:
            ir = AnimationIR.from_dict(data)
            return self.compile_from_ir(ir, output_path)

        theme_name = data.get("theme", "LIGHT_ACADEMIC_CYAN")
        theme_map = {
            "LIGHT_ACADEMIC_CYAN": Themes.LIGHT_ACADEMIC_CYAN,
            "LIGHT_MODERN_STUDIO": Themes.LIGHT_MODERN_STUDIO,
            "LIGHT_ECO_EMERALD": getattr(Themes, "LIGHT_ECO_EMERALD", Themes.LIGHT_ACADEMIC_CYAN),
            "DARK_OBSIDIAN": getattr(Themes, "DARK_OBSIDIAN", Themes.LIGHT_ACADEMIC_CYAN),
        }
        self.theme = theme_map.get(theme_name, Themes.LIGHT_ACADEMIC_CYAN)
        self.new_presentation(
            width_in=data.get("width_in", 13.333),
            height_in=data.get("height_in", 7.5),
        )

        slides_spec = data.get("slides", [])
        for s_data in slides_spec:
            trans = s_data.get("transition", "morph")
            dur_ms = s_data.get("duration_ms", 1400)
            bg_hex = s_data.get("bg_hex")
            slide = self.add_slide(transition=trans, duration_ms=dur_ms, bg_color_hex=bg_hex)

            # Slide Title
            s_title = s_data.get("title")
            if s_title:
                self.add_textbox(
                    slide=slide,
                    left_in=0.8,
                    top_in=0.8,
                    width_in=8.0,
                    height_in=1.0,
                    text=s_title,
                    name=s_data.get("title_name", "!!SlideTitle"),
                    font_size_pt=32,
                    font_bold=True,
                )

            # Subtitle
            s_sub = s_data.get("subtitle")
            if s_sub:
                self.add_textbox(
                    slide=slide,
                    left_in=0.8,
                    top_in=1.8,
                    width_in=8.0,
                    height_in=0.6,
                    text=s_sub,
                    name=s_data.get("subtitle_name", "!!SlideSubtitle"),
                    font_size_pt=18,
                    font_color=self.theme.get("text_secondary", RGBColor(71, 85, 105)),
                )

            # Elements
            elements_map: Dict[str, Any] = {}
            for elem in s_data.get("elements", []):
                e_type = elem.get("type", "card")
                e_name = elem.get("name", "Shape")
                x = elem.get("left", 1.0)
                y = elem.get("top", 2.5)
                w = elem.get("width", 3.5)
                h = elem.get("height", 2.2)

                if e_type == "card":
                    shape_obj = self.add_glass_card(
                        slide=slide,
                        left_in=x,
                        top_in=y,
                        width_in=w,
                        height_in=h,
                        name=e_name,
                        title=elem.get("title"),
                        body=elem.get("body"),
                        bg_hex=elem.get("bg_hex"),
                        border_hex=elem.get("border_hex"),
                    )
                elif e_type == "text":
                    shape_obj = self.add_textbox(
                        slide=slide,
                        left_in=x,
                        top_in=y,
                        width_in=w,
                        height_in=h,
                        text=elem.get("text", ""),
                        name=e_name,
                        font_size_pt=elem.get("font_size_pt", 16),
                        font_bold=elem.get("font_bold", False),
                    )
                elif e_type == "oval":
                    shape_obj = self.add_shape(
                        slide=slide,
                        shape_type=MSO_SHAPE.OVAL,
                        left_in=x,
                        top_in=y,
                        width_in=w,
                        height_in=h,
                        name=e_name,
                        fill_rgb=self.theme.get("accent"),
                        text=elem.get("text"),
                    )
                else:  # rounded rectangle default
                    shape_obj = self.add_shape(
                        slide=slide,
                        shape_type=MSO_SHAPE.ROUNDED_RECTANGLE,
                        left_in=x,
                        top_in=y,
                        width_in=w,
                        height_in=h,
                        name=e_name,
                        fill_rgb=self.theme.get("accent"),
                        text=elem.get("text"),
                    )
                elements_map[e_name] = shape_obj

            # Animations
            for anim in s_data.get("animations", []):
                tgt_name = anim.get("target")
                shape_ref = elements_map.get(tgt_name)
                if not shape_ref:
                    # Find by name on slide
                    for s in slide.shapes:
                        if s.name == tgt_name:
                            shape_ref = s
                            break
                if shape_ref:
                    eff = anim.get("effect", "fade")
                    if eff not in SUPPORTED_EFFECTS and eff not in ("fade_in", "zoom_in", "spin", "grow_shrink"):
                        self.report.unsupported_effects.append(eff)
                    self.add_animation(
                        slide=slide,
                        shape=shape_ref,
                        effect_type=eff,
                        trigger=anim.get("trigger", "with_previous"),
                        duration_ms=anim.get("duration_ms", 800),
                        delay_ms=anim.get("delay_ms", 0),
                        direction=anim.get("direction", "up"),
                        motion_path=anim.get("motion_path"),
                        rotation_deg=anim.get("rotation_deg", 0.0),
                        scale_factor=tuple(anim.get("scale_factor", (1.0, 1.0))),
                    )

        return self.save(output_path)

