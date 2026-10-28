"""
pmg.cinematic_builder
=====================
Professional Cinematic Motion Graphics Builder (v1.3.0).

Combines cross-slide Morph storytelling, multi-plane parallax (BG 0.7x, Mid 1.0x, FG 1.4x),
virtual camera tracking/pan/zoom simulation, animated infographics, 3D orbiting carousels,
animated interface mockups, and reusable motion presets with intra-slide native PowerPoint
Animation Pane timelines (pmg.animations.SlideAnimationTimeline).
"""

import os
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional, Tuple

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

from pmg.engine import (
    MotionDeck,
    Themes,
    add_editable_earth,
    add_editable_meadow_landscape,
    add_editable_kite,
)
from pmg.animations import SlideAnimationTimeline


@dataclass
class MotionChoreographyDecision:
    """
    Explicit narrative justification and parameters for a motion sequence, ensuring
    every significant animation serves a visual or storytelling purpose.
    """
    slide_number: int
    object_id: str
    scope: str               # "inter_slide_morph" or "intra_slide_timeline"
    motion_preset: str       # e.g., "camera_zoom_inspection", "animated_infographic_bars", "multi_plane_parallax"
    narrative_purpose: str   # Why the object moves
    trajectory_desc: str     # Where it moves
    duration_ms: int         # How long it moves
    execution_mode: str      # "simultaneous" or "sequential"

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class MotionPresetRegistry:
    """Catalog of reusable purposeful motion presets for v1.3+ presentations."""

    PRESETS: Dict[str, Dict[str, Any]] = {
        "cinematic_title_reveal": {
            "scope": "intra_slide_timeline",
            "purpose": "Establish thematic hierarchy with simultaneous badge appearance and smooth title fade-in.",
            "effects": ["appear", "fade", "zoom"],
        },
        "multi_plane_parallax": {
            "scope": "inter_slide_morph",
            "purpose": "Simulate 3D depth across background (0.7x), midground (1.0x), and foreground (1.4x) terrain planes.",
            "speed_ratios": {"bg": 0.7, "mid": 1.0, "fg": 1.4},
        },
        "camera_zoom_inspection": {
            "scope": "inter_slide_morph+intra_slide_timeline",
            "purpose": "Push virtual camera from wide establishing shot (1.0x) into macro inspection (1.8x-2.4x) with HUD callouts.",
            "effects": ["morph_scale", "wipe", "scale"],
        },
        "animated_infographic_bars": {
            "scope": "intra_slide_timeline",
            "purpose": "Quantify environmental metrics through staggered bar growth, curved trend trajectory, and target emphasis.",
            "effects": ["wipe", "fly_in", "motion_path", "scale"],
        },
        "orbiting_carousel_3d": {
            "scope": "inter_slide_morph+intra_slide_timeline",
            "purpose": "Rotate focal cards through foreground center stage while flanking nodes recede in scale.",
            "effects": ["morph_orbit", "fade", "scale"],
        },
        "animated_interface_mockup": {
            "scope": "intra_slide_timeline",
            "purpose": "Reveal telemetry dashboard shell and live sensor widgets in structured operational sequence.",
            "effects": ["fade", "fly_in", "wipe", "rotate"],
        },
    }


class CinematicMotionBuilder:
    """
    High-level orchestrator that constructs multi-layered scenes combining
    inter-slide Morph/Transitions with intra-slide native PowerPoint Animation Pane timelines.
    """

    def __init__(self, theme: Dict[str, Any] = None):
        self.deck = MotionDeck(theme=theme or Themes.LIGHT_ECO_EMERALD)
        self.choreography_log: List[MotionChoreographyDecision] = []
        self.slide_timelines: Dict[int, SlideAnimationTimeline] = {}

    def record_decision(
        self,
        slide_number: int,
        object_id: str,
        scope: str,
        motion_preset: str,
        narrative_purpose: str,
        trajectory_desc: str,
        duration_ms: int,
        execution_mode: str,
    ) -> None:
        self.choreography_log.append(
            MotionChoreographyDecision(
                slide_number=slide_number,
                object_id=object_id,
                scope=scope,
                motion_preset=motion_preset,
                narrative_purpose=narrative_purpose,
                trajectory_desc=trajectory_desc,
                duration_ms=duration_ms,
                execution_mode=execution_mode,
            )
        )

    def add_animated_infographic(
        self,
        slide,
        slide_number: int,
        timeline: SlideAnimationTimeline,
        bars_data: List[Tuple[str, str, float, RGBColor]],
        left: float = 0.9,
        top: float = 2.85,
        width: float = 5.6,
        height: float = 3.3,
        name_prefix: str = "Infographic",
    ) -> List[Any]:
        """
        Adds a native editable multi-bar animated infographic with:
        - Frosted container card
        - 4 independent proportional native bar shapes + label badges
        - Animated trend node moving along a curved Bezier motion path across the bars
        """
        created = []
        container = self.deck.add_glass_card(
            slide,
            left=left,
            top=top,
            width=width,
            height=height,
            name=f"{name_prefix}_Container",
        )
        created.append(container)
        timeline.add_effect(container, "fade", trigger="with_previous", duration_ms=300)

        num_bars = len(bars_data)
        row_h = (height - 0.55) / max(1, num_bars)
        max_bar_w = width - 2.35

        for b_idx, (label, val_str, pct, color) in enumerate(bars_data):
            r_top = top + 0.35 + b_idx * row_h
            # Label pill
            lbl_shp = self.deck.add_badge(
                slide,
                f"{label}: {val_str}",
                left=left + 0.2,
                top=r_top,
                width=1.85,
                height=0.36,
                bg_color=color,
                text_color=RGBColor(255, 255, 255),
                name=f"{name_prefix}_Label_{b_idx+1}",
            )
            created.append(lbl_shp)

            # Proportional data bar
            bar_w = max(0.4, max_bar_w * max(0.1, min(1.0, pct)))
            bar_shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left + 2.15),
                Inches(r_top + 0.03),
                Inches(bar_w),
                Inches(0.30),
            )
            bar_shp.name = f"{name_prefix}_Bar_{b_idx+1}"
            bar_shp.fill.solid()
            bar_shp.fill.fore_color.rgb = color
            bar_shp.line.color.rgb = RGBColor(255, 255, 255)
            bar_shp.line.width = Pt(1.0)
            created.append(bar_shp)

            trig = "after_previous" if b_idx == 0 else "with_previous"
            d_ms = b_idx * 70
            timeline.add_effect(lbl_shp, "appear", trigger=trig, duration_ms=150, delay_ms=d_ms)
            timeline.add_effect(bar_shp, "wipe", trigger="with_previous", direction="right", duration_ms=350, delay_ms=d_ms)

        # Curved trajectory indicator node across infographic
        trend_orb = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(left + 2.2),
            Inches(top + height - 0.55),
            Inches(0.34),
            Inches(0.34),
        )
        trend_orb.name = f"{name_prefix}_Trend_Node"
        trend_orb.fill.solid()
        trend_orb.fill.fore_color.rgb = RGBColor(250, 204, 21)
        trend_orb.line.color.rgb = RGBColor(6, 78, 59)
        trend_orb.line.width = Pt(1.5)
        created.append(trend_orb)

        timeline.add_effect(trend_orb, "appear", trigger="after_previous", duration_ms=100)
        timeline.add_effect(
            trend_orb,
            "motion_path",
            trigger="with_previous",
            motion_path="M 0 0 C 0.06 -0.10 0.14 -0.16 0.22 -0.22 E",
            duration_ms=450,
        )

        self.record_decision(
            slide_number=slide_number,
            object_id=f"{name_prefix}_Bar_1..{num_bars}",
            scope="intra_slide_timeline",
            motion_preset="animated_infographic_bars",
            narrative_purpose="Visually quantify ecological restoration metrics via proportional bar wipes and upward curved trend trajectory.",
            trajectory_desc="Horizontal bar wipes (left-to-right) + upward Bezier trend arc.",
            duration_ms=900,
            execution_mode="simultaneous+sequential",
        )
        return created

    def add_orbiting_carousel(
        self,
        slide,
        slide_number: int,
        timeline: SlideAnimationTimeline,
        cards_spec: List[Tuple[str, str, str, RGBColor]],
        active_index: int = 1,
        left: float = 0.85,
        top: float = 2.9,
    ) -> List[Any]:
        """
        Adds a 3-card 3D-style orbiting carousel where the active card sits prominent
        in the foreground and flanking satellite cards sit in recessed depth planes.
        """
        created = []
        positions = [
            (left, top + 0.35, 1.65, 1.35),          # Left satellite (recessed)
            (left + 1.80, top, 2.25, 1.85),          # Center hero stage (foreground)
            (left + 4.20, top + 0.35, 1.65, 1.35),   # Right satellite (recessed)
        ]
        for i, (title, val, desc, col) in enumerate(cards_spec[:3]):
            slot_idx = (i - active_index + 1) % 3
            cx, cy, cw, ch = positions[slot_idx]
            card, tb = self.deck.add_metric_card(
                slide,
                title=title,
                value=val,
                desc=desc,
                left=cx,
                top=cy,
                width=cw,
                height=ch,
                accent_color=col,
                name=f"Carousel_Node_{i+1}",
            )
            created.append(card)
            timeline.add_effect(
                card,
                "fly_in" if slot_idx == 1 else "fade",
                trigger="with_previous",
                direction="up",
                duration_ms=350,
                delay_ms=i * 60,
            )
            if slot_idx == 1:
                timeline.add_effect(
                    card,
                    "scale",
                    trigger="after_previous",
                    scale_factor=(1.08, 1.08),
                    duration_ms=300,
                    auto_reverse=True,
                )

        self.record_decision(
            slide_number=slide_number,
            object_id="Carousel_Node_1..3",
            scope="intra_slide_timeline",
            motion_preset="orbiting_carousel_3d",
            narrative_purpose="Spotlight primary circular-economy pillar in foreground stage while flanking pillars orbit in periphery.",
            trajectory_desc="3-slot spatial carousel layout with foreground scale emphasis.",
            duration_ms=650,
            execution_mode="simultaneous+sequential",
        )
        return created

    def add_interface_mockup(
        self,
        slide,
        slide_number: int,
        timeline: SlideAnimationTimeline,
        left: float = 0.9,
        top: float = 2.85,
        width: float = 5.6,
        height: float = 3.25,
    ) -> List[Any]:
        """
        Adds an editable environmental satellite telemetry UI mockup shell with
        header bar, live sensor cards, and rotating radar indicator.
        """
        created = []
        shell = self.deck.add_glass_card(
            slide, left=left, top=top, width=width, height=height, name="UI_Mockup_Shell"
        )
        created.append(shell)
        hdr = self.deck.add_badge(
            slide,
            "EARTH OBSERVATION SATELLITE HUD · LIVE BIOSPHERE TELEMETRY",
            left=left + 0.18,
            top=top + 0.18,
            width=width - 0.36,
            height=0.36,
            bg_color=RGBColor(2, 132, 199),
            text_color=RGBColor(255, 255, 255),
            name="UI_Mockup_Header",
        )
        created.append(hdr)

        w1, _ = self.deck.add_metric_card(
            slide,
            title="FOREST CANOPY",
            value="+18.4%",
            desc=" LIDAR biomass index",
            left=left + 0.2,
            top=top + 0.70,
            width=2.5,
            height=1.6,
            accent_color=RGBColor(5, 150, 105),
            name="UI_Widget_Canopy",
        )
        w2, _ = self.deck.add_metric_card(
            slide,
            title="OCEAN PH",
            value="8.12 pH",
            desc="Coral reef buffer nominal",
            left=left + 2.9,
            top=top + 0.70,
            width=2.5,
            height=1.6,
            accent_color=RGBColor(2, 132, 199),
            name="UI_Widget_Ocean",
        )
        radar = slide.shapes.add_shape(
            MSO_SHAPE.STAR_4_POINT,
            Inches(left + width - 0.85),
            Inches(top + height - 0.75),
            Inches(0.55),
            Inches(0.55),
        )
        radar.name = "UI_Radar_Spinner"
        radar.fill.solid()
        radar.fill.fore_color.rgb = RGBColor(16, 185, 129)
        radar.line.color.rgb = RGBColor(255, 255, 255)
        created.extend([w1, w2, radar])

        timeline.add_effect(shell, "fade", trigger="with_previous", duration_ms=250)
        timeline.add_effect(hdr, "wipe", trigger="with_previous", direction="right", duration_ms=300)
        timeline.add_effect(w1, "fly_in", trigger="after_previous", direction="up", duration_ms=300)
        timeline.add_effect(w2, "fly_in", trigger="with_previous", direction="up", duration_ms=300, delay_ms=70)
        timeline.add_effect(radar, "rotate", trigger="with_previous", rotation_deg=360.0, duration_ms=500)

        self.record_decision(
            slide_number=slide_number,
            object_id="UI_Mockup_Shell",
            scope="intra_slide_timeline",
            motion_preset="animated_interface_mockup",
            narrative_purpose="Present real-time planetary monitoring HUD with progressive widget unveiling and active radar rotation.",
            trajectory_desc="Shell fade -> Header wipe -> Dual sensor fly-in -> Radar 360 spin.",
            duration_ms=850,
            execution_mode="simultaneous+sequential",
        )
        return created


def build_v1_3_cinematic_environmental_deck(output_path: str, num_slides: int = 10) -> str:
    """
    Canonical v1.3.0 Acceptance Presentation Builder:
    Generates a professional 10-slide presentation about protecting the environment combining:
    - Earth (multi-layered native globe with camera zoom sequences)
    - Meadow & Green landscapes (multi-plane parallax at 0.7x, 1.0x, 1.4x)
    - Blue sky & Sun
    - Flying Kite recurring traveling entity (!!Traveling_Kite) guiding the viewer across all 10 slides
    - Environmental messages
    - Animated infographics (Slides 4 & 8)
    - Camera zoom sequences (Slides 2->3 macro push-in & Slides 6->7 extreme macro dolly zoom)
    - Multi-plane parallax sequences across all 10 slides
    - 3D orbiting carousel (Slide 5) & Animated telemetry UI mockup (Slide 6)
    - Both cross-slide Morph AND intra-slide native PowerPoint Animation Pane effects on every slide!
    """
    builder = CinematicMotionBuilder(theme=Themes.LIGHT_ECO_EMERALD)
    deck = builder.deck

    scenes = [
        {
            "badge": "STAGE 01 · LIVING BIOSPHERE",
            "title": "Protecting Our Living Earth",
            "subtitle": "A guided cinematic journey across skies, green meadows, and living oceans.",
            "mode": "metrics",
            "m1": ("CARBON CAPTURE", "42 GT/YR", "Global forests and seagrass meadows absorb emissions."),
            "m2": ("CLEAN AIR INDEX", "98.4%", "Zero-emission corridors restore atmospheric clarity."),
            "earth": (8.6, 0.9, 2.2, 0.0),
            "kite": (5.9, 0.65, 1.05, 8.0),
            "parallax": 0.0,
            "transition": "none",
        },
        {
            "badge": "STAGE 02 · MULTI-PLANE PARALLAX",
            "title": "Verdant Meadows & Winds",
            "subtitle": "Three depth planes glide at 0.7x, 1.0x, and 1.4x as the kite rides the Jetstream.",
            "mode": "metrics",
            "m1": ("GRASSLAND BIOMASS", "+34%", "Regenerative grazing rebuilds deep prairie root networks."),
            "m2": ("POLLINATOR FLIGHT", "2.4X", "Wildflower corridors reconnect fragmented habitats."),
            "earth": (9.2, 0.75, 1.9, 25.0),
            "kite": (6.6, 0.95, 1.15, -10.0),
            "parallax": 0.85,
            "transition": "morph",
        },
        {
            "badge": "STAGE 03 · CAMERA ZOOM SEQUENCE",
            "title": "Macro Push-In: Ocean Currents",
            "subtitle": "Virtual camera zooms 175% into Earth's marine biosphere as the kite dives.",
            "mode": "metrics",
            "m1": ("KELP FORESTS", "18M HA", "Underwater canopies sequester carbon 20x faster than land."),
            "m2": ("REEF RESILIENCE", "+62%", "Thermal-tolerant coral nurseries restore coastal barriers."),
            "earth": (7.55, 0.45, 3.35, 55.0),
            "kite": (10.9, 3.25, 0.95, 16.0),
            "parallax": 1.65,
            "transition": "morph",
        },
        {
            "badge": "STAGE 04 · ANIMATED INFOGRAPHIC",
            "title": "Clean Energy Transition Metrics",
            "subtitle": "Real-time renewable capacity growth across solar, wind, hydro, and storage.",
            "mode": "infographic_1",
            "earth": (9.3, 1.1, 2.1, 85.0),
            "kite": (6.8, 0.65, 1.1, -6.0),
            "parallax": 2.4,
            "transition": "morph",
        },
        {
            "badge": "STAGE 05 · 3D ORBITING CAROUSEL",
            "title": "Circular Economy Pillars",
            "subtitle": "Three interconnected conservation loops orbiting our zero-waste industrial ecosystem.",
            "mode": "carousel",
            "earth": (8.9, 1.35, 2.35, 115.0),
            "kite": (6.95, 0.55, 1.0, 12.0),
            "parallax": 1.7,
            "transition": "morph",
        },
        {
            "badge": "STAGE 06 · SATELLITE HUD MOCKUP",
            "title": "Planetary Telemetry Interface",
            "subtitle": "Orbital sensors track canopy recovery, ocean pH, and atmospheric purity in real time.",
            "mode": "mockup",
            "earth": (8.7, 0.95, 2.5, 145.0),
            "kite": (6.8, 1.95, 0.95, -14.0),
            "parallax": 0.95,
            "transition": "morph",
        },
        {
            "badge": "STAGE 07 · EXTREME CAMERA ZOOM",
            "title": "Continental Reforestation",
            "subtitle": "Camera pushes 200% into the emerald continental belt to inspect canopy restoration.",
            "mode": "metrics",
            "m1": ("TREES PLANTED", "1.2 TRILLION", "Native biodiversity belts halt desertification."),
            "m2": ("SOIL MOISTURE", "+41%", "Mycorrhizal networks retain monsoon rainfall."),
            "earth": (7.35, 0.35, 3.75, 180.0),
            "kite": (6.55, 0.55, 0.85, 6.0),
            "parallax": 0.3,
            "transition": "morph",
        },
        {
            "badge": "STAGE 08 · BIODIVERSITY INFOGRAPHIC",
            "title": "Species & Habitat Recovery",
            "subtitle": "Decadal restoration trajectory across wetlands, alpine forests, and coral reefs.",
            "mode": "infographic_2",
            "earth": (9.1, 0.85, 2.25, 215.0),
            "kite": (6.8, 0.65, 1.15, -12.0),
            "parallax": 1.15,
            "transition": "morph",
        },
        {
            "badge": "STAGE 09 · GLOBAL STEWARDSHIP",
            "title": "Communities United for Nature",
            "subtitle": "Local conservationists, clean-tech engineers, and youth movements leading action.",
            "mode": "metrics",
            "m1": ("ECO CITIES", "4,500+", "Urban sponge parks and 100% electric transit."),
            "m2": ("YOUTH STEWARDS", "180M", "Next-generation climate guardians across 195 nations."),
            "earth": (8.6, 1.05, 2.55, 250.0),
            "kite": (6.7, 1.35, 1.1, 10.0),
            "parallax": 1.95,
            "transition": "morph",
        },
        {
            "badge": "STAGE 10 · HORIZON 2050 FINALE",
            "title": "A Thriving Blue & Green Planet",
            "subtitle": "The kite ascends into a clear azure sky above restored meadows and a vibrant Earth.",
            "mode": "metrics",
            "m1": ("NET EMISSIONS", "0.0 GT", "Complete climate equilibrium achieved globally."),
            "m2": ("LIVING EARTH", "100% PROTECTED", "Harmony between humanity and nature secured."),
            "earth": (8.4, 0.85, 2.8, 290.0),
            "kite": (10.5, 0.45, 1.0, 20.0),
            "parallax": 0.6,
            "transition": "circle",
        },
    ]

    selected_scenes = scenes[: max(1, min(num_slides, len(scenes)))]
    total = len(selected_scenes)

    for idx, sc in enumerate(selected_scenes, 1):
        trans = sc["transition"] if idx > 1 else "none"
        slide = deck.add_slide(transition=trans, duration_ms=1400)
        tl = SlideAnimationTimeline(slide_number=idx)
        builder.slide_timelines[idx] = tl

        # 1. Multi-Plane Blue Sky & Green Meadow Landscape (0.7x BG, 1.0x Mid, 1.4x FG parallax)
        add_editable_meadow_landscape(slide, pan_x=sc["parallax"], sun_pos=(11.2 - idx * 0.15, 0.7))
        builder.record_decision(
            slide_number=idx,
            object_id="!!Meadow_Back..Front",
            scope="inter_slide_morph",
            motion_preset="multi_plane_parallax",
            narrative_purpose="Create continuous 3D landscape depth across blue sky and green meadows.",
            trajectory_desc=f"Parallax offset={sc['parallax']} across 3 terrain planes.",
            duration_ms=1400,
            execution_mode="simultaneous",
        )

        # 2. Multi-Layered Editable Earth Globe (with camera zoom sequences on Slides 3 & 7)
        ex, ey, edia, erot = sc["earth"]
        earth_shapes = add_editable_earth(slide, left=ex, top=ey, diameter=edia, rot_deg=erot, name_prefix="!!Earth")

        # 3. Header Badge & Hero Title (with intra-slide native entrance animations)
        badge_shp = deck.add_badge(
            slide,
            sc["badge"],
            left=0.9,
            top=0.58,
            width=3.6,
            height=0.38,
            bg_color=RGBColor(5, 150, 105),
            text_color=RGBColor(255, 255, 255),
            name="!!env_badge",
        )
        title_shp = deck.add_hero_title(
            slide,
            sc["title"],
            subtitle=sc["subtitle"],
            left=0.9,
            top=1.10,
            width=6.1,
            height=1.55,
            title_size=23,
            subtitle_size=11.0,
            title_color=RGBColor(6, 78, 59),
            name="!!env_title",
        )
        tl.add_effect(badge_shp, "appear", trigger="with_previous", duration_ms=150)
        tl.add_effect(title_shp, "fade", trigger="with_previous", duration_ms=300, delay_ms=50)

        # 4. Mode-specific foreground choreography (Infographic, Carousel, Mockup, or Metric Cards)
        mode = sc["mode"]
        if mode == "infographic_1":
            builder.add_animated_infographic(
                slide,
                slide_number=idx,
                timeline=tl,
                bars_data=[
                    ("SOLAR", "94%", 0.94, RGBColor(245, 158, 11)),
                    ("WIND", "88%", 0.88, RGBColor(2, 132, 199)),
                    ("HYDRO", "79%", 0.79, RGBColor(14, 165, 233)),
                    ("STORAGE", "91%", 0.91, RGBColor(5, 150, 105)),
                ],
                name_prefix="CleanEnergy_Info",
            )
        elif mode == "infographic_2":
            builder.add_animated_infographic(
                slide,
                slide_number=idx,
                timeline=tl,
                bars_data=[
                    ("WETLANDS", "+82%", 0.82, RGBColor(16, 185, 129)),
                    ("FORESTS", "+91%", 0.91, RGBColor(5, 150, 105)),
                    ("CORAL REEFS", "+74%", 0.74, RGBColor(2, 132, 199)),
                    ("GRASSLANDS", "+89%", 0.89, RGBColor(101, 163, 13)),
                ],
                name_prefix="Biodiversity_Info",
            )
        elif mode == "carousel":
            builder.add_orbiting_carousel(
                slide,
                slide_number=idx,
                timeline=tl,
                cards_spec=[
                    ("REUSE", "92%", "Closed-loop material recovery.", RGBColor(2, 132, 199)),
                    ("RENEW", "100%", "Clean biological & solar cycles.", RGBColor(5, 150, 105)),
                    ("RESTORE", "+68%", "Wild habitat regeneration.", RGBColor(16, 185, 129)),
                ],
                active_index=1,
            )
        elif mode == "mockup":
            builder.add_interface_mockup(slide, slide_number=idx, timeline=tl)
        else:
            t1, v1, d1 = sc["m1"]
            c1, _ = deck.add_metric_card(
                slide,
                title=t1,
                value=v1,
                desc=d1,
                left=0.9,
                top=2.85,
                width=5.6,
                height=1.55,
                accent_color=RGBColor(5, 150, 105),
                name="!!env_card_1",
            )
            t2, v2, d2 = sc["m2"]
            c2, _ = deck.add_metric_card(
                slide,
                title=t2,
                value=v2,
                desc=d2,
                left=0.9,
                top=4.60,
                width=5.6,
                height=1.55,
                accent_color=RGBColor(2, 132, 199),
                name="!!env_card_2",
            )
            tl.add_effect(c1, "fly_in", trigger="with_previous", direction="up", duration_ms=300, delay_ms=80)
            tl.add_effect(c2, "fly_in", trigger="with_previous", direction="up", duration_ms=300, delay_ms=140)
            # On camera zoom slides (3 & 7), add a subtle scale pulse on the Earth atmosphere halo
            if idx in (3, 7) and earth_shapes:
                tl.add_effect(
                    earth_shapes[0],
                    "scale",
                    trigger="with_previous",
                    scale_factor=(1.12, 1.12),
                    duration_ms=350,
                    auto_reverse=True,
                )

        # 5. Flying Kite Recurring Traveling Entity (soaring across all 10 slides)
        kx, ky, ksz, krot = sc["kite"]
        kite_shapes = add_editable_kite(
            slide, left=kx, top=ky, size=ksz, rot_deg=krot, name_prefix="!!Traveling_Kite"
        )
        # Add subtle intra-slide tail flutter rotation on the kite tail ribbon
        if len(kite_shapes) >= 4:
            tl.add_effect(
                kite_shapes[-1],
                "rotate",
                trigger="with_previous",
                rotation_deg=15.0,
                duration_ms=300,
                auto_reverse=True,
            )

        # 6. Bottom Progress Scrubber Bar
        deck.add_progress_bar(slide, current_step=idx, total_steps=total, top=6.95)

        # Compile intra-slide native animations into <p:timing>
        tl.compile_to_openxml(slide)

    deck.save(output_path)
    return os.path.abspath(output_path)
