"""
pmg.engine
==========
Canonical PMG Presentation Engine (v1.1.0+).
Generates presentations using 100% native, individually editable PowerPoint shapes,
text frames, frosted glassmorphism cards, multi-plane parallax layers, and Morph anchors.
"""

import os
import json
from typing import Optional, Dict, Any, List, Tuple

import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml

from scripts.motion_graph_kit import (
    Themes,
    CinematicColors,
    MotionDeck,
    LightThemeCraftDeck,
    apply_transition,
    apply_morph,
    link_morph,
    export_deck_to_video,
)
from scripts.forge_lib import (
    IndustrialPalette,
    ForgeDeck,
    Cinematic3DDeck,
)


def _set_rot(shape, rot_deg: float):
    if rot_deg != 0:
        try:
            shape._element.spPr.xfrm.set("rot", str(int(rot_deg * 60000)))
        except Exception:
            pass


def add_editable_earth(
    slide,
    left: float,
    top: float,
    diameter: float,
    rot_deg: float = 0.0,
    name_prefix: str = "!!Earth",
) -> List[Any]:
    """
    Constructs a multi-layered, individually editable Earth globe using native PowerPoint shapes:
    - Outer Atmosphere Glow Oval
    - Deep Ocean Core Sphere
    - Continents (Green landmass shapes that shift/rotate across slides)
    - Polar/Cloud Layer
    """
    shapes = []
    # 1. Outer Atmosphere Halo
    pad = diameter * 0.08
    atmo = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left - pad),
        Inches(top - pad),
        Inches(diameter + 2 * pad),
        Inches(diameter + 2 * pad),
    )
    atmo.name = f"{name_prefix}_Atmosphere"
    atmo_xml = (
        f'<p:spPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:xfrm><a:off x="{int((left - pad)*914400)}" y="{int((top - pad)*914400)}"/>'
        f'<a:ext cx="{int((diameter + 2*pad)*914400)}" cy="{int((diameter + 2*pad)*914400)}"/></a:xfrm>'
        f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="38BDF8"><a:alpha val="25000"/></a:srgbClr></a:solidFill>'
        f'<a:ln w="12700"><a:solidFill><a:srgbClr val="0284C7"><a:alpha val="45000"/></a:srgbClr></a:solidFill></a:ln>'
        f'</p:spPr>'
    )
    atmo._element.replace(atmo._element.spPr, parse_xml(atmo_xml))
    shapes.append(atmo)

    # 2. Ocean Sphere
    ocean = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left),
        Inches(top),
        Inches(diameter),
        Inches(diameter),
    )
    ocean.name = f"{name_prefix}_Globe"
    ocean.fill.solid()
    ocean.fill.fore_color.rgb = RGBColor(2, 132, 199)  # Vibrant Ocean Blue
    ocean.line.color.rgb = RGBColor(186, 230, 253)
    ocean.line.width = Pt(1.5)
    _set_rot(ocean, rot_deg)
    shapes.append(ocean)

    # 3. Continent 1 (Eurasia/Americas stylized native landmass)
    shift_x = ((rot_deg % 90) / 90.0) * (diameter * 0.12)
    c1 = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left + diameter * 0.18 + shift_x),
        Inches(top + diameter * 0.22),
        Inches(diameter * 0.36),
        Inches(diameter * 0.26),
    )
    c1.name = f"{name_prefix}_Continent_1"
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(16, 185, 129)  # Emerald Green
    c1.line.fill.background()
    _set_rot(c1, rot_deg * 0.5 - 12)
    shapes.append(c1)

    # 4. Continent 2 (Southern green continent)
    c2 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left + diameter * 0.44 - shift_x * 0.6),
        Inches(top + diameter * 0.52),
        Inches(diameter * 0.34),
        Inches(diameter * 0.28),
    )
    c2.name = f"{name_prefix}_Continent_2"
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor(5, 150, 105)  # Deep Forest Green
    c2.line.fill.background()
    _set_rot(c2, rot_deg * 0.4 + 15)
    shapes.append(c2)

    # 5. Cloud band
    cloud = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left + diameter * 0.25),
        Inches(top + diameter * 0.38),
        Inches(diameter * 0.45),
        Inches(diameter * 0.09),
    )
    cloud.name = f"{name_prefix}_CloudBand"
    cloud.fill.solid()
    cloud.fill.fore_color.rgb = RGBColor(240, 249, 255)
    cloud.line.fill.background()
    _set_rot(cloud, -8)
    shapes.append(cloud)

    return shapes


def add_editable_meadow_landscape(
    slide,
    pan_x: float = 0.0,
    zoom: float = 1.0,
    sun_pos: Tuple[float, float] = (10.8, 0.8),
) -> List[Any]:
    """
    Constructs a multi-plane nature & grassland/meadow landscape using native editable shapes:
    - Sky backdrop band (!!Nature_Sky)
    - Sun orb (!!Nature_Sun)
    - Distant mountain/hill ridge (0.7x parallax: !!Meadow_Back)
    - Foreground lush grassland meadow (1.0x parallax: !!Meadow_Front)
    """
    shapes = []

    # 1. Sky Horizon Band
    sky = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.0),
        Inches(0.0),
        Inches(13.333),
        Inches(7.5),
    )
    sky.name = "!!Nature_Sky"
    sky.fill.solid()
    sky.fill.fore_color.rgb = RGBColor(240, 249, 255)  # Crisp Sky Tint
    sky.line.fill.background()
    shapes.append(sky)

    # 2. Golden Sun Orb
    sun = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(sun_pos[0] + pan_x * 0.4),
        Inches(sun_pos[1]),
        Inches(1.1 * zoom),
        Inches(1.1 * zoom),
    )
    sun.name = "!!Nature_Sun"
    sun.fill.solid()
    sun.fill.fore_color.rgb = RGBColor(251, 191, 36)  # Warm Solar Gold
    sun.line.color.rgb = RGBColor(254, 243, 199)
    sun.line.width = Pt(2.0)
    shapes.append(sun)

    # 3. Distant Rolling Hill / Grassland (0.7x Parallax)
    back_hill = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(-2.0 + pan_x * 0.7),
        Inches(5.3),
        Inches(17.5),
        Inches(4.2),
    )
    back_hill.name = "!!Meadow_Back"
    back_hill.fill.solid()
    back_hill.fill.fore_color.rgb = RGBColor(110, 231, 183)  # Soft Mint Meadow
    back_hill.line.fill.background()
    shapes.append(back_hill)

    # 4. Foreground Lush Green Meadow (1.0x Parallax)
    front_meadow = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(-1.5 + pan_x * 1.0),
        Inches(5.85),
        Inches(16.8),
        Inches(3.6),
    )
    front_meadow.name = "!!Meadow_Front"
    front_meadow.fill.solid()
    front_meadow.fill.fore_color.rgb = RGBColor(16, 185, 129)  # Rich Emerald Grassland
    front_meadow.line.fill.background()
    shapes.append(front_meadow)

    return shapes


def add_editable_kite(
    slide,
    left: float,
    top: float,
    size: float = 0.95,
    rot_deg: float = 15.0,
    name_prefix: str = "!!Traveling_Kite",
) -> List[Any]:
    """
    Constructs a vibrant, native editable Flying Kite traveling entity:
    - Primary Diamond Kite Sail (!!Traveling_Kite)
    - Secondary Accent Wing (!!Traveling_Kite_Wing)
    - Flowing Tail Ribbon Nodes (!!Traveling_Kite_Tail1, !!Traveling_Kite_Tail2)
    """
    shapes = []

    # 1. Main Kite Diamond Sail
    kite = slide.shapes.add_shape(
        MSO_SHAPE.DIAMOND,
        Inches(left),
        Inches(top),
        Inches(size),
        Inches(size * 1.18),
    )
    kite.name = name_prefix
    kite.fill.solid()
    kite.fill.fore_color.rgb = RGBColor(245, 158, 11)  # Vibrant Amber/Coral Kite
    kite.line.color.rgb = RGBColor(255, 255, 255)
    kite.line.width = Pt(1.75)
    _set_rot(kite, rot_deg)
    shapes.append(kite)

    # 2. Inner Kite Accent Cross
    wing = slide.shapes.add_shape(
        MSO_SHAPE.DIAMOND,
        Inches(left + size * 0.22),
        Inches(top + size * 0.22),
        Inches(size * 0.56),
        Inches(size * 0.65),
    )
    wing.name = f"{name_prefix}_Core"
    wing.fill.solid()
    wing.fill.fore_color.rgb = RGBColor(14, 165, 233)  # Sky Cyan Core
    wing.line.fill.background()
    _set_rot(wing, rot_deg)
    shapes.append(wing)

    # 3. Kite Tail Ribbon Node 1
    tail1 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left - size * 0.18),
        Inches(top + size * 1.05),
        Inches(size * 0.22),
        Inches(size * 0.22),
    )
    tail1.name = f"{name_prefix}_Tail1"
    tail1.fill.solid()
    tail1.fill.fore_color.rgb = RGBColor(239, 68, 68)
    tail1.line.fill.background()
    shapes.append(tail1)

    # 4. Kite Tail Ribbon Node 2
    tail2 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left - size * 0.36),
        Inches(top + size * 1.28),
        Inches(size * 0.16),
        Inches(size * 0.16),
    )
    tail2.name = f"{name_prefix}_Tail2"
    tail2.fill.solid()
    tail2.fill.fore_color.rgb = RGBColor(16, 185, 129)
    tail2.line.fill.background()
    shapes.append(tail2)

    return shapes


def build_environmental_protection_deck(output_path: str, num_slides: int = 10) -> str:
    """
    Canonical PMG Engine builder for a 10-slide presentation about Environmental Protection.
    Includes:
    - Editable Earth globe (ocean, continents, atmosphere, clouds)
    - Lush green grassland / meadow multi-plane landscapes & blue sky
    - Recurring Flying Kite ('!!Traveling_Kite') guiding the audience across all 10 slides
    - Frosted glassmorphism telemetry cards, metric badges, and progress scrubber
    - Hardware-accelerated 60 FPS Morph & spotlight transitions
    """
    deck = MotionDeck(theme=Themes.LIGHT_ECO_EMERALD)

    slides_content = [
        {
            "badge": "01 // GLOBAL INITIATIVE",
            "title": "PROTECTING OUR LIVING EARTH",
            "subtitle": "Restoring ecosystems, clean skies, and green grasslands for future generations.",
            "m1": ("BIODIVERSITY", "8.7M Species", "Interconnected habitats across land and ocean"),
            "m2": ("CLEAN AIR TARGET", "100% Renewables", "Transitioning global energy to zero-emission sources"),
            "earth": (8.2, 1.5, 3.6, 0.0),
            "kite": (6.3, 4.5, 0.95, -18.0),
            "pan_x": 0.0,
            "sun": (11.2, 0.6),
            "trans": "none",
        },
        {
            "badge": "02 // PLANETARY BIOSPHERE",
            "title": "THE DELICATE ATMOSPHERIC SHIELD",
            "subtitle": "Earth's thin atmosphere regulates climate and sustains every meadow and ocean.",
            "m1": ("CARBON BALANCE", "350 PPM Goal", "Stabilizing atmospheric greenhouse gas concentrations"),
            "m2": ("OCEAN ABSORPTION", "30% CO2 Captured", "Marine ecosystems buffer global temperature rise"),
            "earth": (7.6, 1.3, 4.2, 25.0),
            "kite": (5.8, 1.0, 1.0, -12.0),
            "pan_x": -0.3,
            "sun": (10.6, 0.7),
            "trans": "morph",
        },
        {
            "badge": "03 // GRASSLAND ECOSYSTEMS",
            "title": "LUSH MEADOWS & SOIL CARBON SINKS",
            "subtitle": "Healthy grasslands store vast carbon reserves beneath thriving root networks.",
            "m1": ("SOIL STORAGE", "2,500 GT Carbon", "Grasslands lock carbon safely in deep organic soil"),
            "m2": ("POLLINATOR HABITAT", "+65% Recovery", "Restored wildflower meadows revive native pollinators"),
            "earth": (9.2, 2.2, 2.8, 50.0),
            "kite": (6.8, 1.2, 1.05, 8.0),
            "pan_x": -0.6,
            "sun": (10.0, 0.8),
            "trans": "morph",
        },
        {
            "badge": "04 // CLEAN WIND & SKIES",
            "title": "HARNESSING PURE ATMOSPHERIC CURRENTS",
            "subtitle": "Clean wind energy powers communities while keeping skies crystal clear.",
            "m1": ("WIND CAPACITY", "1,021 GW Global", "Clean turbines displacing fossil fuel generation"),
            "m2": ("AIR QUALITY INDEX", "-48% Particulates", "Clean energy corridors dramatically improve urban health"),
            "earth": (9.6, 2.6, 2.4, 75.0),
            "kite": (8.2, 0.9, 1.1, -18.0),
            "pan_x": -0.9,
            "sun": (9.4, 0.7),
            "trans": "morph",
        },
        {
            "badge": "05 // REFORESTATION CORRIDORS",
            "title": "REBUILDING ANCIENT GREEN CANOPIES",
            "subtitle": "Connecting fragmented landscapes with native trees and watershed buffers.",
            "m1": ("TREES PLANTED", "1.2B Saplings", "Science-guided native reforestation across continents"),
            "m2": ("WATERSHED PURITY", "94% Clean Flow", "Riparian greenbelts filter runoff naturally"),
            "earth": (8.5, 1.8, 3.3, 105.0),
            "kite": (6.6, 2.1, 0.95, 14.0),
            "pan_x": -0.5,
            "sun": (10.2, 0.9),
            "trans": "morph",
        },
        {
            "badge": "06 // CIRCULAR ZERO-WASTE",
            "title": "DESIGNING OUT POLLUTION & WASTE",
            "subtitle": "Transforming linear consumption into regenerative material loops.",
            "m1": ("RECYCLING RATE", "82% Recovery", "Source separation of organic, recyclable, and inert streams"),
            "m2": ("PLASTIC REDUCTION", "-70% Single-Use", "Biodegradable packaging protecting rivers and coastlines"),
            "earth": (8.0, 1.6, 3.7, 135.0),
            "kite": (5.6, 0.6, 0.9, -15.0),
            "pan_x": -0.2,
            "sun": (10.8, 0.8),
            "trans": "morph",
        },
        {
            "badge": "07 // FRESHWATER RIVERS",
            "title": "GUARDING RIVERS, WETLANDS & AQUIFERS",
            "subtitle": "Clean water flows from protected mountain peaks across fertile valleys.",
            "m1": ("WETLAND RESTORATION", "45M Hectares", "Natural sponges preventing floods and droughts"),
            "m2": ("AQUATIC LIFE", "+52% Fish Stocks", "Removing river barriers restores migratory routes"),
            "earth": (8.6, 2.0, 3.1, 165.0),
            "kite": (6.9, 1.1, 1.0, 18.0),
            "pan_x": -0.7,
            "sun": (9.8, 0.6),
            "trans": "morph",
        },
        {
            "badge": "08 // GREEN URBAN DESIGN",
            "title": "BIOPHILIC CITIES & ECO CAMPUSES",
            "subtitle": "Integrating vertical gardens, solar roofs, and green transit into daily life.",
            "m1": ("URBAN CANOPY", "40% Tree Cover", "Lowering urban heat islands by up to 4.5 degrees Celsius"),
            "m2": ("GREEN TRANSIT", "90% Zero-Emission", "Electric mobility and cycling networks across cities"),
            "earth": (9.0, 2.1, 2.9, 195.0),
            "kite": (9.6, 0.8, 0.95, -10.0),
            "pan_x": -1.0,
            "sun": (9.2, 0.8),
            "trans": "morph",
        },
        {
            "badge": "09 // COMMUNITY STEWARDSHIP",
            "title": "EMPOWERING YOUTH & CITIZEN SCIENCE",
            "subtitle": "Every student and citizen acts as a guardian of our shared planet.",
            "m1": ("ACTIVE VOLUNTEERS", "12M Youth Leaders", "Campus sustainability labs driving measurable local impact"),
            "m2": ("OPEN TELEMETRY", "24/7 Air & Water", "Real-time environmental sensors guiding policy"),
            "earth": (7.8, 1.4, 3.8, 225.0),
            "kite": (6.5, 1.5, 1.05, 12.0),
            "pan_x": -0.4,
            "sun": (10.4, 0.7),
            "trans": "morph",
        },
        {
            "badge": "10 // SUSTAINABLE HORIZON",
            "title": "ONE EARTH, ONE SHARED FUTURE",
            "subtitle": "Together we keep the meadows green, the oceans vibrant, and the skies free.",
            "m1": ("NET-ZERO PLEDGE", "2050 Milestone", "Global coalition committed to climate neutrality"),
            "m2": ("LIVING PLANET", "100% Protected", "A thriving biosphere where nature and humanity prosper"),
            "earth": (7.5, 1.3, 4.2, 260.0),
            "kite": (10.6, 0.8, 1.15, 20.0),
            "pan_x": 0.0,
            "sun": (11.0, 0.6),
            "trans": "circle",
        },
    ]

    selected_slides = slides_content[: max(1, min(num_slides, len(slides_content)))]
    total = len(selected_slides)

    for idx, sc in enumerate(selected_slides, 1):
        slide = deck.add_slide(transition=sc["trans"], duration_ms=1600)

        # 1. Multi-plane Nature & Grassland Landscape
        add_editable_meadow_landscape(slide, pan_x=sc["pan_x"], zoom=1.0, sun_pos=sc["sun"])

        # 2. Editable Multi-Layered Earth Globe
        ex, ey, edia, erot = sc["earth"]
        add_editable_earth(slide, left=ex, top=ey, diameter=edia, rot_deg=erot, name_prefix="!!Earth")

        # 3. Header Badge & Hero Title (Strictly bounded, <= 35 words)
        deck.add_badge(
            slide,
            sc["badge"],
            left=0.9,
            top=0.65,
            width=3.2,
            height=0.38,
            bg_color=RGBColor(5, 150, 105),
            text_color=RGBColor(255, 255, 255),
            name="!!env_badge",
        )

        deck.add_hero_title(
            slide,
            sc["title"],
            subtitle=sc["subtitle"],
            left=0.9,
            top=1.2,
            width=6.2,
            height=1.55,
            title_size=24,
            subtitle_size=11.5,
            title_color=RGBColor(6, 78, 59),
            name="!!env_title",
        )

        # 4. Two Frosted Glass Metric Cards
        t1, v1, d1 = sc["m1"]
        deck.add_metric_card(
            slide,
            title=t1,
            value=v1,
            desc=d1,
            left=0.9,
            top=2.9,
            width=5.6,
            height=1.55,
            accent_color=RGBColor(5, 150, 105),
            name="!!env_card_1",
        )

        t2, v2, d2 = sc["m2"]
        deck.add_metric_card(
            slide,
            title=t2,
            value=v2,
            desc=d2,
            left=0.9,
            top=4.65,
            width=5.6,
            height=1.55,
            accent_color=RGBColor(2, 132, 199),
            name="!!env_card_2",
        )

        # 5. Flying Kite Traveling Entity (soaring continuously across slides)
        kx, ky, ksz, krot = sc["kite"]
        add_editable_kite(slide, left=kx, top=ky, size=ksz, rot_deg=krot, name_prefix="!!Traveling_Kite")

        # 6. Bottom Progress Scrubber Bar
        deck.add_progress_bar(slide, current_step=idx, total_steps=total, top=6.95)

    deck.save(output_path)
    return os.path.abspath(output_path)



def build_v1_2_single_slide_animation_showcase(output_path: str) -> str:
    """
    Canonical v1.2.0 Acceptance Presentation Builder:
    Creates a single 16:9 widescreen slide containing:
    - 12 independent editable PowerPoint objects (>= 8)
    - 18 native PowerPoint Animation Pane effects (>= 15)
    - Multiple simultaneous (with_previous) and sequential (after_previous, on_click) animations
    - Curved Bezier motion paths (M ... C ... E)
    - Rotating elements (<p:animRot>)
    - Scaling elements (<p:animScale>)
    - Timed entrance and exit animations
    """
    from pmg.animations import SlideAnimationTimeline

    deck = MotionDeck(theme=Themes.LIGHT_ECO_EMERALD)
    slide = deck.add_slide(transition="none")

    # 1. Sky Horizon Band (Object 1)
    sky = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.0), Inches(0.0), Inches(13.333), Inches(4.5))
    sky.name = "Sky_Horizon_Band"
    sky.fill.solid()
    sky.fill.fore_color.rgb = RGBColor(224, 242, 254)
    sky.line.fill.background()

    # 2. Meadow Terrain Base (Object 2)
    meadow = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-1.0), Inches(4.4), Inches(15.5), Inches(4.2))
    meadow.name = "Meadow_Terrain_Base"
    meadow.fill.solid()
    meadow.fill.fore_color.rgb = RGBColor(167, 243, 208)
    meadow.line.fill.background()

    # 3. Header Badge (Object 3)
    badge = deck.add_badge(
        slide,
        "PMG v1.2 · NATIVE ANIMATION TIMELINE",
        left=0.85,
        top=0.55,
        width=4.0,
        height=0.38,
        bg_color=RGBColor(5, 150, 105),
        text_color=RGBColor(255, 255, 255),
        name="Header_Badge",
    )

    # 4. Hero Title Banner (Object 4)
    title_box = deck.add_hero_title(
        slide,
        "Autonomous Clean Energy Microgrid",
        subtitle="Synchronized solar harvesting, grid storage, and zero-carbon distribution.",
        left=0.85,
        top=1.05,
        width=7.5,
        height=1.35,
        title_size=24,
        subtitle_size=11.5,
        title_color=RGBColor(6, 78, 59),
        name="Title_Banner",
    )

    # 5. Solar Sun Orb (Object 5 - scales & glows)
    sun = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.8), Inches(0.65), Inches(1.35), Inches(1.35))
    sun.name = "Solar_Sun_Orb"
    sun.fill.solid()
    sun.fill.fore_color.rgb = RGBColor(250, 204, 21)
    sun.line.color.rgb = RGBColor(254, 240, 138)
    sun.line.width = Pt(2.5)

    # 6. Wind Turbine Rotor (Object 6 - rotates 360 deg)
    rotor = slide.shapes.add_shape(MSO_SHAPE.STAR_4_POINT, Inches(9.0), Inches(0.8), Inches(1.25), Inches(1.25))
    rotor.name = "Wind_Turbine_Rotor"
    rotor.fill.solid()
    rotor.fill.fore_color.rgb = RGBColor(2, 132, 199)
    rotor.line.color.rgb = RGBColor(255, 255, 255)
    rotor.line.width = Pt(1.5)

    # 7. Solar Array Card (Object 7)
    card_solar = deck.add_metric_card(
        slide,
        title="SOLAR ARRAY",
        value="4.8 MW",
        desc="High-efficiency photovoltaic collectors.",
        left=0.85,
        top=2.95,
        width=3.6,
        height=2.2,
        accent_color=RGBColor(245, 158, 11),
        name="Solar_Array_Card",
    )

    # 8. Battery Hub Card (Object 8 - scales when energy arrives)
    card_battery = deck.add_metric_card(
        slide,
        title="STORAGE MATRIX",
        value="12.4 MWh",
        desc="Grid-scale solid-state battery buffer.",
        left=4.85,
        top=2.95,
        width=3.6,
        height=2.2,
        accent_color=RGBColor(2, 132, 199),
        name="Battery_Hub_Card",
    )

    # 9. Eco Village Card (Object 9)
    card_village = deck.add_metric_card(
        slide,
        title="ECO COMMUNITY",
        value="100% CLEAN",
        desc="Zero-emission residential microgrid.",
        left=8.85,
        top=2.95,
        width=3.6,
        height=2.2,
        accent_color=RGBColor(5, 150, 105),
        name="Eco_Village_Card",
    )

    # 10. Energy Packet Node (Object 10 - travels along curved Bezier motion paths)
    packet = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.35), Inches(2.32), Inches(0.52), Inches(0.52))
    packet.name = "Energy_Packet_Node"
    packet.fill.solid()
    packet.fill.fore_color.rgb = RGBColor(14, 165, 233)
    packet.line.color.rgb = RGBColor(255, 255, 255)
    packet.line.width = Pt(2.0)

    # 11. Smog Warning Alert (Object 11 - timed entrance then timed exit)
    smog_alert = deck.add_badge(
        slide,
        "PEAK FOSSIL LOAD DETECTED — REROUTING TO CLEAN GRID",
        left=4.5,
        top=5.45,
        width=4.6,
        height=0.42,
        bg_color=RGBColor(239, 68, 68),
        text_color=RGBColor(255, 255, 255),
        name="Smog_Warning_Alert",
    )

    # 12. Telemetry Status Bar (Object 12)
    status_bar = deck.add_badge(
        slide,
        "GRID TELEMETRY: 100% SYNCHRONIZED · ZERO CARBON DISPATCH ACTIVE",
        left=0.85,
        top=6.25,
        width=11.6,
        height=0.52,
        bg_color=RGBColor(6, 78, 59),
        text_color=RGBColor(255, 255, 255),
        name="Telemetry_Status_Bar",
    )

    # Build Declarative Animation Timeline with 18 native PowerPoint effects
    tl = SlideAnimationTimeline(slide_number=1)

    # Group 1: Simultaneous opening reveal (3 parallel effects)
    tl.add_effect(badge, "appear", trigger="with_previous", duration_ms=150, delay_ms=50)
    tl.add_effect(title_box, "fade", trigger="with_previous", duration_ms=350, delay_ms=80)
    tl.add_effect(sun, "zoom", trigger="with_previous", duration_ms=400, delay_ms=100)

    # Group 2: Sequential card array deployment (3 staggered parallel effects)
    tl.add_effect(card_solar, "fly_in", trigger="after_previous", direction="up", duration_ms=400)
    tl.add_effect(card_battery, "fly_in", trigger="with_previous", direction="up", duration_ms=400, delay_ms=80)
    tl.add_effect(card_village, "wipe", trigger="with_previous", direction="right", duration_ms=400, delay_ms=140)

    # Group 3: Turbine & Solar generation activation (3 parallel effects: entrance + rotation + scale)
    tl.add_effect(rotor, "fade", trigger="after_previous", duration_ms=250)
    tl.add_effect(rotor, "rotate", trigger="with_previous", rotation_deg=360.0, duration_ms=550)
    tl.add_effect(sun, "scale", trigger="with_previous", scale_factor=(1.22, 1.22), duration_ms=350, auto_reverse=True)

    # Group 4: Timed entrance of transient load alert (2 parallel effects)
    tl.add_effect(smog_alert, "fade", trigger="after_previous", duration_ms=250)
    tl.add_effect(smog_alert, "fly_in", trigger="with_previous", direction="left", duration_ms=300)

    # Group 5: Energy packet dispatch along Curved Bezier Motion Path 1 (2 parallel effects)
    tl.add_effect(packet, "appear", trigger="after_previous", duration_ms=100)
    tl.add_effect(
        packet,
        "motion_path",
        trigger="with_previous",
        motion_path="M 0 0 C 0.10 -0.11 0.20 0.11 0.30 0 E",
        duration_ms=650,
    )

    # Group 6: Storage matrix pulse & second curved Bezier path to Eco Village (2 parallel effects)
    tl.add_effect(
        card_battery,
        "scale",
        trigger="after_previous",
        scale_factor=(1.10, 1.10),
        duration_ms=300,
        auto_reverse=True,
    )
    tl.add_effect(
        packet,
        "motion_path",
        trigger="with_previous",
        motion_path="M 0 0 C 0.10 -0.09 0.20 0.09 0.30 0 E",
        duration_ms=550,
    )

    # Group 7: Timed exit of warning alert + simultaneous status bar reveal (2 parallel effects)
    tl.add_effect(smog_alert, "fade_out", trigger="after_previous", duration_ms=300)
    tl.add_effect(status_bar, "fly_in", trigger="with_previous", direction="up", duration_ms=350)

    # Group 8: Click-triggered final turbine spin + fly_out of transient warning (2 effects)
    tl.add_effect(rotor, "rotate", trigger="on_click", rotation_deg=360.0, duration_ms=500)
    tl.add_effect(smog_alert, "fly_out", trigger="with_previous", direction="down", duration_ms=350)

    tl.compile_to_openxml(slide)

    deck.save(output_path)
    return os.path.abspath(output_path)
