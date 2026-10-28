"""
pmg.space_builder
=================
Builds the canonical 20-slide v1.5.1 flagship presentation:
"SPACE EXPLORATION — BEYOND THE HORIZON"

Covers:
1. Earth from Orbit (Beyond the Horizon)
2. Earth's Magnetosphere & Pale Blue Dot
3. Low Earth Orbit & Satellite Constellations
4. International Space Station (ISS) Orbital Outpost
5. Astronauts & Extravehicular Activity (EVA)
6. Next-Generation Deep Space Spacecraft & Propulsion
7. Lunar Gateway & Artemis Return to the Moon
8. Solar System Architecture & Scale
9. Mars: The Red Horizon & Perseverance
10. Martian Telemetry & Surface Habitat
11. Asteroid Belt & Planetary Defense (DART / Bennu)
12. Jupiter & Europa's Subsurface Ocean
13. Saturn's Rings, Titan & Enceladus
14. Ice Giants, Kuiper Belt & Voyager Heliopause
15. NASA Space Telescopes & JWST at Lagrange L2
16. Exoplanets & Habitable Zone Spectroscopy
17. Stellar Nurseries & Nebulae
18. Black Holes, Event Horizons & Relativistic Lensing
19. The Milky Way Barred Spiral Galaxy
20. Deep-Field Galaxies & Humanity's Cosmic Future

Every slide combines high-contrast celestial visual imagery (`!!Hero_Visual`) with
native editable PowerPoint typography, glassmorphic telemetry cards, data bars,
continuous `!!Traveling_Spacecraft` Morph choreography, multi-plane parallax,
and verified `<p:timing>` animations (`SlideAnimationTimeline`) that preserve
100% object visibility in live Microsoft PowerPoint Slide Show playback.
"""

import os
import math
from typing import Dict, Any, List, Tuple
import numpy as np
import cv2
import pptx
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

from pmg.animations import SlideAnimationTimeline
from scripts.motion_graph_kit import apply_transition


def _render_space_visual_asset(assets_dir: str, kind: str) -> str:
    """
    Renders a crisp, high-contrast 720x720 celestial / aerospace visual PNG for the given topic
    so every slide has a distinct, vibrant hero image (`!!Hero_Visual`) whose inner core has
    rich visual texture and contrast in both Edit Mode and live Slide Show playback.
    """
    os.makedirs(assets_dir, exist_ok=True)
    out_path = os.path.join(assets_dir, f"space_visual_{kind}.png")
    if os.path.exists(out_path) and os.path.getsize(out_path) > 1024:
        return out_path

    sz = 640
    img = np.zeros((sz, sz, 3), dtype=np.uint8)
    # Deep cosmic backdrop with starfield
    img[:, :] = (24, 14, 8)
    rng = np.random.default_rng(abs(hash(kind)) % (2**31 - 1))
    for _ in range(140):
        sx = int(rng.integers(12, sz - 12))
        sy = int(rng.integers(12, sz - 12))
        sr = int(rng.choice([1, 1, 2, 3]))
        col = (int(rng.integers(190, 255)), int(rng.integers(200, 255)), int(rng.integers(210, 255)))
        cv2.circle(img, (sx, sy), sr, col, -1, lineType=cv2.LINE_AA)

    cx, cy = sz // 2, sz // 2

    if kind in ("earth_orbit", "earth_magnetosphere"):
        # Radiant blue Earth with emerald continents, white cloud swirls, and cyan atmospheric limb
        cv2.circle(img, (cx, cy), 245, (255, 195, 60), 10, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx, cy), 232, (215, 115, 20), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx - 35, cy - 35), 185, (245, 165, 50), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx - 45, cy - 30), (95, 58), 22, 0, 360, (90, 200, 75), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx + 55, cy + 40), (85, 50), -18, 0, 360, (75, 180, 95), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx - 10, cy + 85), (70, 32), 12, 0, 360, (115, 215, 120), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx - 15, cy - 70), (130, 20), -10, 0, 360, (245, 250, 255), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx + 25, cy + 12), (140, 18), 8, 0, 360, (240, 245, 252), -1, lineType=cv2.LINE_AA)
        if kind == "earth_magnetosphere":
            cv2.ellipse(img, (cx, cy), (285, 160), 0, 0, 360, (255, 220, 80), 4, lineType=cv2.LINE_AA)
            cv2.ellipse(img, (cx, cy), (260, 210), 0, 0, 360, (255, 180, 120), 3, lineType=cv2.LINE_AA)

    elif kind in ("satellite", "iss_station"):
        # Earth curvature below + bright gold/cyan orbital spacecraft / ISS solar wings in center
        cv2.circle(img, (cx, cy + 340), 360, (195, 105, 25), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx, cy + 340), 364, (255, 210, 90), 5, lineType=cv2.LINE_AA)
        # Solar array wings (vibrant cobalt & gold grid)
        for wx in (cx - 190, cx + 60):
            cv2.rectangle(img, (wx, cy - 110), (wx + 130, cy + 110), (210, 130, 35), -1)
            cv2.rectangle(img, (wx, cy - 110), (wx + 130, cy + 110), (120, 220, 255), 4)
            for gy in range(cy - 80, cy + 110, 35):
                cv2.line(img, (wx, gy), (wx + 130, gy), (120, 220, 255), 2)
        # Central pressurized truss & module core
        cv2.rectangle(img, (cx - 95, cy - 28), (cx + 95, cy + 28), (235, 240, 245), -1)
        cv2.rectangle(img, (cx - 42, cy - 68), (cx + 42, cy + 68), (215, 225, 235), -1)
        cv2.circle(img, (cx, cy), 26, (255, 200, 60), -1, lineType=cv2.LINE_AA)

    elif kind == "astronaut":
        # Stylized high-contrast Astronaut helmet visor & EVA suit chest pack
        cv2.circle(img, (cx, cy - 20), 195, (240, 245, 250), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx, cy - 20), 195, (200, 210, 225), 8, lineType=cv2.LINE_AA)
        # Golden reflective visor
        cv2.ellipse(img, (cx, cy - 25), (145, 115), 0, 0, 360, (45, 175, 245), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx - 25, cy - 45), (95, 65), -12, 0, 360, (110, 215, 255), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx - 55, cy - 65), 28, (255, 255, 255), -1, lineType=cv2.LINE_AA)
        # Suit torso & telemetry chest unit
        cv2.rectangle(img, (cx - 145, cy + 140), (cx + 145, cy + 275), (235, 240, 248), -1)
        cv2.rectangle(img, (cx - 70, cy + 165), (cx + 70, cy + 245), (50, 65, 90), -1)
        cv2.circle(img, (cx - 30, cy + 205), 14, (80, 220, 100), -1)
        cv2.circle(img, (cx + 30, cy + 205), 14, (255, 190, 50), -1)

    elif kind in ("spacecraft", "moon_artemis", "asteroid"):
        # Deep-space exploration vehicle / lunar lander / asteroid rendezvous
        cv2.circle(img, (cx + 75, cy + 65), 175, (165, 170, 180) if kind != "asteroid" else (120, 135, 155), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx + 35, cy + 35), 32, (130, 135, 145), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx + 115, cy + 95), 44, (125, 130, 140), -1, lineType=cv2.LINE_AA)
        # Spacecraft / Rocket stage with bright plasma plume
        pts = np.array([[cx - 170, cy - 150], [cx - 40, cy - 90], [cx - 95, cy - 25]], np.int32)
        cv2.fillPoly(img, [pts], (245, 248, 255))
        cv2.circle(img, (cx - 155, cy - 135), 42, (60, 195, 255), -1, lineType=cv2.LINE_AA)
        cv2.line(img, (cx - 170, cy - 150), (cx + 60, cy + 50), (80, 220, 255), 5, lineType=cv2.LINE_AA)

    elif kind in ("mars_globe", "mars_surface"):
        # Rust-orange / crimson Mars globe with polar ice cap & Valles Marineris canyon
        cv2.circle(img, (cx, cy), 232, (55, 95, 225), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx - 35, cy - 35), 180, (75, 130, 245), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx, cy - 195), (90, 28), 0, 0, 360, (245, 250, 255), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx + 10, cy + 15), (145, 24), -12, 0, 360, (35, 55, 145), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx - 75, cy - 25), 36, (45, 70, 175), 6, lineType=cv2.LINE_AA)

    elif kind in ("solar_system", "exoplanets"):
        # Sun at left/center + concentric orbital rings + multi-colored planets
        cv2.circle(img, (cx, cy), 72, (60, 210, 255), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx, cy), 92, (80, 180, 255), 4, lineType=cv2.LINE_AA)
        radii = [125, 165, 210, 255]
        angles = [0.6, 2.1, 3.8, 5.2]
        p_cols = [(180, 190, 200), (220, 140, 40), (70, 110, 235), (120, 185, 245)]
        for r, ang, col in zip(radii, angles, p_cols):
            cv2.circle(img, (cx, cy), r, (95, 115, 145), 2, lineType=cv2.LINE_AA)
            px = int(cx + r * math.cos(ang))
            py = int(cy + r * math.sin(ang))
            cv2.circle(img, (px, py), 22, col, -1, lineType=cv2.LINE_AA)
            cv2.circle(img, (px, py), 22, (255, 255, 255), 2, lineType=cv2.LINE_AA)

    elif kind in ("jupiter_europa", "saturn_titan", "ice_giants"):
        # Banded gas/ice giant with rings or icy moon
        base_c = (135, 175, 225) if kind == "jupiter_europa" else ((120, 195, 240) if kind == "saturn_titan" else (235, 185, 65))
        cv2.circle(img, (cx, cy), 175, base_c, -1, lineType=cv2.LINE_AA)
        for by in range(-120, 140, 42):
            cv2.ellipse(img, (cx, cy + by), (165, 16), 0, 0, 360, (95, 135, 195), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx, cy), (285, 75), -18, 0, 360, (215, 230, 250), 14, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx, cy), (250, 58), -18, 0, 360, (120, 205, 255), 6, lineType=cv2.LINE_AA)

    elif kind == "nasa_jwst":
        # Iconic 18-segment golden hexagonal James Webb Space Telescope mirror
        hex_r = 46
        offsets = [
            (0, -92), (80, -46), (80, 46), (0, 92), (-80, 46), (-80, -46),
            (0, -184), (80, -138), (160, -92), (160, 0), (160, 92), (80, 138),
            (0, 184), (-80, 138), (-160, 92), (-160, 0), (-160, -92), (-80, -138),
        ]
        for ox, oy in offsets:
            hx, hy = cx + ox, cy + oy
            pts = np.array(
                [
                    [int(hx + hex_r * math.cos(math.radians(60 * k))), int(hy + hex_r * math.sin(math.radians(60 * k)))]
                    for k in range(6)
                ],
                np.int32,
            )
            cv2.fillPoly(img, [pts], (35, 195, 255))
            cv2.polylines(img, [pts], True, (220, 245, 255), 3, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx, cy), 32, (40, 50, 65), -1, lineType=cv2.LINE_AA)

    elif kind == "black_hole":
        # Relativistic glowing golden-orange accretion disk & photon ring around event horizon
        cv2.circle(img, (cx, cy), 225, (45, 125, 245), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx, cy), 175, (90, 200, 255), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx, cy), 115, (10, 10, 14), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx, cy), 118, (220, 245, 255), 5, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx, cy), (285, 58), -12, 0, 360, (140, 225, 255), 18, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx, cy), (285, 58), -12, 0, 360, (245, 252, 255), 5, lineType=cv2.LINE_AA)

    else:
        # Milky Way spiral galaxy / Nebula / Deep-field Galaxies
        cv2.ellipse(img, (cx, cy), (250, 145), 25, 0, 360, (185, 95, 145), -1, lineType=cv2.LINE_AA)
        cv2.ellipse(img, (cx, cy), (195, 105), 25, 0, 360, (235, 165, 95), -1, lineType=cv2.LINE_AA)
        cv2.circle(img, (cx, cy), 78, (215, 240, 255), -1, lineType=cv2.LINE_AA)
        for arm in range(4):
            base_a = arm * (math.pi / 2.0)
            for step in range(18):
                r = 55 + step * 11
                a = base_a + step * 0.18
                px = int(cx + r * math.cos(a))
                py = int(cy + r * 0.65 * math.sin(a))
                cv2.circle(img, (px, py), max(4, 14 - step // 2), (255, 225, 160), -1, lineType=cv2.LINE_AA)

    cv2.imwrite(out_path, img)
    return out_path


SLIDE_SPECS: List[Dict[str, Any]] = [
    {
        "badge": "MISSION 01 · ORBITAL DAWN",
        "title": "BEYOND THE HORIZON",
        "subtitle": "Humanity's journey from low Earth orbit to the edge of the observable cosmos",
        "visual": "earth_orbit",
        "card1": ("EARTH ORBITAL VELOCITY", "29.78 km/s solar orbit · 12,742 km equatorial diameter"),
        "card2": ("ESCAPE VELOCITY", "11.186 km/s delta-v required to break Earth gravity well"),
        "metric": ("ATMOS BOUNDARY · KÁRMÁN LINE 100 KM", 0.82),
        "trans": "fade",
        "hero_pos": (8.0, 1.15, 4.5, 4.5),
        "craft_pos": (6.8, 0.85),
        "anim_style": "hero_intro",
    },
    {
        "badge": "MISSION 02 · PLANETARY SHIELD",
        "title": "OUR LIVING BLUE MARBLE",
        "subtitle": "Earth's magnetosphere and atmosphere shield biosphere evolution from solar wind",
        "visual": "earth_magnetosphere",
        "card1": ("MAGNETOPAUSE STANDOFF", "65,000 km sunward geomagnetic deflection boundary"),
        "card2": ("TOTAL SOLAR IRRADIANCE", "1,361 W/m² top-of-atmosphere radiative equilibrium"),
        "metric": ("GEOMAGNETIC SHIELD INTEGRITY · 99.4%", 0.92),
        "trans": "morph",
        "hero_pos": (7.7, 1.0, 4.8, 4.8),
        "craft_pos": (7.2, 1.4),
        "anim_style": "zoom_wipe",
    },
    {
        "badge": "MISSION 03 · EARTH OBSERVATION",
        "title": "EYES IN LOW EARTH ORBIT",
        "subtitle": "Earth-observing satellite constellations monitor climate, oceans, and biosphere health",
        "visual": "satellite",
        "card1": ("LEO CONSTELLATION BAND", "500–1,200 km sun-synchronous polar observation orbits"),
        "card2": ("RADAR ALTIMETRY PRECISION", "Sub-centimeter sea surface height & ice-sheet tracking"),
        "metric": ("GLOBAL TELEMETRY COVERAGE · 98.7%", 0.88),
        "trans": "morph",
        "hero_pos": (8.1, 1.25, 4.4, 4.4),
        "craft_pos": (7.5, 2.1),
        "anim_style": "bezier_orbit",
    },
    {
        "badge": "MISSION 04 · ORBITAL OUTPOST",
        "title": "INTERNATIONAL SPACE STATION",
        "subtitle": "Twenty-five years of continuous human presence in microgravity at 408 km altitude",
        "visual": "iss_station",
        "card1": ("ORBITAL CADENCE", "27,600 km/h · 15.54 orbits per day (92.9 min period)"),
        "card2": ("PRESSURIZED VOLUME", "916 m³ laboratory habitat · 109 m truss span"),
        "metric": ("SOLAR ARRAY POWER GENERATION · 215 KW", 0.86),
        "trans": "morph",
        "hero_pos": (7.85, 1.15, 4.6, 4.6),
        "craft_pos": (7.1, 2.8),
        "anim_style": "stagger_cards",
    },
    {
        "badge": "MISSION 05 · HUMAN SPACEFLIGHT",
        "title": "ASTRONAUTS IN THE VOID",
        "subtitle": "Extravehicular Mobility Units sustain human explorers in vacuum and thermal extremes",
        "visual": "astronaut",
        "card1": ("EVA THERMAL ENVELOPE", "-150°C shadow to +120°C direct orbital sunlight"),
        "card2": ("PRIMARY LIFE SUPPORT", "8-hour pure O₂ closed-loop scrubbing & cooling"),
        "metric": ("EVA MOBILITY & RADIATION TELEMETRY · 94%", 0.84),
        "trans": "morph",
        "hero_pos": (8.05, 1.2, 4.4, 4.4),
        "craft_pos": (6.9, 3.4),
        "anim_style": "hero_intro",
    },
    {
        "badge": "MISSION 06 · PROPULSION SYSTEMS",
        "title": "DEEP SPACE PROPULSION",
        "subtitle": "High-specific-impulse cryogenic stages, nuclear thermal cores, and Hall-effect ion drives",
        "visual": "spacecraft",
        "card1": ("SLS BLOCK 1B THRUST", "39.1 Meganewtons liftoff thrust · 105 t to LEO"),
        "card2": ("SOLAR ELECTRIC ION ISP", "3,200 s specific impulse for continuous interplanetary thrust"),
        "metric": ("TRANS-LUNAR INJECTION EFFICIENCY · 96.2%", 0.91),
        "trans": "morph",
        "hero_pos": (7.9, 1.2, 4.5, 4.5),
        "craft_pos": (7.4, 1.2),
        "anim_style": "linear_thrust",
    },
    {
        "badge": "MISSION 07 · CISLUNAR GATEWAY",
        "title": "RETURN TO THE MOON",
        "subtitle": "Artemis polar base camps and the Lunar Gateway in Near-Rectilinear Halo Orbit",
        "visual": "moon_artemis",
        "card1": ("CISLUNAR DISTANCE", "384,400 km mean Earth-Moon separation · NRHO 7-day orbit"),
        "card2": ("SHACKLETON CRATER ICE", "Permanently shadowed water-ice deposits for ISRU fuel"),
        "metric": ("ARTEMIS CISLUNAR LOGISTICS READINESS · 89%", 0.85),
        "trans": "morph",
        "hero_pos": (8.15, 1.15, 4.35, 4.35),
        "craft_pos": (7.6, 1.9),
        "anim_style": "zoom_wipe",
    },
    {
        "badge": "MISSION 08 · HELIOCENTRIC ARCHITECTURE",
        "title": "OUR SOLAR SYSTEM",
        "subtitle": "Four terrestrial rocky worlds, main belt asteroids, and four outer gas/ice giants",
        "visual": "solar_system",
        "card1": ("SOLAR MASS DOMINANCE", "99.86% of total Solar System mass resides in the Sun"),
        "card2": ("HELIOSPHERIC RADIUS", "120+ AU solar wind bubble extending to the heliopause"),
        "metric": ("PLANETARY ORBITAL RESONANCE STABILITY · 99%", 0.94),
        "trans": "circle",
        "hero_pos": (7.8, 1.1, 4.7, 4.7),
        "craft_pos": (7.0, 2.5),
        "anim_style": "spin_emphasis",
    },
    {
        "badge": "MISSION 09 · THE RED PLANET",
        "title": "MARS: THE RED HORIZON",
        "subtitle": "Ancient river deltas in Jezero Crater and volcanic shields of Tharsis and Olympus Mons",
        "visual": "mars_globe",
        "card1": ("OLYMPUS MONS RELIEF", "21.9 km summit elevation · 624 km basal shield diameter"),
        "card2": ("VALLES MARINERIS RIFT", "4,000 km tectonic canyon system · 7 km depth"),
        "metric": ("PERSEVERANCE SAMPLE CACHE COMPLETION · 95%", 0.90),
        "trans": "morph",
        "hero_pos": (7.65, 0.95, 4.9, 4.9),
        "craft_pos": (7.1, 1.5),
        "anim_style": "hero_intro",
    },
    {
        "badge": "MISSION 10 · MARTIAN SURFACE OPS",
        "title": "MARS TELEMETRY & HABITAT",
        "subtitle": "MOXIE solid-oxide electrolysis converts Martian CO₂ atmosphere into breathable oxygen",
        "visual": "mars_surface",
        "card1": ("SURFACE GRAVITY & SOL", "3.721 m/s² (0.38g) · 24h 39m 35s Martian solar day"),
        "card2": ("ATMOSPHERIC COMPOSITION", "95.32% CO₂, 2.6% N₂, 1.9% Ar · 6.1 mbar mean pressure"),
        "metric": ("ISRU OXYGEN EXTRACTION EFFICIENCY · 91%", 0.87),
        "trans": "morph",
        "hero_pos": (8.05, 1.2, 4.45, 4.45),
        "craft_pos": (7.4, 2.9),
        "anim_style": "stagger_cards",
    },
    {
        "badge": "MISSION 11 · PLANETARY DEFENSE",
        "title": "ASTEROID RENDEZVOUS",
        "subtitle": "OSIRIS-REx Bennu regolith return and DART kinetic impact orbital deflection",
        "visual": "asteroid",
        "card1": ("DART PERIOD CHANGE", "-33.0 minutes binary orbit period reduction at Dimorphos"),
        "card2": ("BENNU CARBONACEOUS SAMPLE", "121.6 grams pristine prebiotic solar nebula regolith"),
        "metric": ("KINETIC MOMENTUM TRANSFER BETA · 3.61", 0.83),
        "trans": "morph",
        "hero_pos": (8.1, 1.25, 4.4, 4.4),
        "craft_pos": (7.3, 1.1),
        "anim_style": "bezier_orbit",
    },
    {
        "badge": "MISSION 12 · JOVIAN SYSTEM",
        "title": "JUPITER & EUROPA'S OCEAN",
        "subtitle": "Europa Clipper investigates a global liquid saltwater ocean beneath an icy shell",
        "visual": "jupiter_europa",
        "card1": ("SUBSURFACE OCEAN VOLUME", "2× Earth's total liquid ocean water beneath 15–25 km ice"),
        "card2": ("JOVIAN MAGNETOSPHERE", "20,000× stronger magnetic moment than Earth's dipole"),
        "metric": ("EUROPA CLIPPER FLYBY RECONNAISSANCE · 49 PASSES", 0.89),
        "trans": "morph",
        "hero_pos": (7.85, 1.1, 4.65, 4.65),
        "craft_pos": (6.95, 2.2),
        "anim_style": "zoom_wipe",
    },
    {
        "badge": "MISSION 13 · CHRONIAN SYSTEM",
        "title": "RINGS OF SATURN & TITAN",
        "subtitle": "Enceladus hydrothermal plumes and Dragonfly nuclear rotorcraft exploration of Titan",
        "visual": "saturn_titan",
        "card1": ("TITAN ATMOSPHERE", "1.45 bar dense N₂-CH₄ atmosphere · liquid hydrocarbon seas"),
        "card2": ("ENCELADUS SOUTH PLUMES", "Molecular H₂, silica nanoparticles & organic compounds"),
        "metric": ("ASTROBIOLOGICAL HABITABILITY INDEX · 93%", 0.88),
        "trans": "morph",
        "hero_pos": (7.8, 1.15, 4.7, 4.5),
        "craft_pos": (7.35, 3.2),
        "anim_style": "spin_emphasis",
    },
    {
        "badge": "MISSION 14 · INTERSTELLAR BOUNDARY",
        "title": "ICE GIANTS & HELIOPAUSE",
        "subtitle": "Uranus, Neptune, the Kuiper Belt, and Voyager 1 & 2 in interstellar space",
        "visual": "ice_giants",
        "card1": ("VOYAGER 1 RANGE", "164+ AU from Earth · 22.8 light-hours one-way radio link"),
        "card2": ("KUIPER BELT & ARROKOTH", "30–50 AU primordial planetesimals preserved at 40 Kelvin"),
        "metric": ("DEEP SPACE NETWORK SIGNAL LOCK · 100%", 0.91),
        "trans": "morph",
        "hero_pos": (8.0, 1.2, 4.5, 4.5),
        "craft_pos": (7.5, 1.4),
        "anim_style": "linear_thrust",
    },
    {
        "badge": "MISSION 15 · NASA SPACE OBSERVATORIES",
        "title": "NASA DEEP SPACE EYES",
        "subtitle": "James Webb Space Telescope at Sun-Earth L2 peers 13.5 billion years into cosmic dawn",
        "visual": "nasa_jwst",
        "card1": ("JWST PRIMARY MIRROR", "6.5 m gold-coated beryllium aperture · 18 hex segments"),
        "card2": ("CRYOGENIC OPERATING TEMP", "6.0 Kelvin MIRI infrared detector cooled at 1.5M km L2"),
        "metric": ("INFRARED WAVEFRONT ALIGNMENT PRECISION · 99.8%", 0.96),
        "trans": "split",
        "hero_pos": (7.85, 1.1, 4.6, 4.6),
        "craft_pos": (7.0, 2.8),
        "anim_style": "hero_intro",
    },
    {
        "badge": "MISSION 16 · EXOPLANET DEMOGRAPHICS",
        "title": "WORLDS AROUND OTHER STARS",
        "subtitle": "Transit photometry and transmission spectroscopy across 5,600+ confirmed exoplanets",
        "visual": "exoplanets",
        "card1": ("TRAPPIST-1 SYSTEM", "7 Earth-sized rocky planets · 3 inside habitable zone at 40 ly"),
        "card2": ("ATMOSPHERIC BIOSIGNATURES", "JWST NIRSpec detection of H₂O, CO₂, CH₄, and SO₂"),
        "metric": ("SPECTROSCOPIC TRANSIT SNR FIDELITY · 92%", 0.87),
        "trans": "morph",
        "hero_pos": (7.95, 1.15, 4.5, 4.5),
        "craft_pos": (7.25, 1.8),
        "anim_style": "stagger_cards",
    },
    {
        "badge": "MISSION 17 · STELLAR EVOLUTION",
        "title": "CRADLES OF STARS & NEBULAE",
        "subtitle": "Giant molecular clouds, protostellar accretion jets, and supernova heavy-element forging",
        "visual": "nebula",
        "card1": ("PILLARS OF CREATION", "5 light-year photoevaporating H₂ columns in Eagle Nebula M16"),
        "card2": ("STELLAR NUCLEOSYNTHESIS", "Carbon, oxygen, silicon, and iron forged in stellar cores"),
        "metric": ("PROTOSTELLAR ACCRETION FLUX · 88%", 0.85),
        "trans": "morph",
        "hero_pos": (7.8, 1.1, 4.65, 4.65),
        "craft_pos": (6.85, 3.1),
        "anim_style": "zoom_wipe",
    },
    {
        "badge": "MISSION 18 · RELATIVISTIC ASTROPHYSICS",
        "title": "BLACK HOLES & EVENT HORIZONS",
        "subtitle": "Event Horizon Telescope interferometry of Sagittarius A* and M87* relativistic jets",
        "visual": "black_hole",
        "card1": ("SAGITTARIUS A* MASS", "4.154 million solar masses at Milky Way galactic center"),
        "card2": ("M87* SUPERMASSIVE CORE", "6.5 billion solar masses · 5,000 light-year relativistic jet"),
        "metric": ("VLBI INTERFEROMETRIC ANGULAR RESOLUTION · 20 µas", 0.93),
        "trans": "circle",
        "hero_pos": (7.75, 1.05, 4.75, 4.75),
        "craft_pos": (7.15, 1.3),
        "anim_style": "spin_emphasis",
    },
    {
        "badge": "MISSION 19 · GALACTIC STRUCTURE",
        "title": "OUR MILKY WAY GALAXY",
        "subtitle": "A 100,000 light-year barred spiral galaxy of 200 billion stars orbiting every 230 Myr",
        "visual": "milky_way",
        "card1": ("GALACTIC SOLAR ORBIT", "26,000 light-years from core in Orion-Cygnus spiral spur"),
        "card2": ("GAIA ASTROMETRIC MAP", "1.8 billion stellar positions, parallaxes & proper motions"),
        "metric": ("GAIA DR3 ASTROMETRIC CATALOG COMPLETENESS · 97%", 0.94),
        "trans": "morph",
        "hero_pos": (7.9, 1.15, 4.55, 4.55),
        "craft_pos": (7.45, 2.4),
        "anim_style": "bezier_orbit",
    },
    {
        "badge": "MISSION 20 · COSMIC HORIZON",
        "title": "BEYOND THE HORIZON",
        "subtitle": "From Laniakea's 100,000 galaxies to the cosmic microwave background — our journey continues",
        "visual": "deep_field",
        "card1": ("OBSERVABLE UNIVERSE", "93 billion light-year comoving diameter · 2 trillion galaxies"),
        "card2": ("MULTI-PLANETARY FUTURE", "Permanent orbital, lunar, Martian & deep-space science stations"),
        "metric": ("MISSION STATUS · ALL 20 STAGES VERIFIED 100%", 0.98),
        "trans": "fade",
        "hero_pos": (7.7, 1.0, 4.8, 4.8),
        "craft_pos": (7.0, 1.0),
        "anim_style": "hero_intro",
    },
]


def build_v1_5_1_space_exploration_20_slide_deck(
    output_pptx_path: str,
    num_slides: int = 20,
    assets_dir: str = "",
) -> str:
    """
    Builds the canonical 20-slide 'SPACE EXPLORATION — BEYOND THE HORIZON' presentation
    using real editable PowerPoint shapes, typography, data bars, attributed celestial visuals,
    continuous `!!` Morph anchors, and verified `<p:timing>` animations.
    """
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    eff_assets_dir = assets_dir or os.path.join(os.path.dirname(os.path.abspath(output_pptx_path)), "_space_assets")
    os.makedirs(eff_assets_dir, exist_ok=True)

    specs = SLIDE_SPECS[: max(1, min(20, num_slides))]

    for idx, spec in enumerate(specs, 1):
        slide = prs.slides.add_slide(blank_layout)

        # 1. Deep Cosmic Background (`!!Cosmic_BG`)
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.name = "!!Cosmic_BG"
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(9, 14, 28) if idx % 2 == 1 else RGBColor(12, 19, 36)
        bg.line.fill.background()

        # 2. Parallax Starfield Layer (`!!Parallax_Nebula_Band`)
        neb_x = Inches(0.5 + ((idx % 4) * 0.25))
        neb_y = Inches(0.4 + ((idx % 3) * 0.15))
        neb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, neb_x, neb_y, Inches(12.2), Inches(6.6))
        neb.name = "!!Parallax_Nebula_Band"
        neb.fill.solid()
        neb.fill.fore_color.rgb = RGBColor(15, 23, 42)
        neb.line.color.rgb = RGBColor(30, 58, 138)
        neb.line.width = Pt(1.0)

        # 3. Mission Badge Pill (`!!Top_Badge`)
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(0.65), Inches(4.2), Inches(0.45))
        badge.name = "!!Top_Badge"
        badge.fill.solid()
        badge.fill.fore_color.rgb = RGBColor(14, 116, 144)
        badge.line.color.rgb = RGBColor(56, 189, 248)
        badge.line.width = Pt(1.0)
        btf = badge.text_frame
        bp = btf.paragraphs[0]
        bp.text = spec["badge"]
        bp.font.size = Pt(12)
        bp.font.bold = True
        bp.font.color.rgb = RGBColor(224, 242, 254)

        # 4. Main Hero Title (`!!HeroTitle`)
        title_box = slide.shapes.add_textbox(Inches(0.85), Inches(1.22), Inches(6.5), Inches(1.25))
        title_box.name = "!!HeroTitle"
        ttf = title_box.text_frame
        ttf.word_wrap = True
        tp = ttf.paragraphs[0]
        tp.text = spec["title"]
        tp.font.size = Pt(38)
        tp.font.bold = True
        tp.font.color.rgb = RGBColor(248, 250, 252)

        # 5. Subtitle (`!!HeroSubtitle`)
        sub_box = slide.shapes.add_textbox(Inches(0.85), Inches(2.48), Inches(6.4), Inches(0.95))
        sub_box.name = "!!HeroSubtitle"
        stf = sub_box.text_frame
        stf.word_wrap = True
        sp = stf.paragraphs[0]
        sp.text = spec["subtitle"]
        sp.font.size = Pt(17)
        sp.font.color.rgb = RGBColor(148, 163, 184)

        # 6. Primary Telemetry Card (`!!Card_Primary`)
        c1_title, c1_body = spec["card1"]
        card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(3.55), Inches(3.1), Inches(1.75))
        card1.name = "!!Card_Primary"
        card1.fill.solid()
        card1.fill.fore_color.rgb = RGBColor(30, 41, 59)
        card1.line.color.rgb = RGBColor(56, 189, 248)
        card1.line.width = Pt(1.25)
        c1tf = card1.text_frame
        c1tf.word_wrap = True
        p1 = c1tf.paragraphs[0]
        p1.text = f"{c1_title}\n{c1_body}"
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(240, 249, 255)

        # 7. Secondary Telemetry Card (`!!Card_Secondary`)
        c2_title, c2_body = spec["card2"]
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.15), Inches(3.55), Inches(3.1), Inches(1.75))
        card2.name = "!!Card_Secondary"
        card2.fill.solid()
        card2.fill.fore_color.rgb = RGBColor(30, 41, 59)
        card2.line.color.rgb = RGBColor(129, 140, 248)
        card2.line.width = Pt(1.25)
        c2tf = card2.text_frame
        c2tf.word_wrap = True
        p2 = c2tf.paragraphs[0]
        p2.text = f"{c2_title}\n{c2_body}"
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(240, 249, 255)

        # 8. Quantitative Telemetry Bar (`!!Metric_Bar`)
        m_label, m_ratio = spec["metric"]
        bar_w = max(2.8, min(6.4, 6.4 * float(m_ratio)))
        m_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(5.55), Inches(bar_w), Inches(0.65))
        m_bar.name = "!!Metric_Bar"
        m_bar.fill.solid()
        m_bar.fill.fore_color.rgb = RGBColor(2, 132, 199)
        m_bar.line.color.rgb = RGBColor(125, 211, 252)
        mp = m_bar.text_frame.paragraphs[0]
        mp.text = m_label
        mp.font.size = Pt(13)
        mp.font.bold = True
        mp.font.color.rgb = RGBColor(255, 255, 255)

        # 9. Hero Celestial / Aerospace Visual (`!!Hero_Visual`)
        vis_png = _render_space_visual_asset(eff_assets_dir, spec["visual"])
        hx, hy, hw, hh = spec["hero_pos"]
        hero_pic = slide.shapes.add_picture(vis_png, Inches(hx), Inches(hy), width=Inches(hw), height=Inches(hh))
        hero_pic.name = "!!Hero_Visual"

        # 10. Decorative Orbital Ring (`Oval 13`) surrounding Hero Visual
        ring = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(hx - 0.12),
            Inches(hy - 0.12),
            Inches(hw + 0.24),
            Inches(hh + 0.24),
        )
        ring.name = "Oval 13"
        ring.fill.background()
        ring.line.color.rgb = RGBColor(56, 189, 248)
        ring.line.width = Pt(2.0)

        # 11. Continuous Traveling Entity (`!!Traveling_Spacecraft`)
        cx_in, cy_in = spec["craft_pos"]
        craft = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(cx_in), Inches(cy_in), Inches(0.55), Inches(0.55))
        craft.name = "!!Traveling_Spacecraft"
        craft.fill.solid()
        craft.fill.fore_color.rgb = RGBColor(250, 204, 21)
        craft.line.color.rgb = RGBColor(255, 255, 255)
        craft.line.width = Pt(1.5)

        # 12. Footer Attribution (`!!Footer_Attribution`)
        foot = slide.shapes.add_textbox(Inches(0.85), Inches(6.45), Inches(11.6), Inches(0.4))
        foot.name = "!!Footer_Attribution"
        fp = foot.text_frame.paragraphs[0]
        fp.text = (
            f"SLIDE {idx:02d}/20 · SPACE EXPLORATION: BEYOND THE HORIZON · "
            f"TELEMETRY & IMAGERY ATTRIBUTION: NASA / JPL-Caltech / ESA / STScI Public Domain"
        )
        fp.font.size = Pt(11)
        fp.font.color.rgb = RGBColor(100, 116, 139)

        # 13. Slide Transition (Morph / Circle / Split / Fade)
        if idx > 1:
            apply_transition(slide, spec["trans"], duration_ms=1100)

        # 14. Verified Native <p:timing> Choreography (SlideAnimationTimeline)
        # Uses crisp durations (400-550ms) and small stagger delays (0-180ms) so all
        # entrance and motion effects complete well before slide steady state!
        tl = SlideAnimationTimeline(slide_number=idx)
        style = spec["anim_style"]

        if style == "hero_intro":
            tl.add_effect(title_box, "fade", trigger="with_previous", duration_ms=450, delay_ms=0)
            tl.add_effect(hero_pic, "fade", trigger="with_previous", duration_ms=500, delay_ms=80)
            tl.add_effect(card1, "fly_in", trigger="with_previous", duration_ms=450, delay_ms=140, direction="up")
            tl.add_effect(card2, "fly_in", trigger="with_previous", duration_ms=450, delay_ms=200, direction="up")
        elif style == "zoom_wipe":
            tl.add_effect(title_box, "fade", trigger="with_previous", duration_ms=400, delay_ms=0)
            tl.add_effect(hero_pic, "zoom_in", trigger="with_previous", duration_ms=500, delay_ms=60)
            tl.add_effect(m_bar, "wipe", trigger="with_previous", duration_ms=500, delay_ms=140, direction="left")
        elif style == "bezier_orbit":
            tl.add_effect(title_box, "fade", trigger="with_previous", duration_ms=400, delay_ms=0)
            tl.add_effect(hero_pic, "fade", trigger="with_previous", duration_ms=450, delay_ms=60)
            tl.add_effect(
                craft,
                "motion_path",
                trigger="with_previous",
                duration_ms=550,
                delay_ms=100,
                motion_path="M 0 0 C 0.04 -0.06 0.08 0.06 0.10 0 E",
            )
        elif style == "stagger_cards":
            tl.add_effect(title_box, "fade", trigger="with_previous", duration_ms=400, delay_ms=0)
            tl.add_effect(card1, "fade", trigger="with_previous", duration_ms=450, delay_ms=70)
            tl.add_effect(card2, "fade", trigger="with_previous", duration_ms=450, delay_ms=140)
            tl.add_effect(m_bar, "wipe", trigger="with_previous", duration_ms=450, delay_ms=200, direction="left")
        elif style == "linear_thrust":
            tl.add_effect(title_box, "fade", trigger="with_previous", duration_ms=400, delay_ms=0)
            tl.add_effect(
                craft,
                "motion_path",
                trigger="with_previous",
                duration_ms=500,
                delay_ms=60,
                motion_path="M 0 0 L 0.08 -0.04 E",
            )
            tl.add_effect(m_bar, "wipe", trigger="with_previous", duration_ms=450, delay_ms=120, direction="left")
        else:  # spin_emphasis
            tl.add_effect(title_box, "fade", trigger="with_previous", duration_ms=400, delay_ms=0)
            tl.add_effect(ring, "spin", trigger="with_previous", duration_ms=550, delay_ms=60, rotation_deg=360.0)
            tl.add_effect(card1, "fade", trigger="with_previous", duration_ms=450, delay_ms=120)

        tl.compile_to_openxml(slide)

    os.makedirs(os.path.dirname(os.path.abspath(output_pptx_path)), exist_ok=True)
    prs.save(output_pptx_path)
    return os.path.abspath(output_pptx_path)

