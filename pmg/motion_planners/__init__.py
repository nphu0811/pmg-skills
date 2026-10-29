"""
pmg.motion_planners
===================
Specialized Object-Specific Motion Intelligence Planners for PMG v2.5.0.

Provides modular planners for each major PowerPoint object category:
- chart_motion: Quantitative bar, line, doughnut, and numerical counters
- kinetic_typography: Titles, subtitles, word staggers, keyword punches
- image_motion: Photos, visual layers, parallax, dolly zoom, and before/after comparisons
- diagram_motion: Flowcharts, process steps, and directional causal connectors
- map_motion: Geospatial maps, territory highlights, marker pins, and route drawings
- infographic_motion: Timelines, progress bars, circular gauges, and comparison matrices
"""

from pmg.motion_planners.base import (
    BaseMotionPlanner,
    MotionIntent,
    SharedMotionPrimitive,
)
from pmg.motion_planners.chart_motion import ChartMotionPlanner
from pmg.motion_planners.kinetic_typography import KineticTypographyPlanner
from pmg.motion_planners.image_motion import ImageMotionPlanner
from pmg.motion_planners.diagram_motion import DiagramMotionPlanner
from pmg.motion_planners.map_motion import MapMotionPlanner
from pmg.motion_planners.infographic_motion import InfographicMotionPlanner

__all__ = [
    "BaseMotionPlanner",
    "MotionIntent",
    "SharedMotionPrimitive",
    "ChartMotionPlanner",
    "KineticTypographyPlanner",
    "ImageMotionPlanner",
    "DiagramMotionPlanner",
    "MapMotionPlanner",
    "InfographicMotionPlanner",
]
