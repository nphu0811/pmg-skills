"""
pmg.semantic_classifier
=======================
Semantic Object Classification for PMG v2.5.0 Universal Object-Specific Motion Intelligence.

Classifies presentation objects into specialized categories (charts, typography, images,
diagrams, maps, infographics, cards, traveling heroes) with narrative importance,
visual hierarchy, reading order, and audience focus intentions.
"""

from __future__ import annotations
from enum import Enum
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any, Tuple


class ObjectType(str, Enum):
    """Broad category defining the object's fundamental visual and data nature."""
    CHART = "chart"                      # Quantitative charts (bar, line, doughnut, counters)
    TYPOGRAPHY = "typography"            # Headlines, words, paragraphs, labels
    IMAGE = "image"                      # Photos, illustrations, segmented visual assets
    DIAGRAM = "diagram"                  # Flowcharts, trees, process nodes, connectors
    MAP = "map"                          # Geospatial maps, regions, markers, routes
    INFOGRAPHIC = "infographic"          # Timelines, progress bars, comparison tables
    CONTAINER_CARD = "container_card"    # Acrylic frosted glass cards, background panels
    TRAVELING_HERO = "traveling_hero"    # Cross-slide persistent Morph traveling entity (!!)
    DECORATIVE = "decorative"            # Background grids, glowing reticles, ambient lines


class SemanticRole(str, Enum):
    """Granular communicative role within the scene graph."""
    # Typography
    TITLE = "title"
    SUBTITLE = "subtitle"
    KEYWORD = "keyword"
    BODY_TEXT = "body_text"
    CALLOUT_TEXT = "callout_text"
    
    # Quantitative / Chart
    DATA_AXIS = "data_axis"
    DATA_SERIES = "data_series"
    DATA_BAR = "data_bar"
    DATA_LINE = "data_line"
    DATA_SLICE = "data_slice"
    METRIC_VALUE = "metric_value"
    METRIC_LABEL = "metric_label"

    # Diagram
    PROCESS_STEP = "process_step"
    DIAGRAM_NODE = "diagram_node"
    CONNECTOR = "connector"
    DECISION_NODE = "decision_node"

    # Map
    GEO_REGION = "geo_region"
    GEO_MARKER = "geo_marker"
    GEO_ROUTE = "geo_route"

    # Infographic
    TIMELINE_SPINE = "timeline_spine"
    MILESTONE = "milestone"
    PROGRESS_BAR = "progress_bar"
    PROGRESS_FILL = "progress_fill"
    COMPARISON_ITEM = "comparison_item"

    # General
    HERO_SUBJECT = "hero_subject"
    BACKGROUND_IMAGE = "background_image"
    FOREGROUND_CUTOUT = "foreground_cutout"
    CONTAINER = "container"
    ACCENT = "accent"


class NarrativeImportance(str, Enum):
    """Psychological visual hierarchy and storytelling priority."""
    HERO = "hero"              # Level 1: Primary eye-magnet (traveling hero, lead metric)
    PRIMARY = "primary"        # Level 2: Core content being explained (bars, steps, title)
    SECONDARY = "secondary"    # Level 3: Contextual supporting details (subtitles, values)
    SUPPORTING = "supporting"  # Level 4: Structural anchors (axes, spines, connectors)
    BACKGROUND = "background"  # Level 5: Ambient glass containers, decorative textures


class AudienceFocus(str, Enum):
    """Intended temporal timing of audience engagement."""
    IMMEDIATE = "immediate"      # Grounding anchor (shown first to frame the context)
    SEQUENTIAL = "sequential"    # Revealed step-by-step to guide comprehension
    CLIMAX = "climax"            # Emphasized last as the key takeaway or conclusion
    CONTINUOUS = "continuous"    # Glides across transitions (Morph persistent entity)


@dataclass
class ObjectSemanticProfile:
    """
    Complete semantic fingerprint of an object, driving intelligent motion selection.
    """
    object_id: str
    object_type: ObjectType
    semantic_role: SemanticRole
    narrative_importance: NarrativeImportance
    visual_hierarchy_level: int = 1   # 1 (Highest / Title) to 5 (Lowest / Subtext)
    reading_order: int = 0            # Sequential chronological index
    audience_focus: AudienceFocus = AudienceFocus.SEQUENTIAL
    spatial_bounds: Optional[Tuple[float, float, float, float]] = None
    data_context: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "object_id": self.object_id,
            "object_type": self.object_type.value,
            "semantic_role": self.semantic_role.value,
            "narrative_importance": self.narrative_importance.value,
            "visual_hierarchy_level": self.visual_hierarchy_level,
            "reading_order": self.reading_order,
            "audience_focus": self.audience_focus.value,
            "spatial_bounds": self.spatial_bounds,
            "data_context": self.data_context,
        }


class SemanticClassifier:
    """
    Infers the semantic role and motion requirements of objects from scene graphs,
    raw shapes, and presentation specifications.
    """

    @classmethod
    def classify(cls, node_or_dict: Any, index: int = 0) -> ObjectSemanticProfile:
        """
        Classifies an input object (SceneNode, dict, or ObjectIR) into a structured profile.
        """
        # Extract common fields
        if hasattr(node_or_dict, "node_id"):
            obj_id = node_or_dict.node_id
            name = getattr(node_or_dict, "name", obj_id)
            kind = getattr(node_or_dict, "kind", "rect")
            text = getattr(node_or_dict, "text", "")
            font_size = getattr(node_or_dict, "font_size_pt", 18.0)
            role = getattr(node_or_dict, "semantic_role", "content")
            metadata = getattr(node_or_dict, "metadata", {})
            transform = getattr(node_or_dict, "transform", None)
            bounds = transform.bounds() if transform and hasattr(transform, "bounds") else None
        elif isinstance(node_or_dict, dict):
            obj_id = node_or_dict.get("id") or node_or_dict.get("node_id") or f"obj_{index}"
            name = node_or_dict.get("name", obj_id)
            kind = node_or_dict.get("kind", "rect")
            text = node_or_dict.get("text") or node_or_dict.get("title", "")
            font_size = float(node_or_dict.get("font_size_pt", 18.0))
            role = node_or_dict.get("semantic_role", "content")
            metadata = node_or_dict.get("metadata", {})
            bounds = node_or_dict.get("bounds")
        else:
            obj_id = getattr(node_or_dict, "id", f"obj_{index}")
            name = getattr(node_or_dict, "name", obj_id)
            kind = getattr(node_or_dict, "kind", "card")
            text = getattr(node_or_dict, "title", "")
            font_size = 18.0
            role = getattr(node_or_dict, "semantic_role", "content")
            metadata = getattr(node_or_dict, "metadata", {})
            bounds = None

        name_lower = name.lower()
        role_lower = str(role).lower()
        kind_lower = str(kind).lower()
        meta_type = str(metadata.get("type", "")).lower()

        # 1. Traveling Hero Entity (persistent !! cross-slide element)
        if name.startswith("!!") and any(k in name_lower for k in ("hero", "orb", "traveling", "satellite", "core", "pin")):
            return ObjectSemanticProfile(
                object_id=obj_id,
                object_type=ObjectType.TRAVELING_HERO,
                semantic_role=SemanticRole.HERO_SUBJECT,
                narrative_importance=NarrativeImportance.HERO,
                visual_hierarchy_level=1,
                reading_order=index,
                audience_focus=AudienceFocus.CONTINUOUS,
                spatial_bounds=bounds,
                data_context=metadata,
            )

        # 2. Infographics & Timelines
        if any(k in meta_type for k in ("infographic", "timeline", "milestone", "progress")) or \
           any(k in role_lower for k in ("timeline", "milestone", "progress_bar", "infographic")) or \
           any(k in name_lower for k in ("timeline", "milestone", "progress", "gauge_")):
            
            if "spine" in name_lower or "spine" in role_lower:
                sem_role = SemanticRole.TIMELINE_SPINE
                importance = NarrativeImportance.SUPPORTING
                hierarchy = 4
                focus = AudienceFocus.IMMEDIATE
            elif "progress" in name_lower or "progress" in role_lower:
                sem_role = SemanticRole.PROGRESS_BAR
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2
                focus = AudienceFocus.SEQUENTIAL
            else:
                sem_role = SemanticRole.MILESTONE
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2
                focus = AudienceFocus.SEQUENTIAL

            return ObjectSemanticProfile(
                object_id=obj_id,
                object_type=ObjectType.INFOGRAPHIC,
                semantic_role=sem_role,
                narrative_importance=importance,
                visual_hierarchy_level=hierarchy,
                reading_order=index,
                audience_focus=focus,
                spatial_bounds=bounds,
                data_context=metadata,
            )

        # 3. Charts & Quantitative Data
        if any(k in meta_type for k in ("chart", "bar", "doughnut", "pie", "metric")) or meta_type in ("line", "line_chart") or \
           any(k in role_lower for k in ("chart", "bar", "axis", "metric", "doughnut", "series")) or \
           any(k in name_lower for k in ("chart", "bar_", "axis", "metric_value", "metric_label")) or \
           (name_lower.startswith("line_") and "timeline" not in name_lower):
            
            if "axis" in name_lower or "axis" in role_lower:
                sem_role = SemanticRole.DATA_AXIS
                importance = NarrativeImportance.SUPPORTING
                hierarchy = 4
                focus = AudienceFocus.IMMEDIATE
            elif "value" in name_lower or "metric" in role_lower:
                sem_role = SemanticRole.METRIC_VALUE
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2
                focus = AudienceFocus.CLIMAX
            elif "doughnut" in name_lower or "slice" in name_lower:
                sem_role = SemanticRole.DATA_SLICE
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2
                focus = AudienceFocus.SEQUENTIAL
            elif "line" in name_lower and "timeline" not in name_lower:
                sem_role = SemanticRole.DATA_LINE
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2
                focus = AudienceFocus.SEQUENTIAL
            else:
                sem_role = SemanticRole.DATA_BAR
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2
                focus = AudienceFocus.SEQUENTIAL

            return ObjectSemanticProfile(
                object_id=obj_id,
                object_type=ObjectType.CHART,
                semantic_role=sem_role,
                narrative_importance=importance,
                visual_hierarchy_level=hierarchy,
                reading_order=index,
                audience_focus=focus,
                spatial_bounds=bounds,
                data_context=metadata,
            )

        # 3. Diagrams & Processes
        if any(k in meta_type for k in ("diagram", "flowchart", "process", "connector", "node")) or \
           any(k in role_lower for k in ("node", "connector", "process_step", "flowchart")) or \
           any(k in name_lower for k in ("node_", "step_", "connector_", "arrow_flow")):
            
            if "connector" in name_lower or "connector" in role_lower:
                sem_role = SemanticRole.CONNECTOR
                importance = NarrativeImportance.SUPPORTING
                hierarchy = 3
            else:
                sem_role = SemanticRole.PROCESS_STEP
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2

            return ObjectSemanticProfile(
                object_id=obj_id,
                object_type=ObjectType.DIAGRAM,
                semantic_role=sem_role,
                narrative_importance=importance,
                visual_hierarchy_level=hierarchy,
                reading_order=index,
                audience_focus=AudienceFocus.SEQUENTIAL,
                spatial_bounds=bounds,
                data_context=metadata,
            )

        # 4. Maps & Geospatial Elements
        if any(k in meta_type for k in ("map", "geo", "region", "route", "marker")) or \
           any(k in role_lower for k in ("map", "geo_region", "geo_marker", "geo_route", "location")) or \
           any(k in name_lower for k in ("map_", "region_", "route_", "pin_", "marker_")):
            
            if "route" in name_lower or "route" in role_lower:
                sem_role = SemanticRole.GEO_ROUTE
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2
            elif "marker" in name_lower or "pin" in name_lower or "marker" in role_lower:
                sem_role = SemanticRole.GEO_MARKER
                importance = NarrativeImportance.PRIMARY
                hierarchy = 1
            else:
                sem_role = SemanticRole.GEO_REGION
                importance = NarrativeImportance.SECONDARY
                hierarchy = 3

            return ObjectSemanticProfile(
                object_id=obj_id,
                object_type=ObjectType.MAP,
                semantic_role=sem_role,
                narrative_importance=importance,
                visual_hierarchy_level=hierarchy,
                reading_order=index,
                audience_focus=AudienceFocus.SEQUENTIAL,
                spatial_bounds=bounds,
                data_context=metadata,
            )


        # 6. Images & Cutouts
        if kind_lower in ("picture", "image") or getattr(node_or_dict, "media_path", None) or \
           any(k in role_lower for k in ("image", "photo", "cutout", "picture")) or \
           any(k in name_lower for k in ("img_", "photo_", "cutout_", "picture_")):
            
            if "background" in name_lower or "bg" in name_lower:
                sem_role = SemanticRole.BACKGROUND_IMAGE
                importance = NarrativeImportance.BACKGROUND
                hierarchy = 5
                focus = AudienceFocus.IMMEDIATE
            elif "cutout" in name_lower or "hero" in name_lower:
                sem_role = SemanticRole.FOREGROUND_CUTOUT
                importance = NarrativeImportance.HERO
                hierarchy = 1
                focus = AudienceFocus.SEQUENTIAL
            else:
                sem_role = SemanticRole.HERO_SUBJECT
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2
                focus = AudienceFocus.SEQUENTIAL

            return ObjectSemanticProfile(
                object_id=obj_id,
                object_type=ObjectType.IMAGE,
                semantic_role=sem_role,
                narrative_importance=importance,
                visual_hierarchy_level=hierarchy,
                reading_order=index,
                audience_focus=focus,
                spatial_bounds=bounds,
                data_context=metadata,
            )

        # 7. Typography (Headlines, Subtitles, Paragraphs)
        if kind_lower == "textbox" or bool(text and not kind_lower in ("picture", "group")):
            if font_size >= 26.0 or "title" in role_lower or "title" in name_lower:
                sem_role = SemanticRole.TITLE
                importance = NarrativeImportance.PRIMARY
                hierarchy = 1
                focus = AudienceFocus.IMMEDIATE
            elif font_size >= 15.0 or "subtitle" in role_lower or "subtitle" in name_lower:
                sem_role = SemanticRole.SUBTITLE
                importance = NarrativeImportance.SECONDARY
                hierarchy = 2
                focus = AudienceFocus.IMMEDIATE
            elif "keyword" in role_lower or "keyword" in name_lower:
                sem_role = SemanticRole.KEYWORD
                importance = NarrativeImportance.PRIMARY
                hierarchy = 2
                focus = AudienceFocus.CLIMAX
            else:
                sem_role = SemanticRole.BODY_TEXT
                importance = NarrativeImportance.SECONDARY
                hierarchy = 3
                focus = AudienceFocus.SEQUENTIAL

            return ObjectSemanticProfile(
                object_id=obj_id,
                object_type=ObjectType.TYPOGRAPHY,
                semantic_role=sem_role,
                narrative_importance=importance,
                visual_hierarchy_level=hierarchy,
                reading_order=index,
                audience_focus=focus,
                spatial_bounds=bounds,
                data_context={"text": text, "font_size": font_size, **metadata},
            )

        # 8. Container Cards (Glassmorphism / Background Scrim)
        if "card" in role_lower or "card" in name_lower or "container" in role_lower or kind_lower == "card":
            return ObjectSemanticProfile(
                object_id=obj_id,
                object_type=ObjectType.CONTAINER_CARD,
                semantic_role=SemanticRole.CONTAINER,
                narrative_importance=NarrativeImportance.BACKGROUND,
                visual_hierarchy_level=4,
                reading_order=index,
                audience_focus=AudienceFocus.IMMEDIATE,
                spatial_bounds=bounds,
                data_context=metadata,
            )

        # Default fallback: Decorative / Supporting
        return ObjectSemanticProfile(
            object_id=obj_id,
            object_type=ObjectType.DECORATIVE,
            semantic_role=SemanticRole.ACCENT,
            narrative_importance=NarrativeImportance.BACKGROUND,
            visual_hierarchy_level=5,
            reading_order=index,
            audience_focus=AudienceFocus.IMMEDIATE,
            spatial_bounds=bounds,
            data_context=metadata,
        )
