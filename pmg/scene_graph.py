"""Canonical Scene Graph data model for PMG Skills.

Hierarchical scene composition supporting nested groups, local and world transforms,
pivot points, intentional overlap, protected readability regions, and compilation
to native PowerPoint shapes.
"""
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Tuple, Any, Set
import math
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from lxml import etree

A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'


@dataclass
class Transform:
    """2D spatial transform with depth and pivot."""
    x: float = 0.0          # Left position (inches)
    y: float = 0.0          # Top position (inches)
    width: float = 1.0      # Width (inches)
    height: float = 1.0     # Height (inches)
    rotation_deg: float = 0.0
    scale_x: float = 1.0
    scale_y: float = 1.0
    opacity: float = 1.0
    z_index: int = 0
    pivot_x: float = 0.5    # Normalized pivot [0.0, 1.0]
    pivot_y: float = 0.5

    def bounds(self) -> Tuple[float, float, float, float]:
        return (self.x, self.y, self.width * self.scale_x, self.height * self.scale_y)


@dataclass
class SceneNode:
    """Individual object or group in the Scene Graph hierarchy."""
    node_id: str
    name: str = ""
    kind: str = "rect"  # 'group', 'rect', 'ellipse', 'polygon', 'picture', 'textbox'
    transform: Transform = field(default_factory=Transform)
    parent_id: Optional[str] = None
    children: List[str] = field(default_factory=list)
    semantic_role: str = "content"  # 'background', 'subject', 'hud', 'typography', 'overlay', 'card'
    overlap_policy: str = "allow_overlap"  # 'allow_overlap' vs 'protected_readable'
    editability: str = "native_vector"     # 'native_vector', 'raster_layer', 'text', 'group'
    fill: str = "FFFFFF"
    stroke: Optional[str] = None
    stroke_width: float = 0.0
    text: str = ""
    font_size_pt: float = 18.0
    text_color: str = "0F172A"
    bold: bool = False
    media_path: Optional[str] = None
    crop: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    paths: List[List[Tuple[float, float]]] = field(default_factory=list)  # normalized coords for custom polygon
    metadata: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        if not self.name:
            self.name = self.node_id

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class SceneGraph:
    """Hierarchical graph of visual nodes for a single presentation scene."""
    def __init__(self, scene_id: str, width_in: float = 13.333, height_in: float = 7.5):
        self.scene_id = scene_id
        self.width_in = width_in
        self.height_in = height_in
        self.nodes: Dict[str, SceneNode] = {}
        self.root_nodes: List[str] = []

    def add_node(self, node: SceneNode, parent_id: Optional[str] = None) -> SceneNode:
        node.parent_id = parent_id
        self.nodes[node.node_id] = node
        if parent_id is not None:
            if parent_id in self.nodes:
                if node.node_id not in self.nodes[parent_id].children:
                    self.nodes[parent_id].children.append(node.node_id)
            else:
                raise ValueError(f"Parent node '{parent_id}' does not exist in SceneGraph.")
        else:
            if node.node_id not in self.root_nodes:
                self.root_nodes.append(node.node_id)
        return node

    def get_world_transform(self, node_id: str) -> Transform:
        """Compute accumulated world transform up to the scene root."""
        if node_id not in self.nodes:
            raise KeyError(f"Node '{node_id}' not found.")
        node = self.nodes[node_id]
        x = node.transform.x
        y = node.transform.y
        w = node.transform.width
        h = node.transform.height
        rot = node.transform.rotation_deg
        sx = node.transform.scale_x
        sy = node.transform.scale_y
        op = node.transform.opacity
        z = node.transform.z_index

        curr = node
        while curr.parent_id is not None:
            parent = self.nodes[curr.parent_id]
            x = parent.transform.x + x * parent.transform.scale_x
            y = parent.transform.y + y * parent.transform.scale_y
            sx *= parent.transform.scale_x
            sy *= parent.transform.scale_y
            rot = (rot + parent.transform.rotation_deg) % 360.0
            op *= parent.transform.opacity
            z += parent.transform.z_index
            curr = parent

        return Transform(
            x=round(x, 4),
            y=round(y, 4),
            width=round(w, 4),
            height=round(h, 4),
            rotation_deg=round(rot, 2),
            scale_x=round(sx, 4),
            scale_y=round(sy, 4),
            opacity=round(op, 3),
            z_index=z,
        )

    def validate_graph(self) -> List[str]:
        """Validate parent-child cycles, dangling children, and protected collisions."""
        errors = []
        visited = set()
        stack = set()

        def detect_cycle(curr_id: str):
            visited.add(curr_id)
            stack.add(curr_id)
            for child_id in self.nodes[curr_id].children:
                if child_id not in self.nodes:
                    errors.append(f"Node '{curr_id}' references non-existent child '{child_id}'")
                    continue
                if child_id not in visited:
                    detect_cycle(child_id)
                elif child_id in stack:
                    errors.append(f"Cycle detected in SceneGraph involving '{curr_id}' and '{child_id}'")
            stack.remove(curr_id)

        for root in self.root_nodes:
            if root not in visited:
                detect_cycle(root)

        # Collision detection: Flag ONLY accidental collisions among protected readable regions
        readable_nodes = [
            (nid, self.nodes[nid], self.get_world_transform(nid))
            for nid in self.nodes
            if self.nodes[nid].overlap_policy == "protected_readable" and self.nodes[nid].kind == "textbox"
        ]
        for i in range(len(readable_nodes)):
            id1, n1, t1 = readable_nodes[i]
            b1 = t1.bounds()
            for j in range(i + 1, len(readable_nodes)):
                id2, n2, t2 = readable_nodes[j]
                b2 = t2.bounds()
                # Check bounding box overlap
                if not (b1[0] + b1[2] <= b2[0] or b2[0] + b2[2] <= b1[0] or
                        b1[1] + b1[3] <= b2[1] or b2[1] + b2[3] <= b1[1]):
                    errors.append(f"Accidental collision between protected readable text nodes '{id1}' and '{id2}'")

        return errors

    def compile_to_slide(self, slide, name_prefix: str = "!!") -> Dict[str, Any]:
        """Compile the hierarchical Scene Graph into native editable PowerPoint shapes."""
        errors = self.validate_graph()
        if errors:
            raise ValueError("Invalid SceneGraph: " + "; ".join(errors))

        mapping: Dict[str, Any] = {}

        def _custom_geom(shape, paths):
            properties = shape._element.spPr
            for node in list(properties):
                if node.tag in (f'{{{A_NS}}}prstGeom', f'{{{A_NS}}}custGeom'):
                    properties.remove(node)
            geometry = etree.Element(f'{{{A_NS}}}custGeom')
            for tag in ('avLst', 'gdLst', 'ahLst', 'cxnLst'):
                etree.SubElement(geometry, f'{{{A_NS}}}{tag}')
            path_list = etree.SubElement(geometry, f'{{{A_NS}}}pathLst')
            path_node = etree.SubElement(path_list, f'{{{A_NS}}}path', w='100000', h='100000')
            for points in paths:
                for idx, (px, py) in enumerate(points):
                    cmd = etree.SubElement(path_node, f'{{{A_NS}}}{"moveTo" if idx == 0 else "lnTo"}')
                    etree.SubElement(cmd, f'{{{A_NS}}}pt', x=str(round(px * 100000)), y=str(round(py * 100000)))
                etree.SubElement(path_node, f'{{{A_NS}}}close')
            properties.insert(1 if properties.find(f'{{{A_NS}}}xfrm') is not None else 0, geometry)

        def compile_node(node: SceneNode, collection):
            wt = self.get_world_transform(node.node_id)
            geom = [Inches(wt.x), Inches(wt.y), Inches(max(0.01, wt.width * wt.scale_x)), Inches(max(0.01, wt.height * wt.scale_y))]
            shape_name = f"{name_prefix}{node.name}" if not node.name.startswith(name_prefix) else node.name

            if node.kind == "group":
                shape = collection.add_group_shape()
                shape.name = shape_name
                mapping[node.node_id] = shape
                for child_id in node.children:
                    compile_node(self.nodes[child_id], shape.shapes)
                return

            if node.kind == "picture" and node.media_path:
                shape = collection.add_picture(node.media_path, *geom)
                shape.crop_left, shape.crop_top, shape.crop_right, shape.crop_bottom = node.crop
            elif node.kind == "textbox":
                shape = collection.add_textbox(*geom)
                shape.text_frame.text = node.text
                for p in shape.text_frame.paragraphs:
                    p.font.size = Pt(node.font_size_pt)
                    p.font.color.rgb = RGBColor.from_string(node.text_color.lstrip("#"))
                    p.font.bold = node.bold
            elif node.kind in ("rect", "ellipse", "polygon"):
                mso_type = MSO_SHAPE.OVAL if node.kind == "ellipse" else MSO_SHAPE.RECTANGLE
                shape = collection.add_shape(mso_type, *geom)
                if node.kind == "polygon" and node.paths:
                    _custom_geom(shape, node.paths)
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor.from_string(node.fill.lstrip("#"))
                if wt.opacity < 1.0:
                    color = shape._element.spPr.find(f'{{{A_NS}}}solidFill/{{{A_NS}}}srgbClr')
                    if color is not None:
                        etree.SubElement(color, f'{{{A_NS}}}alpha', val=str(round(wt.opacity * 100000)))
                if node.stroke:
                    shape.line.color.rgb = RGBColor.from_string(node.stroke.lstrip("#"))
                    shape.line.width = Pt(node.stroke_width)
                else:
                    shape.line.fill.background()
            else:
                shape = collection.add_shape(MSO_SHAPE.RECTANGLE, *geom)

            shape.name = shape_name
            shape.rotation = wt.rotation_deg
            mapping[node.node_id] = shape

        # Sort roots by z_index
        sorted_roots = sorted(self.root_nodes, key=lambda nid: self.nodes[nid].transform.z_index)
        for root_id in sorted_roots:
            compile_node(self.nodes[root_id], slide.shapes)

        return mapping
