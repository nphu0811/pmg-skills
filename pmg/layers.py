"""Shared asset/layer objects, reused by decomposition and future Scene Graphs.

Coordinates explicitly declare pixels or inches. Parent IDs/children, stable
identity, masks, z-order and editability survive serialization; this is not a
presentation generator or an animation compiler.
"""
import json
import math
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from pmg.models import AuditReport, ReviewDefect


@dataclass
class VisualObject:
    object_id: str
    kind: str
    geometry: Tuple[float, float, float, float]
    parent_id: Optional[str] = None
    children: List[str] = field(default_factory=list)
    units: str = 'px'
    z_order: int = 0
    layer_id: str = ''
    asset_ref: Optional[str] = None
    media_path: Optional[str] = None
    mask_id: Optional[str] = None
    # Each contour is normalized to the object's bounds; opposite winding cuts holes.
    paths: List[List[Tuple[float, float]]] = field(default_factory=list)
    path_holes: List[bool] = field(default_factory=list)
    fill: str = 'FFFFFF'
    stroke: Optional[str] = None
    stroke_width: float = 0.0
    text: str = ''
    opacity: float = 1.0
    rotation_deg: float = 0.0
    scale: Tuple[float, float] = (1.0, 1.0)
    pivot: Tuple[float, float] = (0.5, 0.5)
    crop: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
    editability: str = 'native_vector'
    semantic_role: str = 'content'
    inferred: bool = False
    metadata: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        if not self.layer_id:
            self.layer_id = self.object_id

    def to_dict(self):
        return asdict(self)


@dataclass
class LayerMask:
    mask_id: str
    path: str
    size_px: Tuple[int, int]
    kind: str = 'alpha'
    inferred: bool = False

    def to_dict(self):
        return asdict(self)


@dataclass
class LayerTree:
    root_id: str
    source_path: str
    source_size_px: Tuple[int, int]
    objects: List[VisualObject] = field(default_factory=list)
    masks: List[LayerMask] = field(default_factory=list)
    source_asset_id: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    reassembly: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self):
        return asdict(self)

    def save(self, path: str):
        destination = Path(path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(json.dumps(self.to_dict(), indent=2), encoding='utf-8')
        return str(destination.resolve())

    @classmethod
    def from_dict(cls, data):
        values = dict(data)
        values['objects'] = [VisualObject(**item) for item in data.get('objects', [])]
        values['masks'] = [LayerMask(**item) for item in data.get('masks', [])]
        tree = cls(**values)
        report = tree.audit()
        if not report['is_valid']:
            raise ValueError('; '.join(item['description'] for item in report['defects']))
        return tree

    @classmethod
    def load(cls, path):
        return cls.from_dict(json.loads(Path(path).read_text(encoding='utf-8')))

    def audit(self):
        defects = []
        def error(identity, message):
            defects.append(ReviewDefect('layer_integrity', 'critical', 0, identity, message))
        objects = {obj.object_id: obj for obj in self.objects}
        masks = {mask.mask_id: mask for mask in self.masks}
        if len(objects) != len(self.objects) or any(not obj.object_id for obj in self.objects):
            error(self.root_id, 'Layer object IDs must be non-empty and unique.')
        if self.root_id not in objects or objects[self.root_id].parent_id is not None:
            error(self.root_id, 'Layer root is missing or has a parent.')
        if len(masks) != len(self.masks):
            error(self.root_id, 'Duplicate mask identity.')
        for obj in self.objects:
            if obj.kind not in ('group', 'rect', 'ellipse', 'polygon', 'image', 'text', 'video', 'model3d', 'connector'):
                error(obj.object_id, f'Unknown visual object kind {obj.kind}.')
            if obj.units not in ('px', 'in') or len(obj.geometry) != 4 or not all(math.isfinite(x) for x in obj.geometry):
                error(obj.object_id, 'Invalid geometry or coordinate units.')
            elif obj.geometry[2] <= 0 or obj.geometry[3] <= 0:
                error(obj.object_id, 'Object width and height must be positive.')
            if not 0 <= obj.opacity <= 1:
                error(obj.object_id, 'Opacity is outside [0,1].')
            if obj.kind in ('image', 'video', 'model3d') and (not obj.media_path or not Path(obj.media_path).is_file()):
                error(obj.object_id, 'Media asset is missing.')
            if obj.paths and len(obj.paths) != len(obj.path_holes):
                error(obj.object_id, 'Vector contour/hole declarations disagree.')
            if obj.mask_id and obj.mask_id not in masks:
                error(obj.object_id, f'Missing mask {obj.mask_id}.')
            if obj.parent_id is not None:
                parent = objects.get(obj.parent_id)
                if parent is None or obj.object_id not in parent.children:
                    error(obj.object_id, 'Parent/child references disagree.')
            if len(obj.children) != len(set(obj.children)):
                error(obj.object_id, 'Repeated child identity.')
            for child_id in obj.children:
                if child_id not in objects or objects[child_id].parent_id != obj.object_id:
                    error(obj.object_id, f'Invalid child reference {child_id}.')
            ancestry = set()
            current = obj
            while current is not None:
                if current.object_id in ancestry:
                    error(obj.object_id, 'Cycle in layer hierarchy.')
                    break
                ancestry.add(current.object_id)
                current = objects.get(current.parent_id)
            if obj.object_id != self.root_id and self.root_id not in ancestry:
                error(obj.object_id, 'Object is detached from the layer root.')
        for mask in self.masks:
            if not Path(mask.path).is_file():
                error(mask.mask_id, 'Mask asset is missing.')
        return AuditReport(defects, {'object_count': len(self.objects), 'mask_count': len(self.masks)}).to_dict()
