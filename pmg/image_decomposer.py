"""Image -> shared LayerTree. Never builds a completed presentation.

Flat-color semantics, exact geometric partitions and explicitly approximate depth
layers use shared segmentation. Complex semantic/depth models can be injected;
offline heuristics are identified as such. Raster layers are never called vectors.
"""
import hashlib
import math
from pathlib import Path
from typing import Dict, List, Optional

import cv2
import numpy as np
from PIL import Image

from pmg.layers import LayerMask, LayerTree, VisualObject
from pmg.segmentation import ColorRegionSegmenter, ForegroundSegmenter, Segment


def _load_rgba(path):
    with Image.open(path) as image:
        return np.asarray(image.convert('RGBA')).copy()


def _paths(mask, bounds):
    x, y, width, height = bounds
    contours, hierarchy = cv2.findContours(mask[y:y+height, x:x+width], cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    paths, holes = [], []
    if hierarchy is not None:
        for index, contour in enumerate(contours):
            simplified = cv2.approxPolyDP(contour, .35, True).reshape(-1, 2)
            if len(simplified) < 3:
                continue
            paths.append([(float(px)/max(1, width-1), float(py)/max(1, height-1)) for px, py in simplified])
            holes.append(bool(hierarchy[0][index][3] >= 0))
    return paths, holes


def _color_role(rgb):
    red, green, blue = rgb
    if min(rgb) > 215:
        return 'White_Separator'
    if green > red*1.12 and green > blue*1.1:
        return 'Green_Background'
    if red > green*1.3 and red > blue*1.2:
        return 'Red_Diagonal'
    if red > 130 and green > 85 and blue < min(red, green)*.8:
        return 'Wheat'
    return 'Color_Region'


def _render_paths(paths, width, height):
    mask = np.zeros((height, width), np.uint8)
    contours = [np.rint(np.asarray(path)*np.asarray([max(1, width-1), max(1, height-1)])).astype(np.int32)
                for path in paths]
    if contours:
        cv2.drawContours(mask, contours, -1, 255, cv2.FILLED)
    return mask


def _granular_wheat(segments):
    result = []
    for segment in segments:
        if segment.label not in ('region', 'Wheat') or _color_role(segment.color_rgb) != 'Wheat':
            result.append(segment)
            continue
        core = cv2.erode(segment.mask, np.ones((3, 3), np.uint8))
        count, labels = cv2.connectedComponents(core, connectivity=8)
        if count <= 2:
            result.append(segment)
            continue
        # Nearest eroded core assigns all original pixels, retaining exact coverage.
        _, nearest = cv2.distanceTransformWithLabels(np.uint8(core == 0)*255, cv2.DIST_L2, 3,
                                                    labelType=cv2.DIST_LABEL_CCOMP)
        for index in range(1, count):
            mask = np.uint8((nearest == index) & (segment.mask > 0))*255
            points = cv2.findNonZero(mask)
            if points is None:
                continue
            bounds = cv2.boundingRect(points)
            result.append(Segment(f'{segment.segment_id}_grain_{index}', bounds, mask, segment.color_rgb,
                                  confidence=.7, inferred=True,
                                  metadata={'area': int((mask > 0).sum()), 'strategy': 'morphological_grain_partition'}))
    return result


class ImageDecomposer:
    def __init__(self, segmenter=None, depth_estimator=None):
        self.segmenter = segmenter or ColorRegionSegmenter()
        self.depth_estimator = depth_estimator

    def decompose(self, source_path: str, output_dir: str, mode: str = 'semantic',
                  pattern: str = 'vertical', count: int = 8, rows: int = 3, columns: int = 4,
                  granular: bool = True, polygons=None, source_asset_id: Optional[str] = None,
                  similarity_tolerance: float = .97) -> LayerTree:
        rgba = _load_rgba(source_path)
        height, width = rgba.shape[:2]
        folder = Path(output_dir).resolve()
        folder.mkdir(parents=True, exist_ok=True)
        digest = hashlib.sha256(Path(source_path).read_bytes()).hexdigest()[:12]
        root_id = f'asset_{digest}'
        root = VisualObject(root_id, 'group', (0, 0, width, height), editability='group', semantic_role='asset_root')
        tree = LayerTree(root_id, str(Path(source_path).resolve()), (width, height), [root],
                         source_asset_id=source_asset_id or f'sha256:{digest}',
                         metadata={'mode': mode, 'pattern': pattern if mode == 'geometric' else None,
                                   'semantic_method': type(self.segmenter).__name__ if mode == 'semantic' else None})
        bgr = cv2.cvtColor(rgba, cv2.COLOR_RGBA2BGRA)
        if mode == 'semantic':
            segments = self.segmenter.segment(bgr)
            if granular:
                segments = _granular_wheat(segments)
            tree.metadata['inference'] = 'Palette/color roles are deterministic heuristics, not model-generated semantics.'
        elif mode == 'geometric':
            segments = self._geometric_segments(rgba, pattern, count, rows, columns, polygons)
        elif mode == 'depth':
            segments = self._depth_segments(rgba)
            tree.metadata['depth_method'] = 'injected_estimator' if self.depth_estimator else 'vertical_and_saliency_heuristic'
            tree.metadata['inference'] = 'Depth ordering is estimated; occluded content has not been reconstructed.'
        else:
            raise ValueError(f'Unknown decomposition mode {mode!r}')
        if not segments:
            raise ValueError('No visual components were recovered')

        covered = np.zeros((height, width), bool)
        for segment in segments:
            covered |= segment.mask > 0
        residue = np.uint8((~covered) & (rgba[:, :, 3] > 0))*255
        if residue.any():
            # Preserve tiny unclassified/antialiased details as an honest raster
            # layer, rather than dropping pixels or flattening all components.
            if float((residue > 0).mean()) > .1:
                raise ValueError('Segmentation missed substantial content; refine the strategy before production')
            sample = rgba[:, :, :3][residue > 0]
            segments.append(Segment('detail_residual', cv2.boundingRect(cv2.findNonZero(residue)), residue,
                                    tuple(int(v) for v in np.median(sample, axis=0)), label='Detail_Residual',
                                    metadata={'strategy': 'unclassified_pixel_preservation'}))

        wheat_group = None
        role_counts: Dict[str, int] = {}
        for index, segment in enumerate(segments):
            x, y, component_w, component_h = segment.bbox
            if segment.mask.shape != (height, width):
                raise ValueError('Segment mask must use full source pixel coordinates')
            if mode == 'semantic':
                role = segment.label if segment.label != 'region' else _color_role(segment.color_rgb)
            else:
                role = segment.label if segment.label != 'region' else f'{pattern}_segment'
            if role == 'Wheat':
                if wheat_group is None:
                    wheat_group = VisualObject(root_id+'_wheat', 'group', (0, 0, width, height), parent_id=root_id,
                                               z_order=100, semantic_role='wheat_group', editability='group')
                    tree.objects.append(wheat_group)
                    root.children.append(wheat_group.object_id)
                role = 'Stem' if component_h > 3.5*component_w else 'Grain'
            role_counts[role] = role_counts.get(role, 0)+1
            name = f'{role}_{role_counts[role]:02d}'
            identity = f'{root_id}_{name}'
            mask_path = folder / f'{identity}_mask.png'
            Image.fromarray(segment.mask).save(mask_path)
            mask_id = identity+'_mask'
            tree.masks.append(LayerMask(mask_id, str(mask_path), (width, height), inferred=segment.inferred))
            sample = rgba[:, :, :3][segment.mask > 0]
            is_vector = mode == 'semantic' and sample.size > 0 and float(np.std(sample.astype(float), axis=0).max()) < 14
            paths, holes = _paths(segment.mask, segment.bbox)
            is_vector = is_vector and bool(paths)
            if is_vector:
                vector_mask = _render_paths(paths, component_w, component_h) > 0
                original_mask = segment.mask[y:y+component_h, x:x+component_w] > 0
                intersection = (vector_mask & original_mask).sum()
                union = (vector_mask | original_mask).sum()
                is_vector = float(intersection/max(1, union)) >= .995
            parent = wheat_group if name.startswith(('Stem_', 'Grain_')) and wheat_group else root
            if is_vector:
                kind, editability, media = 'polygon', 'native_vector', None
            else:
                kind, editability = 'image', 'raster_layer'
                layer = rgba.copy()
                layer[:, :, 3] = np.where(segment.mask > 0, layer[:, :, 3], 0)
                media_path = folder / f'{identity}.png'
                Image.fromarray(layer[y:y+component_h, x:x+component_w]).save(media_path)
                media = str(media_path)
            obj = VisualObject(identity, kind, (x, y, component_w, component_h), parent_id=parent.object_id,
                               z_order=index, layer_id=name, asset_ref=tree.source_asset_id,
                               media_path=media, mask_id=mask_id, paths=paths, path_holes=holes,
                               fill=''.join(f'{channel:02X}' for channel in segment.color_rgb),
                               editability=editability, semantic_role=role.lower(), inferred=segment.inferred,
                               metadata={'segmentation': segment.metadata, 'confidence': segment.confidence})
            parent.children.append(identity)
            tree.objects.append(obj)
        tree.reassembly = verify_reassembly(tree, str(folder/'reassembled.png'), tolerance=similarity_tolerance)
        tree.save(str(folder/'layer_tree.json'))
        return tree

    @staticmethod
    def _geometric_segments(rgba, pattern, count, rows, columns, polygons):
        height, width = rgba.shape[:2]
        yy, xx = np.indices((height, width))
        if count < 1 or rows < 1 or columns < 1:
            raise ValueError('Partition counts must be positive')
        if pattern in ('vertical', 'vertical_strips'):
            labels = xx*count//width
        elif pattern in ('horizontal', 'horizontal_strips'):
            labels = yy*count//height
        elif pattern in ('grid', 'tiles', 'puzzle'):
            labels = (yy*rows//height)*columns + xx*columns//width
            if pattern == 'puzzle':
                radius = min(width/columns, height/rows)*.17
                for row in range(rows):
                    for column in range(columns-1):
                        cx, cy = (column+1)*width/columns, (row+.5)*height/rows
                        owner = row*columns+column + ((row+column) % 2)
                        labels[(xx-cx)**2+(yy-cy)**2 < radius**2] = owner
                for row in range(rows-1):
                    for column in range(columns):
                        cx, cy = (column+.5)*width/columns, (row+1)*height/rows
                        owner = row*columns+column + (columns if (row+column) % 2 else 0)
                        labels[(xx-cx)**2+(yy-cy)**2 < radius**2] = owner
        elif pattern in ('radial', 'sectors', 'circular_sectors'):
            angle = (np.arctan2(yy-(height-1)/2, xx-(width-1)/2)+2*math.pi) % (2*math.pi)
            labels = np.minimum(count-1, (angle/(2*math.pi)*count).astype(int))
        elif pattern == 'triangles':
            cell_x, cell_y = xx*columns//width, yy*rows//height
            local_x, local_y = xx*columns/width-cell_x, yy*rows/height-cell_y
            labels = ((cell_y*columns+cell_x)*2 + (local_x+local_y >= 1))
        elif pattern in ('polygon', 'custom'):
            if not polygons:
                raise ValueError('Custom polygon decomposition requires clipping polygons')
            labels = np.full((height, width), -1, np.int32)
            for index, polygon in enumerate(polygons):
                mask = np.zeros((height, width), np.uint8)
                cv2.fillPoly(mask, [np.asarray(polygon, np.int32)], 255)
                labels[(mask > 0) & (labels < 0)] = index
            # Retain uncovered pixels explicitly rather than silently dropping content.
            labels[labels < 0] = len(polygons)
        else:
            raise ValueError(f'Unsupported geometric slicing pattern: {pattern}')
        segments = []
        for index in np.unique(labels):
            mask = np.uint8(labels == index)*255
            bounds = cv2.boundingRect(cv2.findNonZero(mask))
            rgb = tuple(int(v) for v in np.median(rgba[:, :, :3][mask > 0], axis=0))
            segments.append(Segment(f'segment_{index}', bounds, mask, rgb,
                                    label=f'{pattern}_piece', metadata={'partition': int(index)}))
        return segments

    def _depth_segments(self, rgba):
        height, width = rgba.shape[:2]
        if self.depth_estimator:
            depth = np.asarray(self.depth_estimator(rgba))
            if depth.shape != (height, width) or not np.isfinite(depth).all():
                raise ValueError('Depth provider must return a finite source-sized depth map')
            thresholds = np.quantile(depth, [1/3, 2/3])
            labels = np.digitize(depth, thresholds)
        else:
            labels = np.minimum(2, np.indices((height, width))[0]*3//height)
            bgr = cv2.cvtColor(rgba[:, :, :3], cv2.COLOR_RGB2BGR)
            for segment in ForegroundSegmenter().segment(bgr):
                # Salient subject regions are only an estimated foreground hypothesis.
                labels[segment.mask > 0] = 2
        output = []
        for index, role in enumerate(('Background', 'Midground', 'Foreground')):
            mask = np.uint8(labels == index)*255
            points = cv2.findNonZero(mask)
            if points is not None:
                color = tuple(int(v) for v in np.median(rgba[:, :, :3][mask > 0], axis=0))
                output.append(Segment(role.lower(), cv2.boundingRect(points), mask, color,
                                      label=role, inferred=True, confidence=.5,
                                      metadata={'estimated_depth_layer': index}))
        return output


def render_layer_tree(tree: LayerTree) -> np.ndarray:
    width, height = tree.source_size_px
    canvas = Image.new('RGBA', (width, height))
    for obj in sorted(tree.objects, key=lambda item: item.z_order):
        if obj.kind == 'group':
            continue
        x, y, w, h = map(int, obj.geometry)
        if obj.kind == 'polygon':
            mask = _render_paths(obj.paths, w, h)
            color = tuple(int(obj.fill[index:index+2], 16) for index in (0, 2, 4))
            layer = np.zeros((h, w, 4), np.uint8)
            layer[:, :, :3], layer[:, :, 3] = color, mask
            layer = Image.fromarray(layer)
        elif obj.kind == 'image':
            layer = Image.open(obj.media_path).convert('RGBA')
        else:
            continue
        canvas.alpha_composite(layer, (x, y))
    return np.asarray(canvas)


def verify_reassembly(tree: LayerTree, output_path: str, tolerance: float = .97):
    original = _load_rgba(tree.source_path)
    assembled = render_layer_tree(tree)
    Image.fromarray(assembled).save(output_path)
    similarity = float(1-np.abs(original.astype(float)-assembled.astype(float)).mean()/255)
    coverage = float(((assembled[:, :, 3] > 0) | (original[:, :, 3] == 0)).mean())
    source_edges = cv2.Canny(original[:, :, :3], 60, 140)
    assembled_edges = cv2.Canny(assembled[:, :, :3], 60, 140)
    nearby = cv2.dilate(assembled_edges, np.ones((3, 3), np.uint8)) > 0
    contour_alignment = float(nearby[source_edges > 0].mean()) if source_edges.any() else 1.0
    return {'status': 'PASS' if similarity >= tolerance and coverage >= .995 else 'FAIL',
            'pixel_similarity': similarity, 'coverage_ratio': coverage, 'contour_alignment': contour_alignment,
            'tolerance': tolerance, 'rendered_path': str(Path(output_path).resolve()),
            'raster_only_components': [obj.object_id for obj in tree.objects if obj.editability.startswith('raster')],
            'inferred_components': [obj.object_id for obj in tree.objects if obj.inferred]}
