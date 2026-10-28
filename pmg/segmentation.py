"""Shared segmentation strategies for still images and video reconstruction.

The existing video foreground algorithm is moved here unchanged in behavior.
Both consumers accept the same Segmenter protocol and structured Segment result.
Color regions are an efficient flat-art strategy, not claimed AI semantics.
"""
import math
from dataclasses import dataclass, field
from typing import Any, Dict, List, Protocol, Tuple

import cv2
import numpy as np


@dataclass
class Segment:
    segment_id: str
    bbox: Tuple[int, int, int, int]
    mask: np.ndarray
    color_rgb: Tuple[int, int, int]
    label: str = 'region'
    confidence: float = 1.0
    inferred: bool = False
    metadata: Dict[str, Any] = field(default_factory=dict)


class Segmenter(Protocol):
    def segment(self, image: np.ndarray, **context) -> List[Segment]:
        """Return independent regions in source pixel coordinates."""
        ...


def box_iou(first, second):
    x1, y1, w1, h1 = first
    x2, y2, w2, h2 = second
    intersection = max(0, min(x1+w1, x2+w2)-max(x1, x2)) * max(0, min(y1+h1, y2+h2)-max(y1, y2))
    return intersection / max(1, w1*h1 + w2*h2 - intersection)


class ColorRegionSegmenter:
    """Connected palette regions for logos/flat illustrations, without AI weights."""
    def __init__(self, colors: int = 6, min_area: int = 3):
        self.colors, self.min_area = colors, min_area

    def segment(self, image: np.ndarray, **context) -> List[Segment]:
        rgb = cv2.cvtColor(image[:, :, :3], cv2.COLOR_BGR2RGB)
        pixels = rgb.reshape(-1, 3)
        opaque = image[:, :, 3].reshape(-1) > 0 if image.shape[2] == 4 else np.ones(len(pixels), bool)
        if not opaque.any():
            return []
        # Work on color frequencies; flat artwork needs no expensive per-pixel model.
        quantized = (pixels[opaque] // 16) * 16
        palette, counts = np.unique(quantized, axis=0, return_counts=True)
        dominant = palette[np.argsort(counts)[::-1][:self.colors]].astype(np.float32) + 7.5
        labels = np.empty(len(pixels), np.int32)
        for start in range(0, len(pixels), 65536):
            block = pixels[start:start+65536].astype(np.float32)
            distances = ((block[:, None, :] - dominant[None, :, :]) ** 2).sum(axis=2)
            labels[start:start+len(block)] = distances.argmin(axis=1)
        labels = labels.reshape(image.shape[:2])
        alpha = opaque.reshape(image.shape[:2])
        segments = []
        for color_index in range(len(dominant)):
            selected = ((labels == color_index) & alpha).astype(np.uint8)
            count, components, stats, _ = cv2.connectedComponentsWithStats(selected, connectivity=8)
            for index in range(1, count):
                x, y, width, height, area = map(int, stats[index])
                if area < self.min_area:
                    continue
                mask = np.uint8(components == index) * 255
                color = tuple(int(round(value)) for value in np.median(rgb[mask > 0], axis=0))
                segments.append(Segment(f'color_{color_index}_region_{index}', (x, y, width, height), mask, color,
                                        metadata={'area': area, 'strategy': 'palette_connected_components'}))
        return sorted(segments, key=lambda item: (-item.metadata['area'], item.bbox))


class ForegroundSegmenter:
    """Existing video saliency/multi-threshold algorithm, shared with still images."""
    def segment(self, image: np.ndarray, **context) -> List[Segment]:
        background = np.asarray(context.get('background_bgr', np.median(image.reshape(-1, 3), axis=0)), dtype=np.int16)
        boxes = self.foreground_boxes(image, background, image.shape[1], image.shape[0])
        output = []
        for index, box in enumerate(boxes):
            x, y, width, height = box['bbox']
            mask = np.zeros(image.shape[:2], np.uint8)
            mask[y:y+height, x:x+width] = 255
            output.append(Segment(f'foreground_{index}', box['bbox'], mask, box['dominant_rgb'],
                                  confidence=.65, inferred=True, metadata=box))
        return output

    @staticmethod
    def foreground_boxes(frame, bg_bgr, vid_w, vid_h):
        diff = np.abs(frame.astype(np.int16) - bg_bgr).max(axis=2).astype(np.uint8)
        min_area = (vid_w * vid_h) * 0.0012
        max_area = (vid_w * vid_h) * 0.55

        boxes: List[Dict[str, Any]] = []

        # Multi-threshold pass so both flat-background UI videos and textured videos segment cleanly
        for thresh in (24, 52, 85, 115):
            _, mask = cv2.threshold(diff, thresh, 255, cv2.THRESH_BINARY)
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (7, 7))
            closed = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)
            contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

            for cnt in contours:
                area = cv2.contourArea(cnt)
                if area < min_area or area > max_area:
                    continue
                x, y, w, h = cv2.boundingRect(cnt)
                if w < 24 or h < 18:
                    continue
                if any(box_iou((x, y, w, h), b["bbox"]) > 0.35 for b in boxes):
                    continue

                peri = cv2.arcLength(cnt, True)
                circularity = (4.0 * math.pi * area / (peri * peri)) if peri > 0 else 0.0
                aspect = w / float(max(1, h))
                prim = "ellipse" if (circularity > 0.72 and 0.7 <= aspect <= 1.4) else "roundRect"

                roi = frame[y : y + h, x : x + w]
                med_bgr = np.median(roi.reshape(-1, 3), axis=0).astype(int)
                dom_rgb = (int(med_bgr[2]), int(med_bgr[1]), int(med_bgr[0]))
                bdr_rgb = (
                    max(0, dom_rgb[0] - 35),
                    max(0, dom_rgb[1] - 35),
                    max(0, dom_rgb[2] - 35),
                )
                boxes.append({
                    "bbox": (x, y, w, h),
                    "area": area,
                    "primitive": prim,
                    "dominant_rgb": dom_rgb,
                    "border_rgb": bdr_rgb,
                })
            if len(boxes) >= 6:
                break

        # Fallback salient-region extraction for continuous full-bleed / textured video scenes
        if len(boxes) < 5:
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            p_hi = float(np.percentile(gray, 80))
            _, sal_mask = cv2.threshold(gray, max(40, int(p_hi)), 255, cv2.THRESH_BINARY)
            cnts, _ = cv2.findContours(sal_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            for cnt in sorted(cnts, key=cv2.contourArea, reverse=True):
                area = cv2.contourArea(cnt)
                if area < min_area or area > max_area:
                    continue
                x, y, w, h = cv2.boundingRect(cnt)
                if w < 28 or h < 20:
                    continue
                if any(box_iou((x, y, w, h), b["bbox"]) > 0.30 for b in boxes):
                    continue
                roi = frame[y : y + h, x : x + w]
                med_bgr = np.median(roi.reshape(-1, 3), axis=0).astype(int)
                dom_rgb = (int(med_bgr[2]), int(med_bgr[1]), int(med_bgr[0]))
                bdr_rgb = (max(0, dom_rgb[0] - 35), max(0, dom_rgb[1] - 35), max(0, dom_rgb[2] - 35))
                boxes.append({
                    "bbox": (x, y, w, h),
                    "area": area,
                    "primitive": "ellipse" if 0.8 <= (w / float(max(1, h))) <= 1.25 else "roundRect",
                    "dominant_rgb": dom_rgb,
                    "border_rgb": bdr_rgb,
                })
                if len(boxes) >= 7:
                    break

        return boxes
