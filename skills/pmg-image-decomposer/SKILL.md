---
name: pmg-image-decomposer
description: "Decompose flat raster images, logos, and photographs into structured, reusable LayerTree hierarchies with native vector polygons, transparent raster layers, geometric slices, and automatic reassembly validation for PowerPoint animation."
---

# PMG Image Decomposer

`pmg-image-decomposer` converts flat images into structured, reusable collections of independent visual objects consumable by the canonical PMG engine.
Integrated via `pmg.image_decomposer.ImageDecomposer` and `pmg.layers.LayerTree`.

## Core Decomposition Modes

1. **Mode 1 — Semantic Decomposition**:
   - Isolates meaningful brand/visual elements (e.g., Agribank green background, red diagonal bands, white separators, wheat stem and grains).
   - Classifies editability: `native_vector` (OpenXML `<a:custGeom>` freeform polygons) vs. `raster_layer` (transparent PNG).
   - Generates hierarchical grouping (`wheat_group` -> `Grain_01..N`).

2. **Mode 2 — Geometric Slicing**:
   - Partitions images into geometric segments for shutter, mosaic, tile, and carousel reveals.
   - Patterns: `vertical`, `horizontal`, `grid`, `tiles`, `puzzle`, `radial` / `circular_sectors`, `triangles`, `custom` polygons.
   - Guarantees complete pixel coverage without dropped content.

3. **Mode 3 — Depth Decomposition**:
   - Partitions scenes into foreground, midground, and background spatial planes.
   - Supports 2.5D camera parallax and tracking shots.
   - Explicitly flags inferred or inpainted occluded areas.

## Reassembly Validation Gate

Every decomposition executes an automatic reassembly verification:
- Renders reconstructed layers and measures pixel similarity (`>= 0.94` tolerance) and contour alignment against the original asset.
- Reports raster-only vs. vector components.
- Layer trees with failed reassembly are rejected before PowerPoint compilation.

## Native PowerPoint Compilation

Layer trees compile into native PowerPoint objects via `pmg.native_objects.compile_layer_tree`:
- Groups become PowerPoint `!!` group shapes.
- Native vectors become editable `<a:custGeom>` shapes with custom paths.
- Transparent raster segments become positioned and cropped picture shapes.
- Objects retain stable `!!` identities across slides for Morph transitions.

## Python API Usage

```python
from pmg.image_decomposer import ImageDecomposer
from pmg.native_objects import compile_layer_tree
from pmg.layers import LayerTree

decomposer = ImageDecomposer()

# 1. Semantic decomposition
tree = decomposer.decompose(
    source_path="path/to/logo.png",
    output_dir="layers/logo",
    mode="semantic",
    granular=True,
)
assert tree.reassembly["status"] == "PASS"

# 2. Geometric slicing
shutter_tree = decomposer.decompose(
    source_path="path/to/photo.jpg",
    output_dir="layers/shutter",
    mode="geometric",
    pattern="vertical",
    count=8,
)

# 3. Compile to native PowerPoint slide
# slide is a python-pptx Slide object
mapping = compile_layer_tree(slide, tree, bounds=(1.0, 1.5, 11.3, 5.0))
```
