"""Canonical asset demonstrations composed with existing engine/animation APIs."""
import json
import mimetypes
from pathlib import Path
from typing import Optional

from pmg.animations import SlideAnimationTimeline
from pmg.layers import LayerTree
from pmg.native_objects import compile_layer_tree, fit_layer_bounds
from pmg.package_parts import write_package_parts
from scripts.motion_graph_kit import MotionDeck, Themes


def _embed_asset_expectations(path, tree, expectations):
    source_part = 'ppt/pmg_references/'+Path(tree.source_path).name
    mime = mimetypes.guess_type(source_part)[0] or 'application/octet-stream'
    parts = {source_part: Path(tree.source_path).read_bytes(),
             'ppt/pmg_asset_expectations.json': json.dumps({
                 'source_part': source_part, 'root_id': tree.root_id,
                 'reassembly': tree.reassembly, 'scenes': expectations,
                 'asset_provenance': tree.metadata.get('asset_record', {}),
                 'layer_editability': {obj.object_id: obj.editability for obj in tree.objects},
             }, indent=2).encode('utf-8')}
    write_package_parts(path, parts, {source_part: mime, 'ppt/pmg_asset_expectations.json': 'application/json'})


def build_logo_assembly_deck(output_path: str, tree: LayerTree, title: str = 'LOGO ASSEMBLY'):
    deck = MotionDeck(Themes.LIGHT_MODERN_STUDIO)
    components = [obj for obj in tree.objects if obj.kind != 'group']
    stages = [
        ('THE FOUNDATION', lambda obj: obj.semantic_role == 'green_background'),
        ('DIRECTION TAKES SHAPE', lambda obj: obj.semantic_role in ('green_background', 'red_diagonal')),
        ('LINES FIND THEIR PLACE', lambda obj: obj.semantic_role in ('green_background', 'red_diagonal', 'white_separator')),
        ('INDEPENDENT WHEAT COMPONENTS', lambda obj: obj.semantic_role != 'detail_residual'),
        ('THE COMPONENTS CONVERGE', lambda obj: True),
        ('THE COMPLETE EMBLEM', lambda obj: True),
        ('A VISUAL IDENTITY IN MOTION', lambda obj: True),
    ]
    expectations = []
    for number, (heading, selected) in enumerate(stages, 1):
        slide = deck.add_slide(transition='none' if number == 1 else ('circle' if number == 7 else 'morph'))
        deck.add_hero_title(slide, title if number == 1 else heading,
                            subtitle='Independent native vectors and transparent image layers')
        bounds = (4.1, 1.85, 5.1, 4.65) if number < 7 else (.9, 2.0, 4.6, 4.6)
        chosen = [obj.object_id for obj in components if selected(obj)]
        mapping = compile_layer_tree(slide, tree, bounds, include_ids=chosen)
        timeline = SlideAnimationTimeline(number)
        for index, obj in enumerate(components):
            shape = mapping.get(obj.object_id)
            if shape is None:
                continue
            if number == 1:
                if obj.semantic_role == 'green_background':
                    timeline.add_effect(shape, 'zoom', duration_ms=850, delay_ms=index*25)
            elif number == 2 and obj.semantic_role == 'red_diagonal':
                timeline.add_effect(shape, 'fly_in', direction='left', duration_ms=750, delay_ms=index*25)
            elif number == 3 and obj.semantic_role == 'white_separator':
                timeline.add_effect(shape, 'fly_in', direction='right', duration_ms=700)
            elif number in (4, 5) and obj.semantic_role in ('grain', 'stem'):
                delay = index*45
                timeline.add_effect(shape, 'appear', duration_ms=1, delay_ms=delay)
                sign = -1 if index % 2 else 1
                timeline.add_effect(shape, 'motion_path', duration_ms=1500, delay_ms=delay,
                                    motion_path=f'M {sign*.28} -.18 C {sign*.18} -.35 {sign*.08} .16 0 0 E')
                timeline.add_effect(shape, 'rotate', duration_ms=1500, delay_ms=delay, rotation_deg=360)
        timeline.compile_to_openxml(slide)
        if number >= 5:
            x, y, scale = fit_layer_bounds(tree, bounds)
            expectations.append({'slide_number': number, 'roi': [x, y, tree.source_size_px[0]*scale, tree.source_size_px[1]*scale],
                                 'shape_ids': [shape.shape_id for shape in mapping.values()], 'tolerance': .90})
        if number == 7:
            deck.add_glass_card(slide, 6.25, 2.45, 5.75, 2.7, name='!!Brand_Layout_Card')
            deck.add_badge(slide, 'REUSABLE COMPONENTS', 6.65, 2.8, width=4.6)
    deck.save(output_path)
    _embed_asset_expectations(output_path, tree, expectations)
    return str(Path(output_path).resolve())


def build_image_shutter_deck(output_path: str, tree: LayerTree, title='A LANDSCAPE IN MOTION'):
    deck = MotionDeck(Themes.LIGHT_ACADEMIC_CYAN)
    expectations = []
    bounds = (.9, 1.8, 11.5, 5.0)
    for number in range(1, 4):
        slide = deck.add_slide(transition='none' if number == 1 else ('morph' if number == 2 else 'circle'))
        deck.add_hero_title(slide, title if number == 1 else ('ONE IMAGE, INDEPENDENT LAYERS' if number == 2 else 'THE COMPLETE VIEW'))
        mapping = compile_layer_tree(slide, tree, bounds)
        timeline = SlideAnimationTimeline(number)
        if number == 1:
            for index, obj in enumerate(item for item in tree.objects if item.kind != 'group'):
                timeline.add_effect(mapping[obj.object_id], 'fly_in', direction='up' if index % 2 else 'down',
                                    duration_ms=1100, delay_ms=index*120)
        timeline.compile_to_openxml(slide)
        x, y, scale = fit_layer_bounds(tree, bounds)
        expectations.append({'slide_number': number, 'roi': [x, y, tree.source_size_px[0]*scale, tree.source_size_px[1]*scale],
                             'shape_ids': [shape.shape_id for shape in mapping.values()], 'tolerance': .94})
    deck.save(output_path)
    _embed_asset_expectations(output_path, tree, expectations)
    return str(Path(output_path).resolve())
