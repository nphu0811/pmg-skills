"""
Acceptance, Regression, and Integration Tests for PMG Skills v1.6.0:
- Intelligent Visual Asset Engine (pmg-visual-assets)
- Image Decomposition & Vector Reconstruction (pmg-image-decomposer)
- Acceptance Demo A: Agribank Logo Motion Deck (Scenes 1-7)
- Acceptance Demo B: Image Shutter Reveal Deck
- PowerPoint Clean Open Verification (verify_clean_powerpoint_open)
"""
import os
import json
import numpy as np
import pytest
from PIL import Image

from pmg.version import VERSION
from pmg.visual_assets import (
    VisualAssetEngine,
    AssetRequirement,
    AssetRecord,
    WikimediaProvider,
    NASAProvider,
    score_asset,
)
from pmg.layers import LayerTree, VisualObject
from pmg.image_decomposer import ImageDecomposer
from pmg.asset_builder import build_logo_assembly_deck, build_image_shutter_deck
from pmg.validator import StaticValidator, verify_clean_powerpoint_open
from pmg.video_reconstructor import VideoMotionReconstructor


def test_v1_6_version():
    assert VERSION in ("1.6.0", "1.7.0", "1.8.0")


def test_wikimedia_search_and_provenance():
    prov = WikimediaProvider()
    req = AssetRequirement(query="Vietnam nature lotus flower", kind="image", min_width=640, min_height=480)
    results = prov.search(req, limit=2)
    assert len(results) > 0
    first = results[0]
    assert first.provenance == "retrieved"
    assert first.download_url.startswith("http")
    assert first.license_name != ""


def test_nasa_provider_search():
    prov = NASAProvider()
    req = AssetRequirement(query="Apollo Saturn rocket", kind="image", min_width=640, min_height=480)
    results = prov.search(req, limit=2)
    assert len(results) > 0
    first = results[0]
    assert first.provenance == "retrieved"
    assert "NASA" in first.attribution


def test_generator_fallback_distinction(tmp_path):
    class DummyGenerator:
        def generate(self, requirement, output_dir):
            out_img = os.path.join(output_dir, "synth.png")
            os.makedirs(output_dir, exist_ok=True)
            Image.new("RGBA", (800, 600), (255, 0, 0, 255)).save(out_img)
            return AssetRecord(
                asset_id="dummy_gen",
                title="Synthetic Red",
                kind="image",
                source_url="local://generator",
                download_url="file://" + out_img,
                provider="dummy_ai",
                provenance="generated",
                rights_status="compatible",
                local_path=out_img,
                width=800,
                height=600,
            )

    engine = VisualAssetEngine(providers=[], generator=DummyGenerator())
    req = AssetRequirement(query="abstract concept", kind="image")
    res = engine.acquire(req, str(tmp_path))
    assert res["status"] == "READY"
    assert res["provenance"] == "generated"
    assert res["asset"].provenance == "generated"


def test_semantic_decomposition_agribank_emblem(tmp_path):
    source_emblem = "internal_qa/v1_6_assets/agribank_emblem.png"
    if not os.path.exists(source_emblem):
        pytest.skip("agribank_emblem.png not found")

    decomposer = ImageDecomposer()
    out_dir = str(tmp_path / "emblem_test")
    tree = decomposer.decompose(source_emblem, out_dir, mode="semantic", granular=True)
    assert tree.reassembly["status"] == "PASS"
    assert tree.reassembly["pixel_similarity"] >= 0.98

    # Verify semantic roles
    roles = {obj.semantic_role for obj in tree.objects}
    assert "green_background" in roles
    assert "grain" in roles

    # Verify LayerTree JSON roundtrip
    saved_path = tree.save(str(tmp_path / "test_tree.json"))
    loaded = LayerTree.load(saved_path)
    assert len(loaded.objects) == len(tree.objects)


def test_geometric_slicing_vertical_and_radial(tmp_path):
    test_img = tmp_path / "test_pattern.png"
    arr = np.zeros((400, 600, 4), dtype=np.uint8)
    arr[:, :300] = [100, 150, 200, 255]
    arr[:, 300:] = [200, 100, 50, 255]
    Image.fromarray(arr).save(test_img)

    decomposer = ImageDecomposer()
    # Vertical strips
    tree_v = decomposer.decompose(str(test_img), str(tmp_path / "vertical"), mode="geometric", pattern="vertical", count=6)
    assert len(tree_v.objects) == 7  # 1 root group + 6 strips
    assert tree_v.reassembly["status"] == "PASS"

    # Radial sectors
    tree_r = decomposer.decompose(str(test_img), str(tmp_path / "radial"), mode="geometric", pattern="radial", count=4)
    assert len(tree_r.objects) == 5  # 1 root group + 4 sectors
    assert tree_r.reassembly["status"] == "PASS"


def test_demo_a_agribank_logo_assembly_presentation(tmp_path):
    layer_tree_path = "internal_qa/v1_6_assets/emblem_layers_v3/layer_tree.json"
    if not os.path.exists(layer_tree_path):
        pytest.skip("emblem_layers_v3/layer_tree.json not found")

    tree = LayerTree.load(layer_tree_path)
    output_pptx = str(tmp_path / "agribank_assembly_test.pptx")
    deck_path = build_logo_assembly_deck(output_pptx, tree, title="AGRIBANK DEMO")

    assert os.path.exists(deck_path)
    val = StaticValidator().validate(deck_path)
    assert val["is_valid"] is True
    assert len(val["defects"]) == 0
    assert val["total_slides"] == 7
    assert val["total_native_animations"] > 50

    # Strict clean open check
    clean = verify_clean_powerpoint_open(deck_path)
    assert clean["is_clean"] is True, f"Clean open failed: {clean.get('error')}"


def test_demo_b_image_shutter_reveal_presentation(tmp_path):
    layer_tree_path = "internal_qa/v1_6_assets/shutter_layers/layer_tree.json"
    if not os.path.exists(layer_tree_path):
        pytest.skip("shutter_layers/layer_tree.json not found")

    tree = LayerTree.load(layer_tree_path)
    output_pptx = str(tmp_path / "shutter_reveal_test.pptx")
    deck_path = build_image_shutter_deck(output_pptx, tree, title="SHUTTER DEMO")

    assert os.path.exists(deck_path)
    val = StaticValidator().validate(deck_path)
    assert val["is_valid"] is True
    assert len(val["defects"]) == 0
    assert val["total_slides"] == 3

    # Strict clean open check
    clean = verify_clean_powerpoint_open(deck_path)
    assert clean["is_clean"] is True, f"Clean open failed: {clean.get('error')}"


def test_video_reconstructor_shares_segmentation():
    reconstructor = VideoMotionReconstructor()
    assert hasattr(reconstructor, "segmenter")
    # Verify segmenter produces segments matching shared Segment protocol
    dummy_frame = np.zeros((240, 320, 3), dtype=np.uint8)
    boxes = reconstructor._segment_foreground_boxes(dummy_frame, (0, 0, 0), 320, 240)
    assert isinstance(boxes, list)
