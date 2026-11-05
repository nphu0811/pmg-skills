---
name: pmg-visual-assets
description: "Intelligent visual asset research, retrieval, licensing, and media preparation. Discovers photography, illustrations, icons, and videos from authentic sources (Wikimedia, NASA, verified repositories) with explicit provenance, and feeds structured assets to PMG."
---

# PMG Visual Assets Engine

`pmg-visual-assets` provides source-first media research, licensing assessment, and asset preparation for PowerPoint Motion Graphics.
It is integrated with the canonical PMG engine via `pmg.visual_assets.VisualAssetEngine`.

## Key Principles & Boundaries

1. **Search vs. Generation Distinction**:
   - Searching for existing authentic assets and generating synthetic assets are distinct operations.
   - The engine prefers authentic permitted sources (Wikimedia Commons, NASA media archive, open vector repositories).
   - Generative fallback is activated only when no suitable asset exists and a supported generator adapter is configured.
   - Never fabricate search results, URLs, licenses, or generative output.

2. **Licensing & Rights Management**:
   - Every asset records provenance (`retrieved` vs `generated`), license name, license URL, author attribution, and restrictions.
   - Trademarks and brand logos require explicit rights review.

3. **Evaluation Criteria**:
   - Semantic relevance to the presentation topic.
   - Technical quality: resolution, aspect ratio, contrast, transparency.
   - Suitability for downstream decomposition and native PowerPoint animation.

4. **Engine Integration**:
   - Prepares media files and structured `AssetRecord` objects.
   - Does not compile PowerPoint animations directly; hands off assets to `pmg-image-decomposer`, `pmg-core`, and `pmg-morph`.

## Python API Usage

```python
from pmg.visual_assets import VisualAssetEngine, AssetRequirement, WikimediaProvider, NASAProvider

# 1. Initialize engine with configured providers
engine = VisualAssetEngine(providers=[WikimediaProvider(), NASAProvider()])

# 2. Define asset requirements
req = AssetRequirement(
    query="Space exploration Saturn rocket",
    kind="image",
    min_width=1280,
    min_height=720,
    transparency=False,
)

# 3. Research and rank candidates
search_result = engine.research(req, limit=5)

# 4. Acquire permitted asset with download and quality validation
acquired = engine.acquire(req, output_dir="internal_qa/assets")
if acquired["status"] == "READY":
    asset_record = acquired["asset"]
    print(f"Acquired: {asset_record.title} at {asset_record.local_path}")
```
