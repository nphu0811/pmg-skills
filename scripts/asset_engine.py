"""
asset_engine.py (v3.0)
======================
Generates and manages multi-plane 3D isometric cutaways, studio grids,
and transparent energy flow textures for 3D Motion Graphics presentations.
"""

import os
from PIL import Image, ImageDraw, ImageFilter

def create_studio_grid_bg(output_path: str, width: int = 1920, height: int = 1080):
    """
    Creates a cinematic dark carbon studio background with a glowing perspective grid
    and vignette for 3D industrial explainer decks.
    """
    # 1. Dark gradient base (#0B111E to #152238)
    img = Image.new("RGBA", (width, height), (11, 17, 30, 255))
    draw = ImageDraw.Draw(img)
    
    # 2. Draw subtle isometric / perspective grid in lower half
    horizon_y = int(height * 0.45)
    center_x = int(width * 0.5)
    
    # Subtle blue-cyan grid lines (#1B2A47)
    grid_color = (27, 42, 71, 140)
    
    # Horizontal grid lines with increasing spacing
    y = horizon_y
    step = 8
    while y < height:
        draw.line([(0, y), (width, y)], fill=grid_color, width=1)
        step = int(step * 1.25)
        y += step
        
    # Perspective radiating lines from vanishing point
    for x in range(-width // 2, int(width * 1.5), 120):
        draw.line([(center_x, horizon_y), (x, height)], fill=grid_color, width=1)
        
    # 3. Soft atmospheric horizon glow
    glow = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    glow_draw = ImageDraw.Draw(glow)
    glow_draw.ellipse(
        [center_x - 600, horizon_y - 250, center_x + 600, horizon_y + 250],
        fill=(0, 245, 212, 28)
    )
    glow = glow.filter(ImageFilter.GaussianBlur(80))
    img = Image.alpha_composite(img, glow)
    
    # 4. Cinematic vignette
    vignette = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    v_draw = ImageDraw.Draw(vignette)
    v_draw.rectangle([0, 0, width, height], fill=(0, 0, 0, 80))
    # Punch out transparent center
    v_draw.ellipse([width*0.1, height*0.1, width*0.9, height*0.9], fill=(0, 0, 0, 0))
    vignette = vignette.filter(ImageFilter.GaussianBlur(100))
    img = Image.alpha_composite(img, vignette)
    
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    img.convert("RGB").save(output_path, "JPEG", quality=95)
    print(f"[ASSET_ENGINE] Created studio grid: {output_path}")
    return output_path

if __name__ == "__main__":
    out = os.path.join(os.path.dirname(__file__), "..", "assets", "studio_grid_bg.jpg")
    create_studio_grid_bg(out)
