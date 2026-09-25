"""
Generates modern application logo image assets/logo.png.
"""

import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from PIL import Image, ImageDraw
import config


def generate_logo():
    os.makedirs(config.ASSETS_DIR, exist_ok=True)
    logo_path = os.path.join(config.ASSETS_DIR, "logo.png")

    width, height = 450, 110
    img = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    # Card background
    card = [5, 5, width - 5, height - 5]
    draw.rounded_rectangle(card, radius=14, fill=(15, 23, 42, 255), outline=(30, 58, 138, 255), width=2)

    # Icon Badge
    badge = [20, 20, 95, 90]
    draw.rounded_rectangle(badge, radius=10, fill=(30, 58, 138, 255))
    draw.text((36, 42), "AI", fill=(255, 255, 255, 255))

    # Text Titles
    draw.text((115, 26), "Universal AI Excel Converter", fill=(248, 250, 252, 255))
    draw.text((115, 58), "Production-Grade Format Transformation Engine", fill=(148, 163, 184, 255))

    img.save(logo_path, "PNG")
    print(f"Logo generated successfully at: {logo_path}")


if __name__ == "__main__":
    generate_logo()
