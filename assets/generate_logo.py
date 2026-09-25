"""
Generates the OTIS Excel Report Formatter application logo.
"""

import os
import sys

# Ensure parent directory is in sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from PIL import Image, ImageDraw
import config


def generate_logo():
    config.ASSETS_DIR
    os.makedirs(config.ASSETS_DIR, exist_ok=True)
    logo_path = os.path.join(config.ASSETS_DIR, "logo.png")

    width, height = 400, 100
    img = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    # Draw rounded dark blue gradient background card
    card_bounds = [5, 5, width - 5, height - 5]
    draw.rounded_rectangle(card_bounds, radius=12, fill=(15, 23, 42, 255), outline=(51, 65, 85, 255), width=2)

    # Draw OTIS Badge
    badge_bounds = [20, 20, 90, 80]
    draw.rounded_rectangle(badge_bounds, radius=8, fill=(37, 99, 235, 255))
    draw.text((32, 38), "OTIS", fill=(255, 255, 255, 255))

    # Draw text titles
    draw.text((110, 25), "OTIS Excel Report Formatter", fill=(248, 250, 252, 255))
    draw.text((110, 55), "Automated Inspection Report Generator", fill=(148, 163, 184, 255))

    img.save(logo_path, "PNG")
    print(f"Logo generated successfully at: {logo_path}")


if __name__ == "__main__":
    generate_logo()
