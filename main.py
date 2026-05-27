"""Main entry point: process all images in images/ and generate PPTX output.

Applies Gaussian blur + sinusoidal displacement mapping (PPT artistic glass),
then center-crops to 16:9 for the slide background.
"""

import os
import sys
import tempfile

import cv2
import numpy as np
from PIL import Image, ImageFilter

from builder import build_pptx

# Calibrated: PPT blur radius 30 ≈ Pillow GaussianBlur sigma=15.0
BLUR_SIGMA = 15.0

# PPT artisticGlass scale (1-100). Controls displacement amplitude = scale/5.
# Higher = stronger wavy refractive distortion.
GLASS_SCALE = 34

# PPT artisticGlass transparency (0-100). 0 = fully distorted, 100 = original.
GLASS_TRANSPARENCY = 0

# Frequency ratio: image_dim / frequency = 80 (→ 25 cycles for 2000px image).
# Keeps the wave pattern visually consistent across different image sizes.
FREQ_RATIO = 80

# 16:9 aspect ratio
TARGET_RATIO = 16.0 / 9.0

# Paths
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGES_DIR = os.path.join(PROJECT_DIR, "images")
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")


def glass_displace(img_bgr, scale=34, transparency=0):
    """Apply PPT artistic glass effect via sinusoidal displacement mapping.

    The glass effect uses deterministic pixel remapping (NOT blur or
    pixelation): each pixel is shifted by a sinusoidal wave pattern,
    creating a refractive "looking through textured glass" distortion.

    Args:
        img_bgr: BGR image as numpy array (from cv2)
        scale: PPT scale parameter (1-100), controls displacement amplitude
        transparency: PPT transparency (0-100), 0=full distortion, 100=original

    Returns:
        Distorted BGR image.
    """
    rows, cols = img_bgr.shape[:2]

    # Create coordinate grid
    map_x, map_y = np.meshgrid(np.arange(cols), np.arange(rows))
    map_x = map_x.astype(np.float32)
    map_y = map_y.astype(np.float32)

    # Frequency scales with image size for consistent visual effect
    frequency = max(rows, cols) / FREQ_RATIO

    # Displacement amplitude: higher scale = stronger distortion
    amplitude = scale / 5.0

    # Sinusoidal displacement — horizontal shift varies with y,
    # vertical shift varies with x (cross-pattern creates glass-like refraction)
    x_offset = amplitude * np.sin(2 * np.pi * map_y / frequency)
    y_offset = amplitude * np.cos(2 * np.pi * map_x / frequency)

    new_x = map_x + x_offset
    new_y = map_y + y_offset

    distorted = cv2.remap(img_bgr, new_x, new_y, cv2.INTER_LINEAR)

    # Alpha blend with original based on transparency
    alpha = (100 - transparency) / 100.0
    result = cv2.addWeighted(distorted, alpha, img_bgr, 1 - alpha, 0)

    return result


def apply_glass_blur(image_path, output_path):
    """Apply PPT-equivalent blur + glass displacement to an image.

    Steps (matching the manual PPT workflow):
    1. GaussianBlur(sigma=15.0) — PPT artistic blur radius 30
    2. Sinusoidal displacement mapping — PPT artisticGlass
       (scale=34, transparency=0%)
    """
    # Load with PIL, blur, convert to OpenCV BGR
    img = Image.open(image_path).convert("RGB")
    blurred = img.filter(ImageFilter.GaussianBlur(radius=BLUR_SIGMA))

    # Convert PIL → numpy BGR for OpenCV
    blurred_arr = np.array(blurred)
    blurred_bgr = cv2.cvtColor(blurred_arr, cv2.COLOR_RGB2BGR)

    # Apply displacement-based glass effect
    glass_bgr = glass_displace(blurred_bgr, scale=GLASS_SCALE,
                                transparency=GLASS_TRANSPARENCY)

    # Convert back to RGB and save as PNG
    glass_rgb = cv2.cvtColor(glass_bgr, cv2.COLOR_BGR2RGB)
    Image.fromarray(glass_rgb).save(output_path, "PNG")

    return img.size  # (width, height) of original


def crop_to_16_9(img):
    """Center-crop an image to 16:9 aspect ratio."""
    w, h = img.size
    target_h = w / TARGET_RATIO

    if h > target_h:
        excess = h - target_h
        top = int(excess / 2)
        return img.crop((0, top, w, h - int(excess - top)))
    elif w > h * TARGET_RATIO:
        target_w = h * TARGET_RATIO
        excess = w - target_w
        left = int(excess / 2)
        return img.crop((left, 0, w - int(excess - left), h))
    else:
        return img


def process_images():
    """Process all images: blur + glass displacement + 16:9 crop.

    Returns list of dicts with orig_path, glass_path, img_width, img_height.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    temp_dir = tempfile.mkdtemp(prefix="ppt_glass_")

    exts = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff"}
    image_files = sorted(
        f for f in os.listdir(IMAGES_DIR)
        if os.path.splitext(f)[1].lower() in exts
    )

    if not image_files:
        print("No images found in images/ directory.")
        sys.exit(1)

    print(f"Found {len(image_files)} image(s) in images/")

    specs = []

    for i, filename in enumerate(image_files):
        orig_path = os.path.join(IMAGES_DIR, filename)
        glass_path = os.path.join(temp_dir, f"glass_{i:03d}.png")

        print(f"  Processing: {filename}")

        # Apply blur + glass displacement effect
        img_size = apply_glass_blur(orig_path, glass_path)

        # Crop result to 16:9
        img = Image.open(glass_path)
        cropped = crop_to_16_9(img)
        cropped.save(glass_path, "PNG")

        specs.append({
            "orig_path": orig_path,
            "glass_path": glass_path,
            "img_width": img_size[0],
            "img_height": img_size[1],
        })

    return specs


def main():
    specs = process_images()

    output_path = os.path.join(OUTPUT_DIR, "output.pptx")
    print(f"\nBuilding PPTX with {len(specs)} slide(s)...")
    build_pptx(specs, output_path)
    print(f"Done! Output saved to: {output_path}")


if __name__ == "__main__":
    main()
