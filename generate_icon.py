"""Generate nebaxus.ico from nebaxus.svg.
Usage: python generate_icon.py
Requires: pip install cairosvg pillow
"""

import io
import os

try:
    import cairosvg
    from PIL import Image
except ImportError:
    print('Missing dependencies: pip install cairosvg pillow')
    raise

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
SVG_PATH = os.path.join(PROJECT_DIR, 'nebaxus.svg')
ICO_PATH = os.path.join(PROJECT_DIR, 'nebaxus.ico')

SIZES = [16, 24, 32, 48, 64, 128, 256]

def main():
    if not os.path.exists(SVG_PATH):
        print(f'SVG not found: {SVG_PATH}')
        return
    images = []
    for s in SIZES:
        png_data = cairosvg.svg2png(url=SVG_PATH, output_width=s, output_height=s)
        img = Image.open(io.BytesIO(png_data))
        images.append(img)
    images[0].save(
        ICO_PATH,
        format='ICO',
        sizes=[(s, s) for s in SIZES],
        append_images=images[1:],
    )
    print(f'Icon generated: {ICO_PATH} ({len(SIZES)} sizes)')

if __name__ == '__main__':
    main()
