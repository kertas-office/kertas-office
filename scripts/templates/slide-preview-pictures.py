"""Turns the rendered slides (slide-previews.mjs) into the 480x360 gallery pictures in src/templates/:
the first slide large, the next three below it. Run: python3 scripts/templates/slide-preview-pictures.py <out-dir>"""
import glob, os, sys
from PIL import Image, ImageDraw

src = sys.argv[1]
dest = os.path.join(os.path.dirname(__file__), '..', '..', 'src', 'templates')
for d in sorted(os.listdir(src)):
    files = sorted(glob.glob(os.path.join(src, d, '*.png')))
    if not files:
        continue
    pic = Image.new('RGB', (480, 360), '#e9edf3')
    big_w = 448; big_h = round(big_w * 9 / 16)
    pic.paste(Image.open(files[0]).convert('RGB').resize((big_w, big_h), Image.LANCZOS), (16, 16))
    gap = 8; small_w = (big_w - 2 * gap) // 3; small_h = round(small_w * 9 / 16)
    y = 16 + big_h + 10
    for i, f in enumerate(files[1:4]):
        x = 16 + i * (small_w + gap)
        pic.paste(Image.open(f).convert('RGB').resize((small_w, small_h), Image.LANCZOS), (x, y))
    draw = ImageDraw.Draw(pic)
    draw.rectangle([16, 16, 16 + big_w - 1, 16 + big_h - 1], outline='#cbd5e1')
    for i in range(min(3, len(files) - 1)):
        x = 16 + i * (small_w + gap)
        draw.rectangle([x, y, x + small_w - 1, y + small_h - 1], outline='#cbd5e1')
    pic.save(os.path.join(dest, f'{d}.jpg'), quality=88)
    print(d)
