# Lays each editor capture over its window screenshot and writes the website's .webp files.
# Run: python3 scripts/site-shots/compose.py <layers.json> <out-dir> [--new]
# --new also writes new-writer/new-sheets/new-pdf.webp (1200x750) for the changelog's What's new cards.
import sys, json, base64, io, os
from PIL import Image

layers, out = sys.argv[1], sys.argv[2]
data = json.load(open(layers))
k = data['scale']
os.makedirs(out, exist_ok=True)

def image(b64):
    return Image.open(io.BytesIO(base64.b64decode(b64))).convert('RGB')

for name, shot in data['results'].items():
    base = image(shot['base'])
    for layer in shot['layers']:
        b = layer['b']
        top = image(layer['png']).resize((round(b['width'] * k), round(b['height'] * k)), Image.LANCZOS)
        base.paste(top, (round(b['x'] * k), round(b['y'] * k)))
    base = base.resize((2000, 1250), Image.LANCZOS)
    base.save(os.path.join(out, f'{name}.webp'), quality=88, method=6)
    base.resize((1000, 625), Image.LANCZOS).save(os.path.join(out, f'{name.replace("-dark", "")}{"-dark" if name.endswith("-dark") else ""}-sm.webp'), quality=86, method=6)
    if '--new' in sys.argv and not name.endswith('-dark'):
        base.resize((1200, 750), Image.LANCZOS).save(os.path.join(out, f'new-{name}.webp'), quality=88, method=6)
    print('wrote', name)
