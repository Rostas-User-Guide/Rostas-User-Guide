#!/usr/bin/env python3
"""
Regenerate every cropped, highlighted guide screenshot from its original.

    python3 make_crops.py            # rebuild all crops
    python3 make_crops.py board-swap # rebuild only crops whose name contains this

Reads crops.json. Each entry names a source screenshot (kept untouched), a crop
box and up to two highlight boxes, all in the source image's pixel coordinates
as [left, top, right, bottom]. Writes crops/<name>.png and prints the <img>
width/height to use in index.html.

When a screenshot is re-shot: replace the original file, open it, check the
crop box and highlights in crops.json still frame the right things, rerun this.
"""
import json
import os
import sys

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
SS = 3  # supersampling factor for smooth highlight corners


def hex_rgba(h, a=255):
    h = h.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)) + (a,)


def draw_highlights(img, boxes, style):
    """Rounded box in the guide's red with a thin white halo so it reads on any UI."""
    w, h = img.size
    layer = Image.new('RGBA', (w * SS, h * SS), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    stroke, halo, pad, rad = (style[k] * SS for k in ('stroke', 'halo', 'pad', 'radius'))
    for l, t, r, b in boxes:
        box = [l * SS - pad, t * SS - pad, r * SS + pad, b * SS + pad]
        outer = [box[0] - halo, box[1] - halo, box[2] + halo, box[3] + halo]
        d.rounded_rectangle(outer, radius=rad + halo, outline=hex_rgba('#ffffff'),
                            width=stroke + 2 * halo)
        d.rounded_rectangle(box, radius=rad, outline=hex_rgba(style['colour']), width=stroke)
    layer = layer.resize((w, h), Image.LANCZOS)
    return Image.alpha_composite(img.convert('RGBA'), layer)


def main():
    cfg = json.load(open(os.path.join(HERE, 'crops.json')))
    style = cfg['style']
    only = sys.argv[1] if len(sys.argv) > 1 else ''
    os.makedirs(os.path.join(HERE, 'crops'), exist_ok=True)
    problems = 0
    for name, c in cfg['crops'].items():
        if only and only not in name:
            continue
        src = Image.open(os.path.join(HERE, c['src']))
        sw, sh = src.size
        l, t, r, b = c['box']
        if not (0 <= l < r <= sw and 0 <= t < b <= sh):
            print(f'!! {name}: crop box {c["box"]} is outside {c["src"]} ({sw}x{sh})')
            problems += 1
            continue
        hl = c.get('highlights', [])
        if len(hl) > 2:
            print(f'!! {name}: {len(hl)} highlights; the style allows at most two')
            problems += 1
        local = []
        for hb in hl:
            if not (l <= hb[0] < hb[2] <= r and t <= hb[1] < hb[3] <= b):
                print(f'!! {name}: highlight {hb} is outside the crop box')
                problems += 1
            local.append([hb[0] - l, hb[1] - t, hb[2] - l, hb[3] - t])
        out = draw_highlights(src.crop((l, t, r, b)), local, style)
        out.convert('RGB').save(os.path.join(HERE, 'crops', name + '.png'), optimize=True)
        # Show crops a little larger than the old squeezed full windows did,
        # but never wider than the column (CSS max-width handles that).
        dw = round((r - l) * style['display_scale'])
        dh = round((b - t) * style['display_scale'])
        print(f'{name}: {r - l}x{b - t}  ->  width="{dw}" height="{dh}"')
    sys.exit(1 if problems else 0)


if __name__ == '__main__':
    main()
