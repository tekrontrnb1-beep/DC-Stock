# -*- coding: utf-8 -*-
"""Generate PWA PNG icons (gradient + bar-chart glyph) from scratch with Pillow.
Run once: python make_icons.py  → icon-192.png, icon-512.png, icon-512-maskable.png"""
import os
from PIL import Image, ImageDraw

BASE = os.path.dirname(os.path.abspath(__file__))
C1 = (37, 99, 235)    # #2563eb
C2 = (108, 92, 231)   # #6c5ce7

def lerp(a, b, t): return tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))

def gradient(S):
    g = Image.new('RGB', (S, S))
    px = g.load()
    for y in range(S):
        for x in range(S):
            px[x, y] = lerp(C1, C2, (x + y) / (2 * (S - 1)))
    return g

def make(S, rounded=True, scale=1.0):
    img = Image.new('RGBA', (S, S), (0, 0, 0, 0))
    g = gradient(S)
    if rounded:
        mask = Image.new('L', (S, S), 0)
        ImageDraw.Draw(mask).rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.219), fill=255)
        img.paste(g, (0, 0), mask)
    else:
        img.paste(g, (0, 0))
    d = ImageDraw.Draw(img)
    off = (1 - scale) * S / 2
    for (bx, by, bw, bh) in [(15, 32, 8, 17), (28, 23, 8, 26), (41, 15, 8, 34)]:
        x0 = off + (bx / 64) * S * scale; y0 = off + (by / 64) * S * scale
        x1 = off + ((bx + bw) / 64) * S * scale; y1 = off + ((by + bh) / 64) * S * scale
        d.rounded_rectangle([x0, y0, x1, y1], radius=(bw / 64) * S * scale * 0.32, fill=(255, 255, 255, 245))
    return img

make(192).save(os.path.join(BASE, 'icon-192.png'))
make(512).save(os.path.join(BASE, 'icon-512.png'))
make(512, rounded=False, scale=0.72).save(os.path.join(BASE, 'icon-512-maskable.png'))
print('icons written: icon-192.png, icon-512.png, icon-512-maskable.png')
