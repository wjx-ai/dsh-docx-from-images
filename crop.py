#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Crop a picture out of a screenshot for embedding in a .docx.

Usage:
    python crop.py spec.json

spec:
{
  "src":    "page1.png",          # source screenshot
  "out":    "fig1.png",           # output png
  "box":    [x0, y0, x1, y1],     # crop rectangle in source pixels (optional)
  "pad":    8,                     # extra pixels around the box
  "trim":   true,                  # trim uniform white margin after cropping
  "scale":  2.0,                   # upsample for print quality (optional)
  "grid":   false,                 # instead of cropping, draw a labelled grid
  "step":   100,                   # grid step in pixels
  "info":   false                  # only report the source size
}
"""
from __future__ import annotations

import json
import os
import sys

from PIL import Image, ImageDraw


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit("usage: crop.py spec.json")
    spec_path = os.path.abspath(sys.argv[1])
    base = os.path.dirname(spec_path)
    with open(spec_path, "r", encoding="utf-8") as fh:
        spec = json.load(fh)

    src = spec["src"]
    if not os.path.isabs(src):
        src = os.path.join(base, src)
    img = Image.open(src)
    result = {"ok": True, "src": os.path.abspath(src),
              "width": img.width, "height": img.height}

    if spec.get("info"):
        print(json.dumps(result, ensure_ascii=False))
        return

    if spec.get("grid"):
        step = int(spec.get("step", 100))
        out = spec.get("out") or "grid.png"
        if not os.path.isabs(out):
            out = os.path.join(base, out)
        canvas = img.convert("RGB").copy()
        draw = ImageDraw.Draw(canvas)
        for x in range(0, canvas.width, step):
            draw.line([(x, 0), (x, canvas.height)], fill=(255, 0, 0), width=1)
            draw.text((x + 3, 3), str(x), fill=(255, 0, 0))
        for y in range(0, canvas.height, step):
            draw.line([(0, y), (canvas.width, y)], fill=(0, 128, 255), width=1)
            draw.text((3, y + 2), str(y), fill=(0, 128, 255))
        canvas.save(out)
        result.update({"path": os.path.abspath(out), "grid": step})
        print(json.dumps(result, ensure_ascii=False))
        return

    box = spec.get("box")
    if not box:
        raise SystemExit("spec needs `box` unless grid/info is set")
    pad = int(spec.get("pad", 0))
    x0, y0, x1, y1 = [int(v) for v in box]
    x0, y0 = max(0, x0 - pad), max(0, y0 - pad)
    x1, y1 = min(img.width, x1 + pad), min(img.height, y1 + pad)
    if x1 <= x0 or y1 <= y0:
        raise SystemExit("empty crop box")
    fig = img.crop((x0, y0, x1, y1))

    if spec.get("trim"):
        rgb = fig.convert("RGB")
        w, h = rgb.size
        px = rgb.load()
        left, top, right, bottom = w, h, 0, 0
        for y in range(h):
            for x in range(w):
                r, g, b = px[x, y]
                if r < 245 or g < 245 or b < 245:
                    left, right = min(left, x), max(right, x)
                    top, bottom = min(top, y), max(bottom, y)
        if right >= left and bottom >= top:
            pad2 = int(spec.get("trim_pad", 4))
            left = max(0, left - pad2)
            top = max(0, top - pad2)
            right = min(w - 1, right + pad2)
            bottom = min(h - 1, bottom + pad2)
            fig = fig.crop((left, top, right + 1, bottom + 1))

    scale = float(spec.get("scale", 1) or 1)
    if scale != 1:
        fig = fig.resize((int(fig.width * scale), int(fig.height * scale)),
                         Image.LANCZOS)

    out = spec["out"]
    if not os.path.isabs(out):
        out = os.path.join(base, out)
    fig.save(out)
    result.update({"path": os.path.abspath(out), "crop_width": fig.width,
                   "crop_height": fig.height, "box": [x0, y0, x1, y1]})
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()