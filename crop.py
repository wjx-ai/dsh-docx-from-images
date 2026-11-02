#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Crop a picture out of a screenshot / photo for embedding in a .docx.

Usage:
    python crop.py spec.json

spec:
{
  "src":    "page1.png",          # source screenshot
  "out":    "fig1.png",           # output png
  "box":    [x0, y0, x1, y1],     # crop rectangle in source pixels (optional)
  "pad":    8,                    # extra pixels around the box

  # --- cleaning up a photographed page -------------------------------------
  "white_point": 195,             # paper level -> pure white (or true = auto)
  "black_point": 60,              # ink level -> black (auto when white_point=true)
  "grayscale":   true,            # write an 8-bit grey PNG (smaller, crisper)
  "trim":        true,            # trim the margin left around the ink
  "trim_threshold": 205,          # pixels darker than this count as ink

  # --- stitching several panels into one picture ---------------------------
  "compose":   true,              # stitch `panels` into a single image
  "panels":    [{"box":[..], "pad":4, "src":"page2.png"}, ...],
  "direction": "horizontal",      # or "vertical"
  "align":     "bottom",          # bottom | top | center  (cross-axis alignment)
  "gap":       48,                # pixels between panels

  "scale":   2.0,                 # upsample for print quality (optional)
  "grid":    false,               # instead of cropping, draw a labelled grid
  "step":    100,                 # grid step in pixels
  "info":    false                # only report the source size
}

Two things trip up screenshots of real paper, and both are handled here: the
page is photographed grey (not white), so a 245-threshold "trim" removes nothing
and the figure lands in Word as a grey box -- set `white_point`; and one question
often needs several panels (three rulers, two product pictures side by side)
while docx_typeset puts one picture per block -- use `compose`.
"""
from __future__ import annotations

import json
import os
import sys

from PIL import Image, ImageDraw

# A 245 threshold is right for a clean screenshot (pure-white margin) but useless
# on a photograph, whose "white" paper sits near 200.
DEFAULT_TRIM_THRESHOLD = 245
PHOTO_TRIM_THRESHOLD = 205


def _percentile(hist, total, frac):
    """Value at `frac` (0..1) of the luminance distribution, from a histogram."""
    target = total * frac
    acc = 0
    for value, count in enumerate(hist):
        acc += count
        if acc >= target:
            return value
    return len(hist) - 1


def auto_levels(grey):
    """Pick (black, white) so the paper clips to white and the ink deepens.

    The paper of a photographed page sits around the 85th percentile; pulling the
    white point slightly below it (minus a margin) makes the sheet pure white
    without eating into anti-aliased strokes.
    """
    hist = grey.histogram()
    total = sum(hist)
    if total == 0:
        return 0, 255
    lo = _percentile(hist, total, 0.02)
    paper = _percentile(hist, total, 0.85)
    hi = max(lo + 1, paper - 12)
    return lo, min(254, hi)


def apply_levels(grey, lo, hi):
    """Linear stretch: lo -> 0, hi -> 255, everything outside clipped."""
    if hi <= lo:
        return grey
    span = float(hi - lo)
    lut = [min(255, max(0, int((v - lo) / span * 255))) for v in range(256)]
    return grey.point(lut)


def trim_to_ink(img, threshold, trim_pad):
    """Crop away the margin that holds no ink (C-speed, not a Python pixel walk)."""
    grey = img if img.mode == "L" else img.convert("L")
    mask = grey.point(lambda v: 255 if v < threshold else 0)
    bbox = mask.getbbox()
    if bbox is None:
        return img, None
    left, top, right, bottom = bbox
    left = max(0, left - trim_pad)
    top = max(0, top - trim_pad)
    right = min(img.width, right + trim_pad)
    bottom = min(img.height, bottom + trim_pad)
    return img.crop((left, top, right, bottom)), bbox


def ink_box(img, threshold):
    """Bounding box of ink in a finished picture, for diagnostics."""
    grey = img if img.mode == "L" else img.convert("L")
    return grey.point(lambda v: 255 if v < threshold else 0).getbbox()


def prepare_panel(spec, panel, base):
    """Crop + level + trim one panel; shared by single-image and compose modes."""
    src = panel.get("src") or spec.get("src")
    if not os.path.isabs(src):
        src = os.path.join(base, src)
    img = Image.open(src)

    box = panel.get("box") or spec.get("box")
    if not box:
        raise SystemExit("spec needs `box` (or one per panel)")
    pad = int(panel.get("pad", spec.get("pad", 0)))
    x0, y0, x1, y1 = [int(v) for v in box]
    x0, y0 = max(0, x0 - pad), max(0, y0 - pad)
    x1, y1 = min(img.width, x1 + pad), min(img.height, y1 + pad)
    if x1 <= x0 or y1 <= y0:
        raise SystemExit("empty crop box: %r" % (box,))
    fig = img.crop((x0, y0, x1, y1))

    grayscale = bool(panel.get("grayscale", spec.get("grayscale", False)))
    white_point = panel.get("white_point", spec.get("white_point"))

    if white_point is not None:
        grey = fig.convert("L")
        if white_point is True or white_point == "auto":
            lo, hi = auto_levels(grey)
        else:
            hi = int(white_point)
            lo = int(panel.get("black_point", spec.get("black_point", 0)))
        fig = apply_levels(grey, lo, hi)
        grayscale = True

    trim_threshold = int(
        panel.get(
            "trim_threshold",
            spec.get("trim_threshold")
            or (PHOTO_TRIM_THRESHOLD if white_point is not None else DEFAULT_TRIM_THRESHOLD),
        )
    )
    panel_trim = panel.get("trim", spec.get("trim"))
    if panel_trim:
        fig, _ = trim_to_ink(fig, trim_threshold, int(spec.get("trim_pad", 4)))

    if not grayscale and fig.mode != "L":
        fig = fig.convert("RGB")
    return fig, {"box": [x0, y0, x1, y1], "src": os.path.abspath(src)}


def compose_panels(spec, base):
    """Stitch several cropped panels into one picture (row or column)."""
    panels = spec.get("panels") or []
    if not panels:
        raise SystemExit("compose=true needs a non-empty `panels` list")

    figs, meta = [], []
    for panel in panels:
        fig, info = prepare_panel(spec, panel, base)
        figs.append(fig)
        meta.append(info)

    direction = (spec.get("direction") or "horizontal").lower()
    align = (spec.get("align") or "bottom").lower()
    gap = int(spec.get("gap", 32))

    if direction == "vertical":
        width = max(f.width for f in figs)
        height = sum(f.height for f in figs) + gap * (len(figs) - 1)
    else:
        width = sum(f.width for f in figs) + gap * (len(figs) - 1)
        height = max(f.height for f in figs)

    mode = "L" if all(f.mode == "L" for f in figs) else "RGB"
    fill = 255 if mode == "L" else (255, 255, 255)
    canvas = Image.new(mode, (max(1, width), max(1, height)), fill)

    cursor = 0
    for fig in figs:
        if fig.mode != mode:
            fig = fig.convert(mode)
        if direction == "vertical":
            if align == "center":
                x = (width - fig.width) // 2
            elif align == "right":
                x = width - fig.width
            else:
                x = 0
            canvas.paste(fig, (x, cursor))
            cursor += fig.height + gap
        else:
            if align == "center":
                y = (height - fig.height) // 2
            elif align == "top":
                y = 0
            else:
                y = height - fig.height
            canvas.paste(fig, (cursor, y))
            cursor += fig.width + gap
    return canvas, meta


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit("usage: crop.py spec.json")
    spec_path = os.path.abspath(sys.argv[1])
    base = os.path.dirname(spec_path)
    with open(spec_path, "r", encoding="utf-8") as fh:
        spec = json.load(fh)

    src = spec.get("src")
    if not src:
        raise SystemExit("spec needs `src`")
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

    if spec.get("compose"):
        fig, meta = compose_panels(spec, base)
    else:
        if not spec.get("box"):
            raise SystemExit("spec needs `box` unless grid/info/compose is set")
        fig, meta = prepare_panel(spec, {}, base)
        meta = [meta]

    scale = float(spec.get("scale", 1) or 1)
    if scale != 1:
        fig = fig.resize((max(1, int(fig.width * scale)), max(1, int(fig.height * scale))),
                         Image.LANCZOS)

    out = spec.get("out") or "figure.png"
    if not os.path.isabs(out):
        out = os.path.join(base, out)
    fig.save(out, optimize=True)

    threshold = int(spec.get("trim_threshold") or DEFAULT_TRIM_THRESHOLD)
    result.update({
        "path": os.path.abspath(out),
        "crop_width": fig.width,
        "crop_height": fig.height,
        "box": None if spec.get("compose") else meta[0].get("box"),
        "panels": len(meta) if spec.get("compose") else None,
        "ink_box": ink_box(fig, threshold),
    })
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()