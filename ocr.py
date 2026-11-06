#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Read text (with coordinates) out of an image, for a model that cannot see it.

Usage:
    python ocr.py spec.json

spec:
{
  "src": "page1.jpg",          # image to read
  "lang": "zh-Hans-CN",        # Windows OCR language tag
  "region": [x0,y0,x1,y1],     # only read this part (optional)
  "scale": 2,                  # upscale before OCR (small print reads better)
  "white_point": true,         # whiten a photographed page first (see crop.py)
  "figures": true,             # also report where the illustrations are
  "min_gap": 45                # a text gap wider than this may hold a figure
  "out": "ocr.json"            # optional: write the JSON here as well
}

Output (stdout, JSON):
{
  "ok": true, "src": ..., "width": .., "height": .., "lines": 123,
  "text": "line1\\nline2...",
  "boxes": [{"y0":..,"y1":..,"x0":..,"x1":..,"text":".."}, ...],
  "figures": [{"y0":..,"y1":..,"x0":..,"x1":..,"ink":..}, ...]
}

Everything is in the SOURCE image's pixel space, whatever `scale` was used.

`figures` is the trick that replaces looking at the page: illustrations sit in the
empty horizontal bands between text lines, so the gaps between OCR line boxes
that still contain ink ARE the figures. Each one is verified with an ink mask, so
blank space between sections is not reported.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
PS_SCRIPT = os.path.join(HERE, "ocr_ps.ps1")

# A 245 threshold is right for a clean screenshot; a photograph needs whitening first.
DEFAULT_TRIM_THRESHOLD = 245
PHOTO_TRIM_THRESHOLD = 205


def apply_levels(grey, lo, hi):
    if hi <= lo:
        return grey
    span = float(hi - lo)
    lut = [min(255, max(0, int((v - lo) / span * 255))) for v in range(256)]
    return grey.point(lut)


def auto_levels(grey):
    hist = grey.histogram()
    total = sum(hist)
    if total == 0:
        return 0, 255

    def pct(frac):
        target, acc = total * frac, 0
        for value, count in enumerate(hist):
            acc += count
            if acc >= target:
                return value
        return 255

    lo = pct(0.02)
    paper = pct(0.85)
    return lo, max(lo + 1, min(254, paper - 12))


def run_powershell(src, out, lang):
    exe = shutil.which("powershell") or shutil.which("pwsh")
    if not exe:
        raise SystemExit("Windows OCR needs powershell/pwsh on PATH")
    cmd = [exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", PS_SCRIPT,
           "-src", src, "-out", out]
    if lang:
        cmd += ["-lang", lang]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise SystemExit("OCR failed: %s" % ((proc.stderr or proc.stdout).strip()[:400]))


def is_cjk(ch):
    o = ord(ch)
    return (0x4E00 <= o <= 0x9FFF or 0x3400 <= o <= 0x4DBF or 0xF900 <= o <= 0xFAFF)


CJK_PUNCT = set("（）【】《》，。：；、？！“”‘’—…·")


def normalize_text(text):
    """Windows OCR puts a space between every CJK glyph; glue them back.

    Only CJK-CJK and CJK-punctuation boundaries are collapsed, so "66 ÷ 3 + 3"
    keeps the spacing that makes it readable.
    """
    out = []
    for i, ch in enumerate(text):
        if ch == " " and out and i + 1 < len(text):
            prev, nxt = out[-1], text[i + 1]
            if (is_cjk(prev) and is_cjk(nxt)) or prev in CJK_PUNCT or nxt in CJK_PUNCT:
                continue
        out.append(ch)
    return "".join(out).strip()


def rule_columns(mask, frac=0.55):
    """Columns of ink running down most of the page: rules, frames, 密封线.

    A mask resized to width x 1 averages each column (BOX filter), which is much
    cheaper than summing pixels column by column.
    """
    col = mask.resize((mask.width, 1), Image.BOX)
    row = col.load()
    return set(x for x in range(mask.width) if row[x, 0] > int(255 * frac))


def ink_bbox(img, box, threshold, drop_cols=None):
    """Bounding box of ink inside `box`, or None. PIL only, no numpy."""
    x0, y0, x1, y1 = [int(v) for v in box]
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(img.width, x1), min(img.height, y1)
    if x1 <= x0 or y1 <= y0:
        return None
    band = img.crop((x0, y0, x1, y1)).convert("L")
    mask = band.point(lambda v: 255 if v < threshold else 0)
    for x in drop_cols or ():
        if x0 <= x < x1:
            mask.paste(0, (x - x0, 0, x - x0 + 1, mask.height))
    bb = mask.getbbox()
    if bb is None:
        return None
    return (x0 + bb[0], y0 + bb[1], x0 + bb[2], y0 + bb[3])


def merge(boxes, slack=6):
    """Merge line boxes that belong to the same visual line."""
    merged = []
    for b in sorted(boxes, key=lambda b: b["y0"]):
        if merged and b["y0"] <= merged[-1]["y1"] + slack:
            last = merged[-1]
            last["y1"] = max(last["y1"], b["y1"])
            last["x0"] = min(last["x0"], b["x0"])
            last["x1"] = max(last["x1"], b["x1"])
            last["text"] = (last["text"] + " " + b["text"]).strip()
        else:
            merged.append(dict(b))
    return merged


def main() -> None:
    if len(sys.argv) < 2:
        raise SystemExit("usage: ocr.py spec.json")
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
    width, height = img.size

    region = spec.get("region")
    rx, ry = 0, 0
    work = img
    if region:
        rx, ry, rx1, ry1 = [int(v) for v in region]
        work = img.crop((rx, ry, rx1, ry1))

    white_point = spec.get("white_point")
    if white_point is not None:
        grey = work.convert("L")
        lo, hi = auto_levels(grey) if white_point is True else (
            int(spec.get("black_point", 0)), int(white_point))
        work = apply_levels(grey, lo, hi)

    scale = float(spec.get("scale", 1) or 1)
    if scale != 1:
        work = work.resize((max(1, int(work.width * scale)),
                            max(1, int(work.height * scale))), Image.LANCZOS)

    tmpdir = tempfile.mkdtemp(prefix="docx-ocr-")
    png = os.path.join(tmpdir, "page.png")
    tsv = os.path.join(tmpdir, "lines.tsv")
    work.save(png)
    try:
        run_powershell(png, tsv, spec.get("lang"))
        raw = open(tsv, encoding="utf-8-sig").read()
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    boxes = []
    for row in raw.splitlines():
        if not row.strip():
            continue
        parts = row.split("\t")
        if len(parts) < 5:
            continue
        try:
            y0, y1, x0, x1 = [float(v) for v in parts[:4]]
        except ValueError:
            continue
        boxes.append({
            "y0": round(y0 / scale) + ry,
            "y1": round(y1 / scale) + ry,
            "x0": round(x0 / scale) + rx,
            "x1": round(x1 / scale) + rx,
            "text": normalize_text(parts[4].strip()),
        })

    # predictable reading order: top to bottom, then left to right
    boxes.sort(key=lambda b: (b["y0"], b["x0"]))

    result = {
        "ok": True,
        "src": os.path.abspath(src),
        "width": width,
        "height": height,
        "lines": len(boxes),
        "text": "\n".join(b["text"] for b in boxes),
        "boxes": boxes,
    }

    if spec.get("figures"):
        merged = merge(boxes)
        min_gap = int(spec.get("min_gap", 45))
        min_w = int(spec.get("min_w", 60))
        min_h = int(spec.get("min_h", 14))
        threshold = int(spec.get("trim_threshold") or
                        (PHOTO_TRIM_THRESHOLD if white_point is not None
                         else DEFAULT_TRIM_THRESHOLD))
        found, prev_end = [], 0
        right = region[2] if region else width
        probe = work if region else img
        probe_mask = probe.convert("L").point(lambda v: 255 if v < threshold else 0)
        rules = rule_columns(probe_mask) if spec.get("drop_rules", True) else set()
        edges = [(b["y0"], b["y1"]) for b in merged] + [(height, height)]
        for y0, y1 in edges:
            if y0 - prev_end > min_gap:
                bb = ink_bbox(probe, (rx, prev_end, right, y0), threshold, rules)
                if bb and (bb[3] - bb[1]) >= min_h and (bb[2] - bb[0]) >= min_w:
                    found.append({"y0": bb[1], "y1": bb[3], "x0": bb[0], "x1": bb[2],
                                  "w": bb[2] - bb[0], "h": bb[3] - bb[1]})
            prev_end = max(prev_end, y1)
        result["figures"] = found

    out = spec.get("out")
    if out:
        if not os.path.isabs(out):
            out = os.path.join(base, out)
        with open(out, "w", encoding="utf-8") as fh:
            json.dump(result, fh, ensure_ascii=False, indent=1)

    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()