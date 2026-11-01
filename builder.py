#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build a .docx from a JSON layout spec (python-docx).

Usage:  python builder.py spec.json

The spec is a deliberately small, layout-oriented format so an agent can
transcribe a page (screenshot / scan) into it and keep the original typography:
text blocks with 着重号 (emphasis dots), tab-stop columns, shading bands,
tables, pictures and explicit page breaks.
"""
from __future__ import annotations

import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# {d}着重号{/d}  {b}加粗{/b}  {i}斜体{/i}  {u}下划线{/u}
TOKEN = re.compile(r"\{d\}(.*?)\{/d\}|\{b\}(.*?)\{/b\}|\{i\}(.*?)\{/i\}|\{u\}(.*?)\{/u\}", re.S)

WARNINGS: list[str] = []


def warn(msg: str) -> None:
    WARNINGS.append(msg)


def set_fonts(run, east_asia: str, latin: str) -> None:
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:ascii"), latin)
    fonts.set(qn("w:hAnsi"), latin)
    fonts.set(qn("w:eastAsia"), east_asia)


def add_emphasis(run) -> None:
    """着重号 must be w:em/@val=underDot: `dot` draws the marks ABOVE the text."""
    rpr = run._element.get_or_add_rPr()
    for old in rpr.findall(qn("w:em")):
        rpr.remove(old)
    em = OxmlElement("w:em")
    em.set(qn("w:val"), "underDot")
    rpr.append(em)


def shade(paragraph, fill: str) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    ppr.append(shd)


def style_run(run, *, size=12, bold=False, italic=False, underline=False,
              color=None, dot=False, east_asia="宋体", latin="Times New Roman"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    if color:
        run.font.color.rgb = RGBColor.from_string(color.lstrip("#").upper())
    set_fonts(run, east_asia, latin)
    if dot:
        add_emphasis(run)


def emit(paragraph, text, *, size, bold, color, east_asia, latin, base_dot=False):
    """Write text with inline markup into one paragraph."""
    pos = 0
    for m in TOKEN.finditer(text):
        if m.start() > pos:
            seg = text[pos:m.start()]
            r = paragraph.add_run(seg)
            style_run(r, size=size, bold=bold, color=color,
                      east_asia=east_asia, latin=latin, dot=base_dot)
        if m.group(1) is not None:          # 着重号
            r = paragraph.add_run(m.group(1))
            style_run(r, size=size, bold=bold, color=color,
                      east_asia=east_asia, latin=latin, dot=True)
        elif m.group(2) is not None:        # 加粗
            r = paragraph.add_run(m.group(2))
            style_run(r, size=size, bold=True, color=color,
                      east_asia=east_asia, latin=latin, dot=base_dot)
        elif m.group(3) is not None:        # 斜体
            r = paragraph.add_run(m.group(3))
            style_run(r, size=size, bold=bold, italic=True, color=color,
                      east_asia=east_asia, latin=latin, dot=base_dot)
        else:                               # 下划线
            r = paragraph.add_run(m.group(4))
            style_run(r, size=size, bold=bold, underline=True, color=color,
                      east_asia=east_asia, latin=latin, dot=base_dot)
        pos = m.end()
    if pos < len(text):
        r = paragraph.add_run(text[pos:])
        style_run(r, size=size, bold=bold, color=color,
                  east_asia=east_asia, latin=latin, dot=base_dot)


def width_em(text: str) -> float:
    """Rough advance width in em: CJK/fullwidth 1em, ASCII 0.5em."""
    plain = TOKEN.sub(lambda m: next(g for g in m.groups() if g is not None), text)
    total = 0.0
    for ch in plain:
        if ch == "\t":
            continue
        if ord(ch) < 128:
            total += 0.5
        else:
            total += 1.0
    return total


def check_columns(block, text, size, tabs, indent_cm, first_cm, text_width=16.6):
    """Warn when a cell would overflow its tab stop and jump to the default one.

    Model: cell 0 starts at the paragraph indent; cell i>=1 starts at tabs[i-1]
    (a tab jumps forward to that stop). Each cell must end before the next stop.
    """
    cells = text.split("\t")
    if len(cells) == 1:
        return
    em_cm = size * 0.03528
    pos = indent_cm + first_cm
    for i, cell in enumerate(cells):
        if i > 0:
            stop = float(tabs[i - 1]) if i - 1 < len(tabs) else None
            if stop is None:
                warn(f"block {block.get('_i')}: {len(cells)} columns but only "
                     f"{len(tabs)} tab stops -> the extra column lands on a default stop")
                return
            if stop < pos - 0.05:
                warn(f"block {block.get('_i')}: tab stop {stop:.2f}cm is behind the "
                     f"current position {pos:.2f}cm -> the column starts where it already is")
            pos = max(pos, stop)
        end = pos + width_em(cell) * em_cm
        if i + 1 < len(cells):
            nxt = float(tabs[i]) if i < len(tabs) else None
            if nxt is not None and end > nxt + 0.2:
                warn(f"block {block.get('_i')}: column {i + 1} ends at {end:.2f}cm, past "
                     f"the next tab stop {nxt:.2f}cm -> it will be pushed to a default stop")
        elif end > text_width + 0.2:
            warn(f"block {block.get('_i')}: last column ends at {end:.2f}cm, past the "
                 f"{text_width:.1f}cm text width")
        pos = end


def add_paragraph(doc, block, defaults):
    text = block.get("text", "")
    size = float(block.get("size_pt", defaults["size_pt"]))
    line = float(block.get("line_pt", defaults["line_pt"]))
    before = float(block.get("before_pt", 0))
    after = float(block.get("after_pt", 0))
    indent = float(block.get("indent_cm", 0))
    first = float(block.get("first_line_cm", 0))
    tabs = block.get("tabs_cm") or []
    bold = bool(block.get("bold", False))
    color = block.get("color")

    p = doc.add_paragraph()
    fmt = p.paragraph_format
    fmt.left_indent = Cm(indent)
    fmt.first_line_indent = Cm(first)

    rule = block.get("line_rule", "exact")
    # 着重号 hang below the glyph box: give those lines room or Word clips the dots
    if "{d}" in text:
        if rule == "exact" and line < 24:
            line = 24
            rule = "at_least"
        elif rule == "at_least":
            line = max(line, 24)
    fmt.line_spacing = Pt(line)
    if rule == "at_least":
        fmt.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)

    for t in tabs:
        fmt.tab_stops.add_tab_stop(Cm(float(t)))

    align = block.get("align")
    if align:
        p.alignment = ALIGN[align]

    if block.get("fill"):
        shade(p, block["fill"].lstrip("#"))

    emit(p, text, size=size, bold=bold, color=color,
         east_asia=defaults["font"], latin=defaults["latin"])

    check_columns(block, text, size, [float(t) for t in tabs], indent, first)
    return p


def add_title(doc, block, defaults):
    p = doc.add_paragraph()
    fmt = p.paragraph_format
    line = float(block.get("line_pt", 34))
    fmt.line_spacing = Pt(line)
    fmt.space_before = Pt(float(block.get("before_pt", 0)))
    fmt.space_after = Pt(float(block.get("after_pt", 6)))
    fmt.alignment = ALIGN[block.get("align", "center")]
    size = float(block.get("size_pt", 26))
    for seg in block.get("runs", []):
        r = p.add_run(seg.get("text", ""))
        style_run(r, size=float(seg.get("size_pt", size)),
                  bold=bool(seg.get("bold", True)),
                  color=seg.get("color"),
                  east_asia=defaults["font"], latin=defaults["latin"])
    return p


def add_table(doc, block, defaults):
    rows = block.get("rows") or []
    if not rows:
        return None
    cols = max(len(r) for r in rows)
    size = float(block.get("size_pt", defaults["size_pt"]))
    table = doc.add_table(rows=len(rows), cols=cols)
    table.style = block.get("style", "Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    widths = [float(w) for w in (block.get("widths_cm") or [])]
    header = bool(block.get("header", True))

    for r_i, row in enumerate(rows):
        for c_i in range(cols):
            cell = table.cell(r_i, c_i)
            text = row[c_i] if c_i < len(row) else ""
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.alignment = ALIGN[block.get("cell_align", "left")]
            emit(p, str(text), size=size,
                 bold=(header and r_i == 0), color=None,
                 east_asia=defaults["font"], latin=defaults["latin"])
            if widths and c_i < len(widths):
                cell.width = Cm(widths[c_i])
        if header and r_i == 0:
            tr = table.rows[0]._tr
            trPr = tr.get_or_add_trPr()
            tblHeader = OxmlElement("w:tblHeader")
            tblHeader.set(qn("w:val"), "true")
            trPr.append(tblHeader)
    if widths:
        for c_i, w in enumerate(widths):
            for r_i in range(len(rows)):
                table.cell(r_i, c_i).width = Cm(w)
    return table


def add_image(doc, block, defaults):
    src = block["src"]
    if not os.path.isabs(src):
        src = os.path.join(SPEC_DIR, src)
    if not os.path.exists(src):
        raise SystemExit(f"image not found: {src}")
    width = block.get("width_cm")
    picture = doc.add_picture(src, width=Cm(float(width))) if width else doc.add_picture(src)
    p = picture._parent if hasattr(picture, "_parent") else None
    para = doc.paragraphs[-1]
    para.alignment = ALIGN[block.get("align", "center")]
    para.paragraph_format.space_before = Pt(float(block.get("before_pt", 6)))
    para.paragraph_format.space_after = Pt(float(block.get("after_pt", 0)))
    cap = block.get("caption")
    if cap:
        cp = doc.add_paragraph()
        cp.alignment = ALIGN["center"]
        cp.paragraph_format.space_before = Pt(2)
        cp.paragraph_format.space_after = Pt(float(block.get("caption_after_pt", 8)))
        r = cp.add_run(cap)
        style_run(r, size=float(block.get("caption_size_pt", 10.5)),
                  color=block.get("caption_color"),
                  east_asia=defaults["font"], latin=defaults["latin"])
        return cp
    return para


def add_pagebreak_carrier(doc) -> None:
    """A 1pt paragraph that carries only a page break (used after tables)."""
    p = doc.add_paragraph()
    fmt = p.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = Pt(1)
    fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    r = p.add_run()
    r.font.size = Pt(1)
    r.add_break(WD_BREAK.PAGE)


def build(spec: dict) -> dict:
    doc = Document()
    body = spec.get("body") or {}
    defaults = {
        "font": body.get("font", "宋体"),
        "latin": body.get("latin", "Times New Roman"),
        "size_pt": float(body.get("size_pt", 12)),
        "line_pt": float(body.get("line_pt", 20)),
    }

    normal = doc.styles["Normal"]
    normal.font.name = defaults["latin"]
    normal.font.size = Pt(defaults["size_pt"])
    rpr = normal.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    fonts.set(qn("w:eastAsia"), defaults["font"])
    fonts.set(qn("w:ascii"), defaults["latin"])
    fonts.set(qn("w:hAnsi"), defaults["latin"])

    page = spec.get("page") or {}
    section = doc.sections[0]
    section.page_width = Cm(float(page.get("width_cm", 21.0)))
    section.page_height = Cm(float(page.get("height_cm", 29.7)))
    margin = page.get("margin_cm") or {}
    section.top_margin = Cm(float(margin.get("top", 2.2)))
    section.bottom_margin = Cm(float(margin.get("bottom", 2.2)))
    section.left_margin = Cm(float(margin.get("left", 2.2)))
    section.right_margin = Cm(float(margin.get("right", 2.2)))
    section.header_distance = Cm(float(margin.get("header", 1.5)))
    section.footer_distance = Cm(float(margin.get("footer", 1.5)))

    counts = {"paragraphs": 0, "tables": 0, "images": 0, "pagebreaks": 0}
    last_para = None

    for i, block in enumerate(spec.get("blocks") or []):
        block["_i"] = i
        # models often write a literal backslash-t inside JSON; treat it as a tab
        if isinstance(block.get("text"), str):
            text = block["text"].replace("\\t", "\t").replace("\\n", "\n")
            # "A.\t" is an option label, not a column: a tab there jumps to the
            # first column stop and pushes the whole row right (classic bug).
            text = re.sub(r"^([A-D])\.\t+", r"\1. ", text)
            block["text"] = text
        kind = block.get("type", "text")

        if kind == "pagebreak":
            if last_para is not None and last_para.runs:
                last_para.runs[-1].add_break(WD_BREAK.PAGE)
            else:
                add_pagebreak_carrier(doc)
            counts["pagebreaks"] += 1
            continue

        if kind == "title":
            last_para = add_title(doc, block, defaults)
            counts["paragraphs"] += 1
        elif kind == "table":
            add_table(doc, block, defaults)
            counts["tables"] += 1
            last_para = None          # a table is not addressable as a paragraph
        elif kind == "image":
            last_para = add_image(doc, block, defaults)
            counts["images"] += 1
        else:
            last_para = add_paragraph(doc, block, defaults)
            counts["paragraphs"] += 1

    out = spec.get("out") or "out.docx"
    if not os.path.isabs(out):
        out = os.path.join(SPEC_DIR, out)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    doc.save(out)
    return {"ok": True, "path": os.path.abspath(out), "counts": counts,
            "warnings": WARNINGS}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("usage: builder.py spec.json")
    spec_path = os.path.abspath(sys.argv[1])
    SPEC_DIR = os.path.dirname(spec_path)
    with open(spec_path, "r", encoding="utf-8") as fh:
        spec = json.load(fh)
    result = build(spec)
    sys.stdout.write(json.dumps(result, ensure_ascii=False))