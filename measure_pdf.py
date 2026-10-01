"""
measure_pdf.py - exact layout measurements for ICT practical exam printouts.

Reads a PDF and reports, in centimetres / points:
  - page size and orientation
  - text area (margins), header / footer text and alignment
  - columns and the gap between them
  - indented paragraphs and bulleted lists (bullet position and text position)
  - bordered / shaded paragraphs (indent from margins, text-to-border distance,
    border width, shading %)
  - tables (width, offsets from margins, centred or not, rows x columns)
  - shapes / images (width x height)
  - PowerPoint handouts (slides per page, slide frames, shape sizes and
    font sizes converted back to real slide size)
  - font sizes of headings / styled lines

Pages that contain real text (PDF exported from Word / PowerPoint / Access)
are measured from the PDF geometry (accurate to about 0.02 cm).
Scanned pages (images only) are measured from a rendered image
(accurate to about 0.05-0.1 cm, depending on scan quality).

Usage:  python measure_pdf.py file.pdf [file2.pdf ...]
"""
from __future__ import annotations

import io
import re
import statistics
from collections import Counter

import numpy as np
import pdfplumber
import pypdfium2 as pdfium

PT_PER_CM = 72 / 2.54

# A slide is 25.4 x 19.05 cm (4:3) or 33.867 x 19.05 cm (16:9) in PowerPoint
SLIDE_WIDTHS = {4 / 3: 25.4, 16 / 9: 33.867, 16 / 10: 25.4 * 1.0}

BULLET_CHARS = set("•●○◦■□▪▫◆◇★☆✓✔➢➤►▸‣⁃–—-*·")


def cm(pt: float) -> float:
    return pt / PT_PER_CM


def fmt(pt: float) -> str:
    return f"{cm(pt):.2f} cm"


def snippet(text: str, n: int = 60) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    return text if len(text) <= n else text[: n - 1] + "…"


def paper_name(w: float, h: float) -> str:
    a, b = sorted((cm(w), cm(h)))
    if abs(a - 21.0) < 0.3 and abs(b - 29.7) < 0.3:
        name = "A4"
    elif abs(a - 21.59) < 0.3 and abs(b - 27.94) < 0.3:
        name = "Letter"
    else:
        name = "custom size"
    orient = "landscape" if w > h else "portrait"
    return f"{name} {orient} ({cm(w):.1f} x {cm(h):.1f} cm)"


# --------------------------------------------------------------------------- #
#  Geometry helpers
# --------------------------------------------------------------------------- #


def is_bullet_char(ch: dict) -> bool:
    t = ch.get("text", "")
    if not t or t.isspace():
        return False
    font = (ch.get("fontname") or "").lower()
    if any(k in font for k in ("symbol", "wingding", "webding", "opensymbol")):
        return True
    if t in BULLET_CHARS:
        return True
    return 0xF000 <= ord(t[0]) <= 0xF0FF  # private-use glyphs used by Office bullets


def style_of(chars: list[dict]) -> dict:
    real = [c for c in chars if c.get("text", "").strip()]
    if not real:
        real = chars
    sizes = [round(c.get("size", 0), 1) for c in real]
    fonts = Counter((c.get("fontname") or "") for c in real)
    font = fonts.most_common(1)[0][0] if fonts else ""
    fl = font.lower()
    return {
        "size": statistics.median(sizes) if sizes else 0,
        "font": re.sub(r"^[A-Z]{6}\+", "", font),
        "bold": any(k in fl for k in ("bold", "black", "heavy", "semibold")),
        "italic": any(k in fl for k in ("italic", "oblique")),
    }


def raw_lines(page, split_gap: float = 0.8 * PT_PER_CM) -> list[list[dict]]:
    """Group characters into visual lines; split a row where there is a big horizontal gap
    (e.g. between two newspaper columns or between table cells)."""
    chars = [c for c in page.chars if c.get("text") and c.get("upright", True)]
    chars.sort(key=lambda c: ((c["top"] + c["bottom"]) / 2, c["x0"]))
    rows: list[list[dict]] = []
    for c in chars:
        mid = (c["top"] + c["bottom"]) / 2
        h = max(1.0, c["bottom"] - c["top"])
        for row in reversed(rows[-6:]):
            rmid = row[0]["_mid"]
            if abs(mid - rmid) <= 0.35 * h:
                row.append(c)
                break
        else:
            c = dict(c)
            rows.append([c])
            c["_mid"] = mid
            continue
    out = []
    for row in rows:
        row.sort(key=lambda c: c["x0"])
        seg = [row[0]]
        for prev, cur in zip(row, row[1:]):
            if cur["x0"] - prev["x1"] > split_gap and prev["text"].strip() and cur["text"].strip():
                out.append(seg)
                seg = [cur]
            elif cur["x0"] - prev["x1"] > split_gap and not prev["text"].strip():
                # gap after a space: still a split if very large
                out.append(seg)
                seg = [cur]
            else:
                seg.append(cur)
        out.append(seg)
    return out


def build_lines(page) -> list[dict]:
    """Text lines with geometry and style."""
    out = []
    for chars in raw_lines(page):
        chars = [c for c in chars if c.get("text")]
        if not any(c["text"].strip() for c in chars):
            continue
        chars.sort(key=lambda c: c["x0"])
        text = ""
        for prev, cur in zip([None] + chars[:-1], chars):
            if prev is not None and cur["x0"] - prev["x1"] > 0.25 * max(1.0, cur.get("size", 10)) and not text.endswith(" "):
                text += " "
            text += cur["text"]
        text = text.strip()
        ln = {"top": min(c["top"] for c in chars), "bottom": max(c["bottom"] for c in chars)}
        bullet = None
        text_x0 = chars[0]["x0"]
        first_real = next((c for c in chars if c["text"].strip()), None)
        if first_real is not None and is_bullet_char(first_real):
            rest = [c for c in chars if c["x0"] > first_real["x1"] - 0.1 and c["text"].strip()]
            if rest and rest[0]["x0"] - first_real["x1"] > 1.0:
                bullet = {"char": first_real["text"], "x0": first_real["x0"]}
                text_x0 = rest[0]["x0"]
        elif first_real is not None:
            text_x0 = first_real["x0"]
        st = style_of(chars)
        out.append(
            {
                "x0": min(c["x0"] for c in chars if c["text"].strip()) if any(c["text"].strip() for c in chars) else chars[0]["x0"],
                "x1": max(c["x1"] for c in chars if c["text"].strip()),
                "top": ln["top"],
                "bottom": ln["bottom"],
                "text": text,
                "text_x0": text_x0,
                "bullet": bullet,
                **st,
            }
        )
    out.sort(key=lambda l: (round(l["top"], 0), l["x0"]))
    return out


def segments_from_page(page) -> tuple[list, list, list]:
    """Horizontal segments, vertical segments and filled boxes from vector graphics."""
    hs, vs, fills = [], [], []

    def add_h(y, x0, x1, lw):
        if x1 - x0 > 3:
            hs.append({"y": y, "x0": min(x0, x1), "x1": max(x0, x1), "lw": lw})

    def add_v(x, y0, y1, lw):
        if y1 - y0 > 3:
            vs.append({"x": x, "y0": min(y0, y1), "y1": max(y0, y1), "lw": lw})

    for ln in page.lines:
        lw = ln.get("linewidth") or 0.5
        if abs(ln["top"] - ln["bottom"]) < 1.0:
            add_h((ln["top"] + ln["bottom"]) / 2, ln["x0"], ln["x1"], lw)
        elif abs(ln["x0"] - ln["x1"]) < 1.0:
            add_v((ln["x0"] + ln["x1"]) / 2, ln["top"], ln["bottom"], lw)

    for r in page.rects:
        w, h = r["x1"] - r["x0"], r["bottom"] - r["top"]
        lw = r.get("linewidth") or 0.5
        filled = r.get("fill")
        stroked = r.get("stroke")
        # Thin filled rectangles are how Word draws border lines and table rules
        if h <= 2.5 and w > 3:
            add_h((r["top"] + r["bottom"]) / 2, r["x0"], r["x1"], max(h, lw))
            continue
        if w <= 2.5 and h > 3:
            add_v((r["x0"] + r["x1"]) / 2, r["top"], r["bottom"], max(w, lw))
            continue
        if stroked:
            add_h(r["top"], r["x0"], r["x1"], lw)
            add_h(r["bottom"], r["x0"], r["x1"], lw)
            add_v(r["x0"], r["top"], r["bottom"], lw)
            add_v(r["x1"], r["top"], r["bottom"], lw)
        if filled and w > 3 and h > 3:
            fills.append({"x0": r["x0"], "x1": r["x1"], "top": r["top"], "bottom": r["bottom"],
                          "color": r.get("non_stroking_color")})
    return hs, vs, fills


def gray_percent(color) -> float | None:
    """Darkness of a fill colour in percent (0 = white, 100 = black)."""
    if color is None:
        return None
    if isinstance(color, (int, float)):
        color = [color]
    try:
        vals = [float(v) for v in color]
    except (TypeError, ValueError):
        return None
    if len(vals) == 1:
        lum = vals[0]
    elif len(vals) == 3:
        lum = 0.299 * vals[0] + 0.587 * vals[1] + 0.114 * vals[2]
    elif len(vals) == 4:  # CMYK
        c, m, y, k = vals
        lum = 1 - min(1, 0.299 * c + 0.587 * m + 0.114 * y + k)
    else:
        return None
    return round((1 - lum) * 100)


def components(hs: list, vs: list, tol: float = 2.5) -> list[dict]:
    """Group touching horizontal / vertical segments into boxes and grids."""
    segs = [("h", s) for s in hs] + [("v", s) for s in vs]
    n = len(segs)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def bbox(kind, s):
        if kind == "h":
            return s["x0"], s["y"], s["x1"], s["y"]
        return s["x"], s["y0"], s["x"], s["y1"]

    boxes = [bbox(k, s) for k, s in segs]
    for i in range(n):
        ax0, ay0, ax1, ay1 = boxes[i]
        for j in range(i + 1, n):
            bx0, by0, bx1, by1 = boxes[j]
            if ax0 - tol <= bx1 and bx0 - tol <= ax1 and ay0 - tol <= by1 and by0 - tol <= ay1:
                pi, pj = find(i), find(j)
                if pi != pj:
                    parent[pi] = pj
    groups: dict[int, list] = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(segs[i])
    comps = []
    for g in groups.values():
        h = [s for k, s in g if k == "h"]
        v = [s for k, s in g if k == "v"]
        if not h or not v:
            continue
        x0 = min([s["x0"] for s in h] + [s["x"] for s in v])
        x1 = max([s["x1"] for s in h] + [s["x"] for s in v])
        top = min([s["y"] for s in h] + [s["y0"] for s in v])
        bottom = max([s["y"] for s in h] + [s["y1"] for s in v])
        if x1 - x0 < 10 or bottom - top < 6:
            continue
        inner_h = sorted({round(s["y"], 0) for s in h if top + 2 < s["y"] < bottom - 2})
        inner_v = sorted({round(s["x"], 0) for s in v if x0 + 2 < s["x"] < x1 - 2})
        lw = statistics.median([s["lw"] for s in h + v])
        comps.append({"x0": x0, "x1": x1, "top": top, "bottom": bottom,
                      "inner_h": inner_h, "inner_v": inner_v, "lw": lw})
    return comps


def cluster_values(values: list[float], tol: float) -> list[tuple[float, int]]:
    """Cluster 1-D values; returns (centre, count) sorted by centre."""
    if not values:
        return []
    values = sorted(values)
    clusters = [[values[0]]]
    for v in values[1:]:
        if v - clusters[-1][-1] <= tol:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return [(statistics.median(c), len(c)) for c in clusters]


# --------------------------------------------------------------------------- #
#  Layout analysis shared by digital and scanned pages
# --------------------------------------------------------------------------- #


def find_text_frame(lines: list[dict], page_w: float) -> tuple[float, float]:
    """Left and right edge of the main text area (page margins)."""
    lefts = cluster_values([l["x0"] for l in lines], 2.0)
    if not lefts:
        return 0, page_w
    # the left margin is the left-most edge that many lines share
    big = [c for c in lefts if c[1] >= max(2, 0.08 * len(lines))]
    left = min(c[0] for c in (big or lefts))
    rights = sorted(l["x1"] for l in lines)
    right = rights[int(0.97 * (len(rights) - 1))] if rights else page_w
    return left, right


def find_columns(lines: list[dict], left: float, right: float) -> list[tuple[float, float]] | None:
    """Detect newspaper columns: a vertical strip that (almost) no line crosses."""
    if len(lines) < 6 or right - left < 100:
        return None
    width = int(right - left) + 1
    cover = np.zeros(width)
    for l in lines:
        a = max(0, int(l["x0"] - left))
        b = min(width, int(l["x1"] - left) + 1)
        cover[a:b] += 1
    peak = cover.max()
    if peak == 0:
        return None
    low = cover <= max(1, 0.3 * peak)
    best = None
    i = int(0.2 * width)
    while i < int(0.8 * width):
        if low[i]:
            j = i
            while j < width and low[j]:
                j += 1
            dip = cover[i:j].max()
            side = int(3 * PT_PER_CM)
            lm = cover[max(0, i - side):i].mean() if i > 0 else 0
            rm = cover[j:min(width, j + side)].mean() if j < width else 0
            if (j - i) >= 8 and lm >= 2 * max(dip, 1) and rm >= 2 * max(dip, 1) \
                    and (best is None or j - i > best[1] - best[0]):
                best = (i, j)
            i = j
        else:
            i += 1
    if best is None:
        return None
    gap_l, gap_r = left + best[0], left + best[1]
    col2 = [l for l in lines if l["x0"] >= gap_r - 3]
    col1 = [l for l in lines if l["x1"] <= gap_l + 3]
    if len(col1) < 3 or len(col2) < 3:
        return None
    c2_left = cluster_values([l["x0"] for l in col2], 2.0)
    c2_left = max(c2_left, key=lambda c: c[1])[0]
    c1_right = max(l["x1"] for l in col1)
    return [(left, c1_right), (c2_left, right)]


def group_paragraphs(lines: list[dict]) -> list[list[dict]]:
    """Split a column's lines into paragraphs using vertical gaps and bullets."""
    if not lines:
        return []
    lines = sorted(lines, key=lambda l: l["top"])
    heights = [l["bottom"] - l["top"] for l in lines if l["bottom"] > l["top"]]
    lh = statistics.median(heights) if heights else 10
    paras = [[lines[0]]]
    for prev, cur in zip(lines, lines[1:]):
        gap = cur["top"] - prev["bottom"]
        new = gap > 0.55 * lh or cur.get("bullet") or abs(cur.get("size", 0) - prev.get("size", 0)) > 1.5
        # a change of left edge also starts a new paragraph (except a first-line indent on line 2)
        if not new and abs(cur["x0"] - prev["x0"]) > 0.3 * PT_PER_CM and len(paras[-1]) >= 2:
            new = True
        if not new and len(paras[-1]) == 2 and abs(paras[-1][1]["x0"] - paras[-1][0]["x0"]) > 0.3 * PT_PER_CM \
                and abs(cur["x0"] - paras[-1][1]["x0"]) > 0.3 * PT_PER_CM:
            # the "first-line indent" guess was wrong: split before the 2nd line
            second = paras[-1].pop()
            paras.append([second])
        if new:
            paras.append([cur])
        else:
            paras[-1].append(cur)
    return paras


# --------------------------------------------------------------------------- #
#  Digital (text) pages
# --------------------------------------------------------------------------- #


def detect_slide_frames(comps: list[dict], page_w: float, page_h: float) -> list[dict]:
    frames = []
    for c in comps:
        w, h = c["x1"] - c["x0"], c["bottom"] - c["top"]
        if w < 7 * PT_PER_CM or c["inner_h"] or c["inner_v"]:
            continue
        ratio = w / h
        for r, slide_w in ((4 / 3, 25.4), (16 / 9, 33.867), (16 / 10, 27.517)):
            if abs(ratio - r) < 0.035:
                frames.append({**c, "ratio": r, "scale": cm(w) / slide_w, "slide_w": slide_w})
                break
    frames.sort(key=lambda f: (round(f["top"]), f["x0"]))
    return frames


def inside(obj, box, pad=1.5) -> bool:
    return (obj["x0"] >= box["x0"] - pad and obj["x1"] <= box["x1"] + pad
            and obj["top"] >= box["top"] - pad and obj["bottom"] <= box["bottom"] + pad)


def graphic_objects(page, exclude_boxes: list[dict], include_rects: bool = False) -> list[dict]:
    """Images and filled vector shapes (not table / border / frame lines)."""
    objs = []
    for im in page.images:
        objs.append({"kind": "image", "x0": im["x0"], "x1": im["x1"],
                     "top": im["top"], "bottom": im["bottom"]})
    for cv in page.curves:
        objs.append({"kind": "shape", "x0": cv["x0"], "x1": cv["x1"],
                     "top": cv["top"], "bottom": cv["bottom"]})
    if include_rects:
        for rc in page.rects:
            w, h = rc["x1"] - rc["x0"], rc["bottom"] - rc["top"]
            g = gray_percent(rc.get("non_stroking_color"))
            big = any(w > 0.9 * (b["x1"] - b["x0"]) and h > 0.9 * (b["bottom"] - b["top"]) for b in exclude_boxes)
            if rc.get("fill") and w > 0.3 * PT_PER_CM and h > 0.3 * PT_PER_CM and (g is None or g >= 3) and not big:
                objs.append({"kind": "shape", "x0": rc["x0"], "x1": rc["x1"],
                             "top": rc["top"], "bottom": rc["bottom"]})
    # merge overlapping pieces (charts and shapes are often several paths)
    merged: list[dict] = []
    for o in sorted(objs, key=lambda o: (o["top"], o["x0"])):
        if o["x1"] - o["x0"] < 4 and o["bottom"] - o["top"] < 4:
            continue
        for m in merged:
            if o["x0"] <= m["x1"] + 1 and m["x0"] <= o["x1"] + 1 and o["top"] <= m["bottom"] + 1 and m["top"] <= o["bottom"] + 1:
                m["x0"], m["x1"] = min(m["x0"], o["x0"]), max(m["x1"], o["x1"])
                m["top"], m["bottom"] = min(m["top"], o["top"]), max(m["bottom"], o["bottom"])
                m["parts"] += 1
                if o["kind"] == "image":
                    m["kind"] = "image"
                break
        else:
            merged.append({**o, "parts": 1})
    out = []
    for m in merged:
        w, h = m["x1"] - m["x0"], m["bottom"] - m["top"]
        if any(abs(m["x0"] - b["x0"]) < 2 and abs(m["x1"] - b["x1"]) < 2 and abs(m["top"] - b["top"]) < 2 for b in exclude_boxes):
            continue
        if w < 0.3 * PT_PER_CM or h < 0.3 * PT_PER_CM:
            continue
        out.append(m)
    return out


def analyse_digital_page(page, pno: int) -> list[str]:
    W, H = page.width, page.height
    rep = [f"Page {pno}: {paper_name(W, H)}, measured from the PDF itself (exact)."]
    lines = build_lines(page)
    hs, vs, fills = segments_from_page(page)
    comps = components(hs, vs)
    frames = detect_slide_frames(comps, W, H)

    # ---------------- PowerPoint handouts / slides ----------------
    page_ratio = W / H
    single_slide = (not frames) and any(abs(page_ratio - r) < 0.035 for r in (4 / 3, 16 / 9, 16 / 10)) and len(lines) < 60
    if frames or single_slide:
        if single_slide:
            sw = 25.4 if abs(page_ratio - 4 / 3) < 0.035 else (33.867 if abs(page_ratio - 16 / 9) < 0.035 else 27.517)
            frames = [{"x0": 0, "x1": W, "top": 0, "bottom": H, "scale": cm(W) / sw, "slide_w": sw}]
            rep.append("Layout: one full slide per page.")
        else:
            rep.append(f"Layout: PowerPoint handout with {len(frames)} slide frame(s) on this page.")
        for i, f in enumerate(frames, 1):
            s = f["scale"]
            fl = [l for l in lines if inside(l, f)]
            title = snippet(fl[0]["text"], 50) if fl else "(no text)"
            rep.append(f"  Slide frame {i} ('{title}'): printed {cm(f['x1']-f['x0']):.2f} x {cm(f['bottom']-f['top']):.2f} cm "
                       f"= {s*100:.0f}% of the real slide ({f['slide_w']:.2f} cm wide). Real sizes below are converted back to the slide.")
            for l in fl:
                if l["bullet"] or l["x0"] - f["x0"] > 2:
                    pass
            # text positions inside the slide (real cm)
            if fl:
                lefts = cluster_values([l["text_x0"] - f["x0"] for l in fl], 2.0)
                levels = [c for c in lefts]
                if len(levels) > 1:
                    rep.append("    Text start positions from slide left edge (real cm): "
                               + ", ".join(f"{cm(c)/s:.2f} cm ({n} lines)" for c, n in levels))
                for l in fl:
                    tag = []
                    if l["bullet"]:
                        tag.append(f"bullet '{l['bullet']['char']}' at {cm(l['bullet']['x0']-f['x0'])/s:.2f} cm, text at {cm(l['text_x0']-f['x0'])/s:.2f} cm")
                    st = f"{l['size']/s:.0f} pt" + (" bold" if l["bold"] else "") + (" italic" if l["italic"] else "")
                    if l is fl[0] or l["bullet"] or l["italic"] or l["bold"]:
                        rep.append(f"    '{snippet(l['text'], 45)}': {st}" + (f"; {tag[0]}" if tag else ""))
            # footer / slide number inside the frame
            for l in fl:
                rel_y = (l["top"] - f["top"]) / max(1, f["bottom"] - f["top"])
                if rel_y > 0.88 or rel_y < 0.06:
                    pos_x = (l["x0"] + l["x1"]) / 2 - f["x0"]
                    fw = f["x1"] - f["x0"]
                    where = "left" if pos_x < fw * 0.33 else ("right" if pos_x > fw * 0.66 else "centre")
                    edge = "bottom" if rel_y > 0.5 else "top"
                    rep.append(f"    Text near {edge} {where} of slide: '{snippet(l['text'], 40)}'")
            # shapes / images in this frame
            objs = [o for o in graphic_objects(page, [f], include_rects=True) if inside(o, f, pad=2)]
            for o in objs:
                w, h = cm(o["x1"] - o["x0"]) / s, cm(o["bottom"] - o["top"]) / s
                if w > 0.95 * f["slide_w"]:
                    continue
                kind = "picture" if o["kind"] == "image" else ("chart or grouped graphic" if o["parts"] > 6 else "shape")
                txt_left = [l for l in fl if l["x1"] <= o["x0"] + 2 and l["bottom"] > o["top"] and l["top"] < o["bottom"]]
                txt_right = [l for l in fl if l["x0"] >= o["x1"] - 2 and l["bottom"] > o["top"] and l["top"] < o["bottom"]]
                pos = []
                if txt_left:
                    pos.append("to the right of text")
                if txt_right:
                    pos.append("to the left of text")
                rep.append(f"    {kind.capitalize()}: real size {w:.2f} cm wide x {h:.2f} cm high, "
                           f"at {cm(o['x0']-f['x0'])/s:.2f} cm from slide left, {cm(o['top']-f['top'])/s:.2f} cm from slide top"
                           + (f" ({', '.join(pos)})" if pos else ""))
        # page header/footer outside the frames
        outside = [l for l in lines if not any(inside(l, f, 3) for f in frames)]
        for l in outside:
            rep.append(f"  Text outside slide frames: '{snippet(l['text'], 50)}' at {cm(l['x0']):.2f} cm from page left, "
                       f"{cm(l['top']):.2f} cm from page top")
        return rep

    # ---------------- Documents / reports ----------------
    if not lines:
        rep.append("  No text found on this page.")
        return rep
    top_zone, bottom_zone = 0.09 * H, 0.91 * H
    header = [l for l in lines if l["bottom"] < top_zone]
    footer = [l for l in lines if l["top"] > bottom_zone]
    body = [l for l in lines if l not in header and l not in footer]
    tl, tr = find_text_frame(body or lines, W)
    rep.append(f"  Text area: left margin {fmt(tl)}, right edge of text {fmt(tr)} from page left "
               f"(right margin about {fmt(W - tr)}); text width {fmt(tr - tl)}.")

    def align(l, a=None, b=None):
        a = tl if a is None else a
        b = tr if b is None else b
        mid = (l["x0"] + l["x1"]) / 2
        if abs(mid - (a + b) / 2) < 0.6 * PT_PER_CM and l["x0"] > a + PT_PER_CM:
            return "centred"
        if l["x1"] > b - 0.4 * PT_PER_CM and l["x0"] > a + 2 * PT_PER_CM:
            return "right aligned"
        if abs(l["x0"] - a) < 0.4 * PT_PER_CM:
            return "left aligned"
        return f"starting {fmt(l['x0'] - a)} from the left {'margin' if (a, b) == (tl, tr) else 'column edge'}"

    for zone, items in (("Header", header), ("Footer", footer)):
        for l in items:
            rep.append(f"  {zone}: '{snippet(l['text'], 70)}' - {align(l)} ({l['size']:.0f} pt)")

    # tables, boxes, shading
    tables = [c for c in comps if c["inner_h"] or c["inner_v"]]
    boxes = [c for c in comps if not c["inner_h"] and not c["inner_v"]]
    # shaded areas without borders also count as boxes
    for fbox in fills:
        g = gray_percent(fbox["color"])
        if g is None or g < 3 or g > 95:
            continue
        if not any(inside(fbox, b, 4) for b in boxes) and not any(inside(fbox, t, 4) for t in tables):
            if any(inside(l, fbox, 4) for l in body):
                boxes.append({**fbox, "inner_h": [], "inner_v": [], "lw": 0, "shade_only": True})

    columns = find_columns([l for l in body if not any(inside(l, t, 3) for t in tables)], tl, tr)
    if columns:
        (c1l, c1r), (c2l, c2r) = columns
        w_eq = c2r - c2l
        gap_eq = c2l - tl - w_eq
        rep.append(f"  Two columns: column 1 starts at {fmt(c1l)}, column 2 starts at {fmt(c2l)} from page left. "
                   f"Gap between columns: {cm(gap_eq):.2f} cm (assuming equal column widths, as Word's default); "
                   f"gap between the text edges is {cm(c2l - c1r):.2f} cm.")
        col_bounds = [(c1l, c1l + w_eq), (c2l, c2r)]
    else:
        col_bounds = [(tl, tr)]

    def column_of(obj):
        cx = (obj["x0"] + obj["x1"]) / 2
        for a, b in col_bounds:
            if a - 3 <= cx <= b + 3:
                return a, b
        # spans columns -> full width
        return tl, tr

    def full_width(obj):
        return obj["x0"] < (col_bounds[0][1] if len(col_bounds) > 1 else tr) - 5 and obj["x1"] > (col_bounds[-1][0] if len(col_bounds) > 1 else tl) + 5 and len(col_bounds) > 1

    for t in tables:
        a, b = (tl, tr) if full_width(t) else column_of(t)
        left_off, right_off = t["x0"] - a, b - t["x1"]
        rows, cols = len(t["inner_h"]) + 1, len(t["inner_v"]) + 1
        tx = [l for l in lines if inside(l, t, 2)]
        first = snippet(" ".join(l["text"] for l in tx[:2]), 50) if tx else ""
        row_edges = [t["top"]] + sorted(t["inner_h"]) + [t["bottom"]]
        multi = 0
        for r0, r1 in zip(row_edges, row_edges[1:]):
            row_lines = cluster_values([l["top"] for l in tx if r0 - 1 <= l["top"] <= r1 + 1], 2.0)
            if len(row_lines) > 1:
                multi += 1
        centred = abs(left_off - right_off) <= 0.25 * PT_PER_CM
        rep.append(f"  Table ('{first}'): width {fmt(t['x1'] - t['x0'])}, height {fmt(t['bottom'] - t['top'])}; "
                   f"{rows} rows x {cols} columns; {fmt(left_off)} from the left edge and {fmt(right_off)} from the right edge "
                   f"of its {'text area' if (a, b) == (tl, tr) else 'column'} -> {'centred' if centred else 'NOT centred'}; "
                   + ("every row is on one line." if multi == 0 else f"{multi} row(s) wrap onto more than one line.")
                   + f" Border lines {t['lw']:.1f} pt.")

    for bx in boxes:
        tx = [l for l in body if inside(l, bx, 3)]
        if not tx and not bx.get("shade_only") and (bx["x1"] - bx["x0"]) > 0.8 * PT_PER_CM and (bx["bottom"] - bx["top"]) > 0.8 * PT_PER_CM \
                and (bx["x1"] - bx["x0"]) < 0.9 * (tr - tl):
            rep.append(f"  Shape/box (no text): {fmt(bx['x1'] - bx['x0'])} wide x {fmt(bx['bottom'] - bx['top'])} high, "
                       f"{fmt(bx['x0'] - tl)} from the left margin, {fmt(bx['top'])} from page top.")
            continue
        if not tx or (bx["x1"] - bx["x0"]) > 0.98 * (tr - tl) + 30 and (bx["bottom"] - bx["top"]) > 0.6 * H:
            continue
        a, b = (tl, tr) if full_width(bx) else column_of(bx)
        txt_l = min(l["x0"] for l in tx)
        txt_r = max(l["x1"] for l in tx)
        gap_l, gap_t = txt_l - bx["x0"], min(l["top"] for l in tx) - bx["top"]
        gap_r = bx["x1"] - txt_r
        # Word sets the paragraph indent where the text is; the border is drawn outside it.
        ind_l = txt_l - a
        ind_r_box = b - bx["x1"]
        ind_r = ind_r_box + min(gap_l, gap_r) if gap_r > gap_l else ind_r_box + gap_l
        shade = None
        for fbox in fills:
            if inside(fbox, bx, 4) or (abs(fbox["x0"] - bx["x0"]) < 6 and abs(fbox["x1"] - bx["x1"]) < 6 and abs(fbox["top"] - bx["top"]) < 8):
                shade = gray_percent(fbox["color"])
                break
        kind = "Shaded paragraph (no border)" if bx.get("shade_only") else "Bordered paragraph"
        rep.append(f"  {kind} ('{snippet(tx[0]['text'], 55)}'): paragraph text starts {cm(ind_l):.2f} cm from the left "
                   f"{'margin' if (a, b) == (tl, tr) else 'edge of its column'} and the paragraph's right indent is about "
                   f"{cm(ind_r):.2f} cm. Box edges: {cm(bx['x0'] - a):.2f} cm from the left and {cm(ind_r_box):.2f} cm from the right "
                   f"{'margin' if (a, b) == (tl, tr) else 'column edge'}."
                   + ("" if bx.get("shade_only") else f" Distance from text to border: {gap_l:.1f} pt left, {gap_t:.1f} pt top. Border line {bx['lw']:.1f} pt.")
                   + (f" Shading about {shade}% grey." if shade is not None else " No shading."))

    # indented paragraphs and bullets
    in_struct = lambda l: any(inside(l, t, 3) for t in tables) or any(inside(l, bx, 3) for bx in boxes)
    by_col: dict = {}
    for l in body:
        if in_struct(l):
            continue
        by_col.setdefault(column_of(l) if not full_width(l) else (tl, tr), []).append(l)
    bullet_groups = []
    for (a, b), cl in by_col.items():
        for para in group_paragraphs(cl):
            first = para[0]
            if first["bullet"]:
                bullet_groups.append((a, first))
                continue
            if len(para) <= 2 and all(align(l, a, b) == "centred" for l in para):
                continue  # centred heading, reported under styled lines
            rest = para[1:] or para
            left = min(l["x0"] for l in rest)
            fi = first["x0"] - left
            ind = left - a
            if ind > 0.15 * PT_PER_CM or abs(fi) > 0.15 * PT_PER_CM:
                txt = snippet(" ".join(l["text"] for l in para), 55)
                rep.append(f"  Indented paragraph ('{txt}'): left indent {cm(ind):.2f} cm"
                           + (f", first line {cm(fi):+.2f} cm" if abs(fi) > 0.15 * PT_PER_CM else "")
                           + f" from the {'margin' if (a, b) == (tl, tr) else 'column edge'}.")
    if bullet_groups:
        # summarise consecutive bullets with the same geometry
        summary = Counter((round(cm(l["bullet"]["x0"] - a), 2), round(cm(l["text_x0"] - a), 2), l["bullet"]["char"]) for a, l in bullet_groups)
        for (bpos, tpos, ch), n in summary.items():
            ex = [snippet(l["text"], 30) for a, l in bullet_groups if round(cm(l["bullet"]["x0"] - a), 2) == bpos][:2]
            shown = ch if ch.isprintable() and not (0xF000 <= ord(ch[0]) <= 0xF0FF) else "symbol bullet"
            rep.append(f"  Bulleted lines x{n} (e.g. {', '.join(repr(e) for e in ex)}): bullet ('{shown}') at {bpos:.2f} cm, "
                       f"text starts at {tpos:.2f} cm from the margin/column edge.")

    # headings / styled lines (font sizes)
    sizes = Counter(round(l["size"]) for l in body)
    body_size = sizes.most_common(1)[0][0] if sizes else 0
    styled = [l for l in body if (round(l["size"]) != body_size or l["bold"] or l["italic"]) and len(l["text"].strip()) > 2]
    seen = set()
    styled_out = []
    for l in styled:
        key = snippet(l["text"], 40)
        if key in seen:
            continue
        seen.add(key)
        a, b = (tl, tr) if full_width(l) else column_of(l)
        styled_out.append(f"'{key}' {l['size']:.0f} pt{' bold' if l['bold'] else ''}{' italic' if l['italic'] else ''}, {align(l, a, b)}")
    rep.append(f"  Main body text size: {body_size} pt.")
    for s in styled_out[:15]:
        rep.append(f"  Styled line: {s}")

    # pictures / shapes in documents
    for o in graphic_objects(page, comps):
        if inside(o, {"x0": 0, "x1": W, "top": 0, "bottom": top_zone}) or o["top"] > bottom_zone:
            continue
        rep.append(f"  {'Picture' if o['kind']=='image' else 'Shape/graphic'}: {fmt(o['x1']-o['x0'])} wide x {fmt(o['bottom']-o['top'])} high, "
                   f"{fmt(o['x0'] - tl)} from the left margin.")
    return rep


# --------------------------------------------------------------------------- #
#  Scanned (image) pages
# --------------------------------------------------------------------------- #

def runs(mask_1d: np.ndarray, min_len: int) -> list[tuple[int, int]]:
    d = np.diff(np.concatenate(([0], mask_1d.astype(np.int8), [0])))
    starts, ends = np.where(d == 1)[0], np.where(d == -1)[0]
    return [(s, e) for s, e in zip(starts, ends) if e - s >= min_len]


def analyse_scanned_page(pdf_bytes: bytes, index: int, W: float, H: float, pno: int, dpi: int = 150) -> list[str]:
    rep = [f"Page {pno}: {paper_name(W, H)}, SCANNED / image page - measured from the image "
           f"(about ±0.05-0.1 cm; positions are given from the top of the page so they can be matched to the content)."]
    doc = pdfium.PdfDocument(pdf_bytes)
    try:
        pg = doc[index]
        img = pg.render(scale=dpi / 72).to_numpy()
    finally:
        doc.close()
    if img.ndim == 2:
        rgb = np.stack([img] * 3, axis=-1)
    else:
        rgb = img[..., :3]
    # pdfium returns BGR(A)
    b, g, r = rgb[..., 0].astype(int), rgb[..., 1].astype(int), rgb[..., 2].astype(int)
    gray = (0.299 * r + 0.587 * g + 0.114 * b)
    red_ink = (r > 150) & (g < 120) & (b < 120)  # teacher's red marks
    dark = (gray < 140) & ~red_ink
    Hp, Wp = dark.shape
    px_cm = 2.54 / dpi

    from scipy import ndimage

    # 1) solid shapes / pictures / charts: big connected ink areas that are well filled
    lab, nlab = ndimage.label(dark)
    blobs = []
    solid = np.zeros_like(dark)
    if nlab:
        for k, sl in enumerate(ndimage.find_objects(lab), 1):
            if sl is None:
                continue
            hh, ww = (sl[0].stop - sl[0].start), (sl[1].stop - sl[1].start)
            if hh * px_cm >= 0.8 and ww * px_cm >= 0.8:
                fill = (lab[sl] == k).sum() / float(hh * ww)
                if fill >= 0.2:
                    blobs.append({"x0": sl[1].start * 72 / dpi, "x1": sl[1].stop * 72 / dpi,
                                  "top": sl[0].start * 72 / dpi, "bottom": sl[0].stop * 72 / dpi})
                    solid[sl] |= (lab[sl] == k)
    dark_lines = dark & ~solid

    # 2) long thin straight lines -> borders / tables / slide frames
    hlines, vlines = [], []
    min_h = int(2.0 / px_cm)
    for y in range(Hp):
        for s, e in runs(dark_lines[y], min_h):
            hlines.append((y, s, e))
    min_v = int(1.0 / px_cm)
    for x in range(Wp):
        for s, e in runs(dark_lines[:, x], min_v):
            vlines.append((x, s, e))

    def merge(lines_):
        lines_.sort()
        out = []
        for a, s, e in lines_:
            if out and a - out[-1][0] <= 3 and abs(s - out[-1][1]) < 6 and abs(e - out[-1][2]) < 6:
                continue
            out.append((a, s, e))
        return out

    hlines, vlines = merge(hlines), merge(vlines)
    hs = [{"y": y * 72 / dpi, "x0": s * 72 / dpi, "x1": e * 72 / dpi, "lw": 0.5} for y, s, e in hlines]
    vs = [{"x": x * 72 / dpi, "y0": s * 72 / dpi, "y1": e * 72 / dpi, "lw": 0.5} for x, s, e in vlines]
    comps = components(hs, vs, tol=4)

    # light or coloured filled shapes (e.g. a blue sun) that are not border boxes
    sat = rgb.max(axis=-1).astype(int) - rgb.min(axis=-1).astype(int)
    light_ink = ((gray < 215) | (sat > 60)) & ~red_ink & ~solid
    lab2, n2 = ndimage.label(light_ink)
    if n2:
        for k, sl in enumerate(ndimage.find_objects(lab2), 1):
            if sl is None:
                continue
            hh, ww = (sl[0].stop - sl[0].start), (sl[1].stop - sl[1].start)
            if hh * px_cm < 0.8 or ww * px_cm < 0.8:
                continue
            fill = (lab2[sl] == k).sum() / float(hh * ww)
            bb = {"x0": sl[1].start * 72 / dpi, "x1": sl[1].stop * 72 / dpi,
                  "top": sl[0].start * 72 / dpi, "bottom": sl[0].stop * 72 / dpi}
            near_box = any(abs(bb["x0"] - c["x0"]) < 8 and abs(bb["x1"] - c["x1"]) < 8 and abs(bb["top"] - c["top"]) < 8 for c in comps)
            if fill >= 0.04 and not near_box and (bb["x1"] - bb["x0"]) < 0.6 * W:
                if not any(inside(bb, b2, 3) or inside(b2, bb, 3) for b2 in blobs):
                    blobs.append(bb)
                    solid[sl] |= (lab2[sl] == k)

    textmask = dark_lines.copy() & ~solid
    for y, s, e in hlines:
        textmask[max(0, y - 2):y + 3, s:e] = False
    for x, s, e in vlines:
        textmask[s:e, max(0, x - 2):x + 3] = False

    # 3) text blocks (paragraph / column pieces), then lines inside each block
    dil = ndimage.binary_dilation(textmask, structure=np.ones((5, max(3, int(0.45 / px_cm)))))
    blab, nb = ndimage.label(dil)
    tlines = []
    for k, sl in enumerate(ndimage.find_objects(blab), 1):
        if sl is None:
            continue
        sub = textmask[sl] & (blab[sl] == k)
        rowsum = sub.sum(axis=1)
        for rs, re_ in runs(rowsum >= 2, 3):
            band = sub[rs:re_]
            cols = np.where(band.any(axis=0))[0]
            if len(cols) < 4:
                continue
            x0 = sl[1].start + cols[0]
            x1 = sl[1].start + cols[-1]
            top, bot = sl[0].start + rs, sl[0].start + re_
            tlines.append({"x0": x0 * 72 / dpi, "x1": x1 * 72 / dpi, "top": top * 72 / dpi, "bottom": bot * 72 / dpi,
                           "text": "", "bullet": None, "size": (bot - top) * 72 / dpi, "bold": False, "italic": False,
                           "text_x0": x0 * 72 / dpi})
    tlines.sort(key=lambda l: (l["top"], l["x0"]))
    if not tlines and not blobs:
        rep.append("  No text detected on this page.")
        return rep

    def shade_in(box):
        y0, y1 = int(box["top"] * dpi / 72) + 3, int(box["bottom"] * dpi / 72) - 3
        x0, x1 = int(box["x0"] * dpi / 72) + 3, int(box["x1"] * dpi / 72) - 3
        if y1 <= y0 or x1 <= x0:
            return None
        region = gray[y0:y1, x0:x1]
        light = region[region >= 140]
        if light.size < 50:
            return None
        return round((1 - float(np.median(light)) / 255) * 100)

    frames = detect_slide_frames(comps, W, H)
    if frames:
        rep.append(f"  Layout: PowerPoint handout with {len(frames)} slide frame(s).")
        for i, f in enumerate(frames, 1):
            s = f["scale"]
            rep.append(f"  Slide frame {i} (top at {cm(f['top']):.1f} cm from page top): printed {cm(f['x1']-f['x0']):.2f} x "
                       f"{cm(f['bottom']-f['top']):.2f} cm = {s*100:.0f}% of the real slide ({f['slide_w']:.2f} cm wide).")
            fl = [l for l in tlines if inside(l, f, 3)]
            lefts = cluster_values([l["x0"] - f["x0"] for l in fl], 3.0)
            if lefts:
                rep.append("    Text line start positions from slide left edge (real cm): "
                           + ", ".join(f"{cm(c)/s:.2f} cm ({n} lines)" for c, n in lefts))
            for bl in blobs:
                if inside(bl, f, 3) and (bl["x1"] - bl["x0"]) < 0.95 * (f["x1"] - f["x0"]):
                    rep.append(f"    Shape/picture: real size {cm(bl['x1']-bl['x0'])/s:.2f} cm wide x {cm(bl['bottom']-bl['top'])/s:.2f} cm high, "
                               f"{cm(bl['x0']-f['x0'])/s:.2f} cm from slide left, {cm(bl['top']-f['top'])/s:.2f} cm from slide top.")
        return rep

    top_zone, bottom_zone = 0.09 * H, 0.91 * H
    header = [l for l in tlines if l["bottom"] < top_zone]
    footer = [l for l in tlines if l["top"] > bottom_zone]
    body = [l for l in tlines if l not in header and l not in footer]
    tl, tr = find_text_frame(body or tlines, W)
    rep.append(f"  Text area: left margin {fmt(tl)}, right edge of text {fmt(tr)} from page left "
               f"(right margin about {fmt(W - tr)}); text width {fmt(tr - tl)}.")

    def align(l, a, b):
        mid = (l["x0"] + l["x1"]) / 2
        if abs(mid - (a + b) / 2) < 0.6 * PT_PER_CM and l["x0"] > a + PT_PER_CM:
            return "centred"
        if l["x1"] > b - 0.4 * PT_PER_CM and l["x0"] > a + 2 * PT_PER_CM:
            return "right aligned"
        if abs(l["x0"] - a) < 0.4 * PT_PER_CM:
            return "left aligned"
        return f"starting {fmt(l['x0'] - a)} from the left margin"

    for zone, items in (("Header", header), ("Footer", footer)):
        for l in items:
            rep.append(f"  {zone} text ({cm(l['x1']-l['x0']):.1f} cm long): {align(l, tl, tr)}")

    tables = [c for c in comps if c["inner_h"] or c["inner_v"]]
    boxes = [c for c in comps if not c["inner_h"] and not c["inner_v"]]
    columns = find_columns([l for l in body if not any(inside(l, t, 4) for t in tables + boxes)], tl, tr)
    if columns:
        (c1l, c1r), (c2l, c2r) = columns
        w_eq = c2r - c2l
        rep.append(f"  Two columns: column 2 starts at {fmt(c2l)} from page left; gap between columns {cm(c2l - tl - w_eq):.2f} cm "
                   f"(assuming equal widths); gap between text edges {cm(c2l - c1r):.2f} cm.")
        col_bounds = [(c1l, c1l + w_eq), (c2l, c2r)]
    else:
        col_bounds = [(tl, tr)]

    def column_of(obj):
        if len(col_bounds) > 1 and obj["x0"] < col_bounds[0][1] - 5 and obj["x1"] > col_bounds[1][0] + 5:
            return tl, tr
        cx = (obj["x0"] + obj["x1"]) / 2
        for a, b in col_bounds:
            if a - 3 <= cx <= b + 3:
                return a, b
        return tl, tr

    where = lambda a, b: "margin" if (a, b) == (tl, tr) else "column edge"
    for t in tables:
        a, b = column_of(t)
        lo, ro = t["x0"] - a, b - t["x1"]
        rep.append(f"  Table at {cm(t['top']):.1f} cm from page top: width {fmt(t['x1']-t['x0'])}; {len(t['inner_h'])+1} rows x "
                   f"{len(t['inner_v'])+1} columns; {fmt(lo)} from the left and {fmt(ro)} from the right {where(a, b)} -> "
                   f"{'centred' if abs(lo-ro) <= 0.25*PT_PER_CM else 'NOT centred'}.")
    for bx in boxes:
        tx = [l for l in body if inside(l, bx, 4)]
        if not tx:
            continue
        a, b = column_of(bx)
        txt_l = min(l["x0"] for l in tx)
        gap_l = txt_l - bx["x0"]
        ind_r = (b - bx["x1"]) + gap_l
        sh = shade_in(bx)
        rep.append(f"  Bordered box at {cm(bx['top']):.1f}-{cm(bx['bottom']):.1f} cm from page top ({len(tx)} text lines): "
                   f"paragraph text starts {cm(txt_l - a):.2f} cm from the left {where(a, b)}, right indent about {cm(ind_r):.2f} cm; "
                   f"box edges {cm(bx['x0']-a):.2f} cm (left) and {cm(b-bx['x1']):.2f} cm (right); text-to-border gap about {gap_l:.1f} pt"
                   + (f"; shading about {sh}% grey." if sh is not None and sh >= 3 else "; no shading."))
    for bl in blobs:
        rep.append(f"  Shape/picture at {cm(bl['top']):.1f} cm from page top: {cm(bl['x1']-bl['x0']):.2f} cm wide x "
                   f"{cm(bl['bottom']-bl['top']):.2f} cm high.")
    free = [l for l in body if not any(inside(l, x, 4) for x in tables + boxes)]
    for a, b in col_bounds + ([(tl, tr)] if len(col_bounds) > 1 else []):
        cl = sorted([l for l in free if column_of(l) == (a, b)], key=lambda l: l["top"])
        run: list = []

        def flush():
            if run:
                ind = statistics.median([l["x0"] - a for l in run])
                rep.append(f"  Indented text at {cm(run[0]['top']):.1f}-{cm(run[-1]['bottom']):.1f} cm from page top "
                           f"({len(run)} line{'s' if len(run) > 1 else ''}): starts {cm(ind):.2f} cm from the {where(a, b)}.")

        for l in cl:
            ind = l["x0"] - a
            if ind <= 0.2 * PT_PER_CM or align(l, a, b) == "centred":
                flush()
                run = []
                continue
            if run and abs((run[-1]["x0"] - a) - ind) > 0.15 * PT_PER_CM:
                flush()
                run = []
            run.append(l)
        flush()
    return rep


# --------------------------------------------------------------------------- #
#  Public entry point
# --------------------------------------------------------------------------- #

def measure_pdf(pdf_bytes: bytes, name: str = "document.pdf", max_pages: int = 40) -> dict:
    report: list[str] = []
    pages_info = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        n = len(pdf.pages)
        report.append(f"EXACT MEASUREMENTS for '{name}' ({n} page{'s' if n != 1 else ''}), computed by measuring software "
                      f"from the PDF. All positions are horizontal distances unless stated.")
        for i, page in enumerate(pdf.pages[:max_pages]):
            nchars = len([c for c in page.chars if c.get("text", "").strip()])
            img_area = sum((im["x1"] - im["x0"]) * (im["bottom"] - im["top"]) for im in page.images)
            scanned = nchars < 15 and img_area > 0.5 * page.width * page.height
            try:
                if scanned:
                    lines = analyse_scanned_page(pdf_bytes, i, page.width, page.height, i + 1)
                else:
                    lines = analyse_digital_page(page, i + 1)
            except Exception as e:  # never fail the whole document
                lines = [f"Page {i + 1}: could not be measured ({e})."]
            pages_info.append({"page": i + 1, "scanned": scanned})
            report.extend(lines)
        if n > max_pages:
            report.append(f"(Only the first {max_pages} pages were measured.)")
    return {"name": name, "pages": pages_info, "report": "\n".join(report)}


if __name__ == "__main__":
    import sys

    for path in sys.argv[1:]:
        with open(path, "rb") as fh:
            print(measure_pdf(fh.read(), path)["report"])
            print()
