#!/usr/bin/env python3
"""
add_examples.py — Add textbook score Examples to chapter JS/PDF files.

Usage:
  python3 scripts/add_examples.py --chapters 4 20 31        # pilot
  python3 scripts/add_examples.py --chapters 4 --dry-run    # no write
  python3 scripts/add_examples.py --all                     # all 36 chapters

Phases per chapter:
  1. Discover: find which PDF page each EX AMPLE caption is on
  2. Extract:  render page, crop score image, save PNG
  3. Patch:    insert example slides into chapter JS, update TOC
  4. Rebuild:  node chXX.js → soffice → PDF
"""

import argparse
import math
import os
import re
import subprocess
import sys
from pathlib import Path

# ── Constants ─────────────────────────────────────────────────────────────────

ROOT = Path(__file__).parent.parent          # repo root
PDF  = str(ROOT / "A HISTORY of WESTERN TENTН MUSIC.pdf")

# PDF page offset: printed page N → PDF page N + 37
PDF_OFFSET = 37

# Split a score image into multiple slides when it would be too narrow to read.
# For side-by-side layout: rendered W = 4.55" × asp; split if asp < 0.70 (W < 3.2")
SPLIT_THRESHOLD = 0.70

# Manual y_end_pct overrides for examples where auto-detection includes body text.
# Key: (chapter_int, example_num_str)  Value: y_end_pct (fraction of PDF page height)
MANUAL_Y_END = {
    # ── Ch01 ─────────────────────────────────────────────────────────────────
    (1,  '4'):  0.460,   # Seikilos: score + 2-line English translation; stop before FIGURE 1.10 photo at 75%
    # ── Ch02 ─────────────────────────────────────────────────────────────────
    (2,  '1'):  0.340,   # Deus Deus meus 4 verse phrases: body text at 35.6%
    (2,  '2'):  0.620,   # Viderunt Solesmes notation + English translation; body text "editions" at 65.2%
    (2,  '3'):  0.640,   # Viderunt modern notation; body text "small notes" at 65.2%
    (2,  '5'):  0.720,   # Church modes diagram; body text "listeners may find" at 73.4%
    (2,  '6'):  0.290,   # Ut queant laxis + translation; body text "Greek names" at 30.7%
    (2,  '8'):  0.420,   # Viderunt solmization (2nd of two on page); body text "required shifting" at 42.2%
    # ── Ch08 ─────────────────────────────────────────────────────────────────
    (8,  '2'):  0.46,    # Carol burden: score 2 sections, body text "elaborated" at ~48%
    (8,  '3'):  0.34,    # Dunstable cantus vs chant: score ends ~31%, body text at 36%
    (8,  '4'):  0.795,   # Binchois chanson: label at 58%, score+translation ends ~79%, body text at 80%
    (8,  '6'):  0.472,   # Se la face ay pale: score + translation, body text "a ballade" just after
    (8,  '7'):  0.96,    # Missa Gloria a-d: full page (all 4 versions)
    # ── Ch09 ─────────────────────────────────────────────────────────────────
    (9,  '1'):  0.57,    # Busnoys a+b: body text "Popularity" at ~60%
    (9,  '2'):  0.32,    # Range comparison diagram: body text "The bassus" at ~34%
    (9,  '3'):  0.690,   # Ockeghem prolationum: score b ends ~68%, "Example 9.3b." body text at 70.1%
    (9,  '4'):  0.42,    # Obrecht Gloria: score ends ~40%, body text at 43%
    (9,  '5'):  0.980,   # Isaac Puer natus: 4-voice motet b section extends to 98%, footer at 98.5%
    (9,  '6'):  0.36,    # Carnival song: score+translation, body text at actual ~38%
    (9,  '7'):  0.53,    # Isaac Innsbruck: 2 systems, Josquin bio at ~58%
    (9,  '8'):  0.32,    # Josquin Mille regretz: 1 system, body text at ~35%
    (9,  '9'):  0.734,   # Josquin Ave Maria a-c: stop before "projection of the text" body text
    (9,  '10'): 0.79,    # Févin Missa a+b: body text "Lord God..." at ~80%
    # ── Ch10 ─────────────────────────────────────────────────────────────────
    (10, '2'):  0.79,    # Willaert a+b: body text at ~82%
    (10, '3'):  0.46,    # Rore Da le belle: score ends ~44%, body text at 48%
    (10, '5'):  0.63,    # Gesualdo Io parto: 2 systems, body text "Carlo Gesualdo" at ~67%
    (10, '6'):  0.345,   # Sermisy Tant que vivray: score + translation, exclude body text
    # ── Ch05 ─────────────────────────────────────────────────────────────────
    (5,  '7'):  0.83,    # stop before "in Example 5.7" body text (3 discant systems)
    (5,  '8'):  0.3507,  # clausulae: 2 systems only, body text at 51%
    (5,  '9'):  0.7008,  # Pérotin 4-voice: 3 systems, body text at 75%
    (5,  '13'): 0.4480,  # Adam de la Halle motet: keep score + translation
    (5,  '14'): 0.4251,  # Petrus de Cruce: keep score + translation
    (5,  '15'): 0.5885,  # Cadence forms: very short score, body text at 22%
    (22, '1'):  0.96,    # all a–e sections fill the page (was cut at a,b only)
    (23, '3'):  0.96,    # all a–e sections fill the page (was cut at a,b only)
    (23, '4'):  0.3548,  # Haydn 104 finale: 2 systems, body text at 35%
    (23, '6'):  0.46,    # stop before "FREELANCING" body text (at 48%); include both a,b
    (23, '7'):  0.635,   # stop before "Contrasting styles" body text (at 64%); 3 score systems
    (23, '8'):  0.2020,  # Mozart Jupiter theme: 1 system, body text at 21%
    (25, '1'):  0.96,    # all a–d sections fill the page (was cut at a only)
    (25, '3'):  0.960,   # Schumann Dichterliebe: label at 80.7%, single system fills to ~96%, footer 98.5%
    (25, '7'):  0.96,    # all a–d sections fill the page (was cut at a only)
    (25, '8'):  0.960,   # Chopin Mazurka: a at 56.9%, b at 70.9%, both sections fill to footer
    (27, '2'):  0.700,   # Rossini Una voce: a+b + translations to 66.6%, body text at 71.7%
    (33, '2'):  0.960,   # Schoenberg Op.25: a/b/c sections at 13.4%/40.2%/73.0%, no body text on page
    (33, '3'):  0.2557,  # Schoenberg row: row diagram only, body text at 30%
    (33, '7'):  0.2498,  # Stravinsky Petrushka: short passage, body text at 20%
    (33, '11'): 0.96,    # all a–g sections fill the page (was cut at a,b only)
    (33, '12'): 0.2393,  # Bartók xylophone palindrome: 1 system, body text at 18%
    (33, '15'): 0.820,   # Ives Concord Alcotts: a-e sections at 13%/28%/45%/61%, body text at 83.2%
}

# x_start override for two-column pages where body text is in the left column
MANUAL_X_START = {
    (13, '2'): 0.49,   # right column only: body text on left, EXAMPLE 13.2 on right
    (1,  '4'): 0.52,   # right column only: Seikilos epitaph + translation in right col
}

# Chapter page ranges (printed-book pagination)
CH_PAGES = {
     1: ( 4,  19),  2: ( 20,  41),  3: ( 42,  62),  4: ( 63,  79),
     5: ( 80, 105),  6: (106, 132),  7: (136, 158),  8: (159, 179),
     9: (180, 204), 10: (205, 228), 11: (229, 253), 12: (254, 277),
    13: (278, 296), 14: (297, 316), 15: (317, 338), 16: (339, 370),
    17: (371, 401), 18: (402, 423), 19: (424, 453), 20: (454, 470),
    21: (471, 493), 22: (494, 513), 23: (514, 553), 24: (554, 579),
    25: (580, 617), 26: (618, 645), 27: (646, 670), 28: (671, 710),
    29: (711, 730), 30: (731, 755), 31: (756, 769), 32: (770, 803),
    33: (804, 847), 34: (848, 868), 35: (869, 897), 36: (898, 918),
    37: (919, 953), 38: (954, 989), 39: (990,1020),
}

# Chapters with no Examples — skip entirely
NO_EXAMPLES = {7, 36, 39}

# ── Phase 1: Discover ─────────────────────────────────────────────────────────

def discover_examples(ch: int) -> dict:
    """
    Scan chapter PDF pages to find which page each EX AMPLE N.M caption is on.
    Returns {ex_num_str: pdf_page_int}, e.g. {'1': 100, '2': 103}.
    """
    p_start, p_end = CH_PAGES[ch]
    pdf_start = p_start + PDF_OFFSET
    pdf_end   = p_end   + PDF_OFFSET
    found = {}

    for page in range(pdf_start, pdf_end + 1):
        result = subprocess.run(
            ["pdftotext", "-layout", "-f", str(page), "-l", str(page), PDF, "-"],
            capture_output=True, text=True
        )
        # Match "EX AMPLE N.M" where N = chapter number
        for m in re.finditer(rf'EX\s*AMPLE\s+{ch}\.(\d+)', result.stdout):
            num = m.group(1)
            if num not in found:
                found[num] = page

    return found


# ── Phase 2: Extract ──────────────────────────────────────────────────────────

def get_label_y_pct(ch: int, ex_num: str, pdf_page: int) -> float:
    """
    Return the Y position of the EX AMPLE label as a fraction of page height.
    Uses pdftotext -bbox (HTML word-level bounding boxes).
    """
    result = subprocess.run(
        ["pdftotext", "-bbox", "-f", str(pdf_page), "-l", str(pdf_page), PDF, "-"],
        capture_output=True, text=True
    )
    html = result.stdout

    # Page height from <page width="W" height="H">
    h_match = re.search(r'height="([\d.]+)"', html)
    page_h = float(h_match.group(1)) if h_match else 792.0

    # Find the "EX" word, then confirm "AMPLE" follows, then "N.M"
    # bbox HTML format: <word xMin="x" yMin="y" xMax="x2" yMax="y2">TEXT</word>
    words = re.findall(r'yMin="([\d.]+)"[^>]*>([^<]+)</word>', html)

    for i, (y, word) in enumerate(words):
        text = word.strip()
        if text == "EX" and i + 2 < len(words):
            next_word  = words[i+1][1].strip()
            after_word = words[i+2][1].strip()
            if next_word == "AMPLE" and after_word == f"{ch}.{ex_num}":
                return float(y) / page_h
        # Also handle "EXAMPLE" without space (some PDF encodings)
        if text in ("EXAMPLE", "EX AMPLE") and i + 1 < len(words):
            after_word = words[i+1][1].strip()
            if after_word == f"{ch}.{ex_num}":
                return float(y) / page_h

    # Fallback: search for "ch.num" label position directly
    for y, word in words:
        if word.strip() == f"{ch}.{ex_num}":
            return float(y) / page_h

    return 0.05  # fallback: near top of page


def find_example_bottom_y_pct(pdf_page: int, y_label_abs: float,
                               page_h: float) -> float:
    """
    Auto-detect where the example figure ends by finding the paragraph break
    after the score image.

    Strategy:
    1. Collect text lines (groups of words at same Y) after y_label_abs.
    2. Find the first BIG gap (>80pt) = the score image area.
    3. After that gap, collect poem/caption lines.
    4. Find the next gap > 22pt = paragraph break before body text.
    5. Crop bottom = y just before that paragraph break.
    """
    result = subprocess.run(
        ["pdftotext", "-bbox", "-f", str(pdf_page), "-l", str(pdf_page), PDF, "-"],
        capture_output=True, text=True
    )
    html = result.stdout

    # Collect all (y, words_on_line) pairs
    word_hits = re.findall(r'yMin="([\d.]+)"[^>]*>([^<]+)</word>', html)
    if not word_hits:
        return 0.92

    # Group words into lines (within 3pt)
    lines = {}
    footer_y = page_h * 0.93   # running heads/footers live here — exclude
    for y_str, word in word_hits:
        y = float(y_str)
        if y <= y_label_abs + 4:   # +4pt tolerance: exclude label line itself
            continue
        if y >= footer_y:          # exclude running head / page number area
            continue
        # Find the line bucket
        bucket = None
        for existing_y in list(lines.keys()):
            if abs(existing_y - y) <= 3:
                bucket = existing_y
                break
        if bucket is None:
            lines[y] = []
            bucket = y
        lines[bucket].append(word.strip())

    if not lines:
        return 0.92

    sorted_ys = sorted(lines.keys())

    # Step A: The score image gap is BETWEEN y_label_abs and sorted_ys[0].
    # If that gap > 80pt, the score image is there and sorted_ys[0] starts
    # poem/caption text after the score.
    if sorted_ys and (sorted_ys[0] - y_label_abs) > 80:
        # Score image sits between label and first text.
        # Now find the paragraph break in the text that follows.
        for i in range(len(sorted_ys) - 1):
            gap = sorted_ys[i+1] - sorted_ys[i]
            if gap > 22:
                bottom_y = sorted_ys[i] + gap * 0.4
                return min(bottom_y / page_h, 0.97)
        # No paragraph break → text right after score IS body text.
        # Crop just before the first text line (at score_image_gap_end).
        return max(0.10, (sorted_ys[0] - 8) / page_h)

    # Step B: First text appears close to the label (no big initial gap).
    # The score might be embedded differently. Find any gap > 80pt then > 22pt.
    score_image_gap_end = None
    for i in range(len(sorted_ys) - 1):
        if sorted_ys[i+1] - sorted_ys[i] > 80:
            score_image_gap_end = sorted_ys[i+1]
            break

    if score_image_gap_end is None:
        return 0.90  # no score gap found; fallback

    after_score = [y for y in sorted_ys if y >= score_image_gap_end]
    FOOTER_Y = page_h * 0.93   # lines below this are running heads/footers — ignore
    for i in range(len(after_score) - 1):
        if after_score[i] >= FOOTER_Y:
            break
        gap = after_score[i+1] - after_score[i]
        if gap > 22:
            bottom_y = after_score[i] + gap * 0.4
            return min(bottom_y / page_h, 0.97)

    # No paragraph break in body text → text immediately following IS body text.
    # Crop at score_image_gap_end (just before first body text line).
    return max(0.10, (score_image_gap_end - 8) / page_h)


def crop_example(ch: int, ex_num: str, pdf_page: int, out_path: str,
                 y_end_pct: float = None) -> tuple:
    """
    Render pdf_page at 200dpi, crop from just above the EX AMPLE label to
    y_end_pct (or 0.97 if None). White-out top 65px (running head).
    Returns (width_px, height_px) of saved image.
    """
    from PIL import Image, ImageDraw

    # Render the page
    tmp_prefix = f"/tmp/ex_render_{ch}_{ex_num}"
    subprocess.run(
        ["pdftoppm", "-r", "200", "-png",
         "-f", str(pdf_page), "-l", str(pdf_page), PDF, tmp_prefix],
        capture_output=True
    )
    # pdftoppm names output as prefix-NNNN.png
    render_path = f"{tmp_prefix}-{pdf_page:04d}.png"
    if not os.path.exists(render_path):
        # Try alternate naming (single page → -1.png)
        render_path = f"{tmp_prefix}-1.png"
    if not os.path.exists(render_path):
        # Try without zero-padding
        candidates = sorted(Path("/tmp").glob(f"ex_render_{ch}_{ex_num}-*.png"))
        if candidates:
            render_path = str(candidates[-1])
        else:
            raise FileNotFoundError(f"No render found for ch{ch} ex{ex_num} page{pdf_page}")

    img = Image.open(render_path)
    W, H = img.size

    y1_pct = get_label_y_pct(ch, ex_num, pdf_page)
    y1 = max(0, int((y1_pct - 0.02) * H) - 20)

    if y_end_pct is None:
        # Auto-detect the example bottom; get actual page height from bbox
        bbox_result = subprocess.run(
            ["pdftotext", "-bbox", "-f", str(pdf_page), "-l", str(pdf_page), PDF, "-"],
            capture_output=True, text=True
        )
        hm = re.search(r'height="([\d.]+)"', bbox_result.stdout)
        page_h_pts = float(hm.group(1)) if hm else 762.0
        y_label_abs = y1_pct * page_h_pts
        y_end_pct = find_example_bottom_y_pct(pdf_page, y_label_abs, page_h_pts)

    y2 = int(y_end_pct * H)
    y2 = min(y2, H)

    # Crop with small horizontal margins (remove page gutters)
    # MANUAL_X_START overrides left edge for two-column pages
    x_start = MANUAL_X_START.get((ch, ex_num), 0.04)
    left  = int(x_start * W)
    right = int(0.96 * W)
    cropped = img.crop((left, y1, right, y2))

    # White-out running head (top 65px)
    draw = ImageDraw.Draw(cropped)
    draw.rectangle([0, 0, cropped.width, 65], fill='white')

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    cropped.save(out_path)
    os.unlink(render_path)
    return cropped.size


def find_natural_splits(img_path: str, n_parts: int) -> list:
    """
    Find n_parts-1 natural vertical split points by locating the whitest
    (most blank) horizontal row near each N-th fraction of the image height.
    Returns list of y-pixel positions to split at.
    """
    from PIL import Image
    img = Image.open(img_path).convert('L')
    W, H = img.size
    split_ys = []

    for k in range(1, n_parts):
        target = H * k // n_parts
        search_start = max(0, target - H // 7)
        search_end   = min(H, target + H // 7)

        try:
            import numpy as np
            arr = np.array(img)[search_start:search_end]
            row_means = arr.mean(axis=1)
            best_local = int(row_means.argmax())
            split_ys.append(search_start + best_local)
        except ImportError:
            # Fallback: scan rows manually
            best_y, best_mean = target, 0.0
            for y in range(search_start, search_end):
                row_pixels = list(img.crop((0, y, W, y + 1)).getdata())
                m = sum(row_pixels) / len(row_pixels)
                if m > best_mean:
                    best_mean, best_y = m, y
            split_ys.append(best_y)

    return split_ys


def split_image_vertically(img_path: str, n_parts: int) -> list:
    """
    Split image into n_parts at natural break points.
    Saves part files as <base>a.png, <base>b.png, …
    Deletes the original.
    Returns list of (part_path, W_px, H_px).
    """
    from PIL import Image
    img = Image.open(img_path)
    W, H = img.size
    base = img_path[:-4]   # strip .png

    split_ys = find_natural_splits(img_path, n_parts)
    boundaries = [0] + split_ys + [H]

    parts = []
    for i in range(n_parts):
        y1, y2 = boundaries[i], boundaries[i + 1]
        part = img.crop((0, y1, W, y2))
        suffix = chr(ord('a') + i)
        part_path = f"{base}{suffix}.png"
        part.save(part_path)
        parts.append((part_path, W, y2 - y1))

    os.unlink(img_path)
    return parts


def extract_all_examples(ch: int, pages: dict, dry_run: bool = False) -> dict:
    """
    Extract all Examples for a chapter.
    pages = {ex_num_str: pdf_page_int}
    Returns {ex_num_str: (W_px, H_px)}
    """
    out_dir = ROOT / "examples" / f"ch{ch:02d}"
    dims = {}

    # Group examples by PDF page to detect adjacency
    by_page = {}
    for num, page in sorted(pages.items(), key=lambda x: (x[1], int(x[0]))):
        by_page.setdefault(page, []).append(num)

    # Build y_end_pct map: if two examples on same page, split at midpoint
    y_end = {}
    for page, nums in by_page.items():
        if len(nums) == 1:
            y_end[nums[0]] = None  # full page crop
        else:
            # Get y positions for all examples on this page
            y_pcts = [get_label_y_pct(ch, n, page) for n in nums]
            for i, num in enumerate(nums):
                if i + 1 < len(nums):
                    # End just before the next example starts
                    y_end[num] = y_pcts[i+1] - 0.01
                else:
                    y_end[num] = None

    for num in sorted(pages.keys(), key=int):
        page = pages[num]
        out_path = str(out_dir / f"ex{ch:02d}_{num}.png")

        if dry_run:
            print(f"  [dry-run] would crop ch{ch} ex{num} from page {page} → {out_path}")
            dims[num] = (800, 300)  # fake dims for dry-run
            continue

        print(f"  Cropping ex{ch}.{num} from PDF page {page}…", end=" ", flush=True)
        try:
            manual = MANUAL_Y_END.get((ch, num))
            w, h = crop_example(ch, num, page, out_path, y_end_pct=manual if manual is not None else y_end.get(num))
            asp = w / h if h > 0 else 2.0
            print(f"{w}×{h}px (asp={asp:.2f})", end="")

            if asp < SPLIT_THRESHOLD:
                # Too tall/narrow — split into multiple pages
                n_parts = max(2, math.ceil(1.3 / asp))  # target part_asp ≥ 1.3
                n_parts = min(n_parts, 4)
                parts = split_image_vertically(out_path, n_parts)
                dims[num] = [(pw, ph) for _, pw, ph in parts]
                print(f" → split into {n_parts} pages")
            else:
                dims[num] = (w, h)
                print(f" → {os.path.basename(out_path)}")
        except Exception as e:
            print(f"ERROR: {e}")
            dims[num] = (800, 300)

    return dims


# ── Phase 3: Patch JS ─────────────────────────────────────────────────────────

# Chapter JS filename patterns
def find_js_file(ch: int) -> Path:
    """Find the JS source file for a chapter."""
    for path in sorted(ROOT.glob(f"ch{ch:02d}_*.js")):
        return path
    raise FileNotFoundError(f"No JS file found for chapter {ch}")


def read_c_palette(js: str) -> dict:
    """Extract the C = { ... } color palette from a chapter JS file."""
    m = re.search(r'const C = \{([^}]+)\}', js, re.DOTALL)
    if not m:
        return {}
    palette = {}
    for line in m.group(1).split('\n'):
        km = re.match(r'\s*(\w+)\s*:\s*["\']([0-9A-Fa-f]{6})["\']', line)
        if km:
            palette[km.group(1)] = km.group(2)
    return palette


def safe_color(palette: dict, *keys) -> str:
    """Return the first key that exists in palette, else a neutral fallback."""
    for k in keys:
        if k in palette:
            return f"C.{k}"
    # Return hex literal fallback
    fallbacks = {
        'ivory':  '"E8DFC0"',
        'bronze': '"9A7840"',
        'slate':  '"2A2A2A"',
        'cream':  '"F5ECD8"',
    }
    for k in keys:
        if k in fallbacks:
            return fallbacks[k]
    return '"888888"'


def find_insertion_point(js: str) -> int:
    """Return char index just before the Key Terms slide comment."""
    # Standard comment format across all chapters
    m = re.search(r'\n// ── SLIDE[^\n]*[Kk]ey [Tt]erm', js)
    if m:
        return m.start() + 1
    # Fallback: before pres.writeFile
    i = js.rfind('\npres.writeFile')
    if i >= 0:
        return i + 1
    return len(js)


def count_slides_before(js: str, insertion_idx: int) -> int:
    """Count how many slides exist before the insertion point (exclude function defs)."""
    snippet = js[:insertion_idx]
    # Match calls but not function definitions (function darkSlide(...))
    all_matches = re.findall(r'(?:darkSlide|lightSlide)\s*\(', snippet)
    def_matches = re.findall(r'function\s+(?:darkSlide|lightSlide)\s*\(', snippet)
    return len(all_matches) - len(def_matches)


def build_example_slide_js(ch: int, ex_num: str, palette: dict,
                            img_dims: tuple,
                            title_zh: str, subtitle_en: str,
                            explanation_title: str, explanation_zh: str,
                            img_suffix: str = '') -> str:
    """
    Generate one pptxgenjs slide block for a textbook example.
    - aspect >= 1.3: stacked layout (score top, text panel below)
    - aspect <  1.3: side-by-side layout (score left, text panel right)
    """
    W_px, H_px = img_dims
    asp = W_px / H_px if H_px > 0 else 2.0

    c_bronze = safe_color(palette, 'bronze', 'gold', 'tan')
    c_slate  = safe_color(palette, 'slate', 'panel', 'darkBg')
    c_ivory  = safe_color(palette, 'ivory', 'cream', 'lightText')
    img_key  = f"ex{ch:02d}_{ex_num}{img_suffix}"

    def esc(s):
        return s.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n')

    header = f"""
// ── EXAMPLE {ch}.{ex_num} ────────────────────────────────────────────────────────
{{
  const s = darkSlide(pres);
  topBar(s); bottomBar(s);
  s.addText("{esc(title_zh)}", {{ x:0.4, y:0.14, w:9.2, h:0.38, fontSize:18, bold:true, color:C.gold, fontFace:"Georgia", align:"center" }});
  s.addText("{esc(subtitle_en)}", {{ x:0.4, y:0.56, w:9.2, h:0.28, fontSize:12, italic:true, color:{c_ivory}, fontFace:"Calibri", align:"center" }});"""

    if asp >= 1.3:
        # ── Stacked layout ──────────────────────────────────────────────────
        H_img = 2.65 if asp >= 2.0 else 3.05
        W_img = min(9.0, round(H_img * asp, 2))
        H_img = round(W_img / asp, 2)  # natural height — no whitespace for very wide images
        y0 = 1.05
        x0 = round(0.5 + (9.0 - W_img) / 2, 2)
        fx = round(x0 - 0.08, 2); fy = round(y0 - 0.06, 2)
        fw = round(W_img + 0.16, 2); fh = round(H_img + 0.12, 2)
        ty = round(y0 + H_img + 0.10, 2); th = round(5.38 - ty, 2)

        body = f"""
  s.addShape(pres.ShapeType.rect, {{ x:{fx}, y:{fy}, w:{fw}, h:{fh}, fill:{{color:"FFFFFF"}}, line:{{color:{c_bronze}, width:0.5}} }});
  s.addImage({{ path: __dirname+"/examples/ch{ch:02d}/{img_key}.png", x:{x0}, y:{y0}, w:{W_img}, h:{H_img} }});
  s.addShape(pres.ShapeType.rect, {{ x:0.30, y:{ty}, w:9.40, h:{th:.2f}, fill:{{color:{c_slate}}} }});
  s.addText("{esc(explanation_title)}", {{ x:0.40, y:{ty+0.05:.2f}, w:9.20, h:0.24, fontSize:12, bold:true, color:C.gold, fontFace:"Georgia", valign:"middle" }});
  s.addText("{esc(explanation_zh)}", {{ x:0.40, y:{ty+0.30:.2f}, w:9.20, h:{th-0.33:.2f}, fontSize:12, color:{c_ivory}, fontFace:"Calibri", paraSpaceAfter:0, valign:"top" }});
}}"""
    else:
        # ── Side-by-side layout (tall score) ────────────────────────────────
        # Score fills left column at full content height.
        # Column widths are NOT equal — score gets its natural width (H×asp),
        # capped only to ensure text panel is at least 2.8" wide.
        y0      = 0.90          # content starts below title/subtitle
        H_img   = round(5.45 - y0, 2)          # e.g. 4.55"
        MIN_TEXT_W = 2.80       # guarantee readable text column
        W_img   = round(min(9.40 - MIN_TEXT_W - 0.28, H_img * asp), 2)  # ≤ 6.32"
        x0      = 0.30
        fx      = round(x0 - 0.08, 2); fy = round(y0 - 0.06, 2)
        fw      = round(W_img + 0.16, 2); fh = round(H_img + 0.12, 2)
        # Right panel — gets whatever space is left after the score
        xp      = round(x0 + W_img + 0.28, 2)
        wp      = round(9.70 - xp, 2)
        th_text = round(H_img - 0.42, 2)        # text area height inside panel

        body = f"""
  s.addShape(pres.ShapeType.rect, {{ x:{fx}, y:{fy}, w:{fw}, h:{fh}, fill:{{color:"FFFFFF"}}, line:{{color:{c_bronze}, width:0.5}} }});
  s.addImage({{ path: __dirname+"/examples/ch{ch:02d}/{img_key}.png", x:{x0}, y:{y0}, w:{W_img}, h:{H_img} }});
  s.addShape(pres.ShapeType.rect, {{ x:{xp}, y:{fy}, w:{wp}, h:{fh}, fill:{{color:{c_slate}}} }});
  s.addText("{esc(explanation_title)}", {{ x:{xp+0.10:.2f}, y:{y0+0.05:.2f}, w:{wp-0.20:.2f}, h:0.28, fontSize:12, bold:true, color:C.gold, fontFace:"Georgia", valign:"middle" }});
  s.addText("{esc(explanation_zh)}", {{ x:{xp+0.10:.2f}, y:{y0+0.40:.2f}, w:{wp-0.20:.2f}, h:{th_text:.2f}, fontSize:13, color:{c_ivory}, fontFace:"Calibri", paraSpaceAfter:0, lineSpacingMultiple:1.3, valign:"top" }});
}}"""

    return header + body


def update_toc(js: str, first_example_slide: int, num_added: int) -> str:
    """Increment TOC slide numbers >= first_example_slide by num_added."""
    def increment(m):
        n = int(m.group(1))
        new_n = n + num_added if n >= first_example_slide else n
        return f'[{new_n},'
    return re.sub(r'\[\s*(\d+)\s*,', increment, js)


def patch_js(ch: int, js_path: Path, examples_data: dict,
             explanations: dict, dry_run: bool = False) -> bool:
    """
    Insert example slides into chapter JS before Key Terms.
    examples_data = {ex_num: (W_px, H_px)}
    explanations  = {ex_num: {'title_zh': str, 'subtitle_en': str,
                               'explanation_title': str, 'explanation_zh': str}}
    Returns True on success.
    """
    js = js_path.read_text(encoding='utf-8')
    palette = read_c_palette(js)

    # ── Remove any previously inserted example blocks ─────────────────────
    marker = f'// ── EXAMPLE {ch}.'
    if marker in js:
        first_ex_pos = js.find(marker)
        insertion_idx_old = find_insertion_point(js)
        old_block = js[first_ex_pos:insertion_idx_old]
        old_count = old_block.count(marker)
        old_first_slide = count_slides_before(js[:first_ex_pos], len(js[:first_ex_pos])) + 1
        # Remove old blocks
        js = js[:first_ex_pos] + js[insertion_idx_old:]
        # Reverse TOC increments: decrement entries >= old_first_slide by old_count
        def decrement(m):
            n = int(m.group(1))
            return f'[{n - old_count},' if n >= old_first_slide + old_count else f'[{n},'
        js = re.sub(r'\[\s*(\d+)\s*,', decrement, js)

    insertion_idx = find_insertion_point(js)
    first_example_slide = count_slides_before(js, insertion_idx) + 1

    # Build all example slide blocks
    blocks = []
    for ex_num in sorted(examples_data.keys(), key=int):
        dims = examples_data[ex_num]
        ex_info    = explanations.get(ex_num, {})
        title_base  = ex_info.get('title_zh',          f"譜例 {ch}.{ex_num}")
        subtitle_en = ex_info.get('subtitle_en',        f"Example {ch}.{ex_num}")
        exp_title   = ex_info.get('explanation_title',  f"譜例說明  教科書")
        exp_zh      = ex_info.get('explanation_zh',     f"• （說明文字待補）")

        if isinstance(dims, list):
            # Split example: one slide per part
            n = len(dims)
            for i, part_dims in enumerate(dims):
                suffix   = chr(ord('a') + i)
                page_lbl = f"（{i+1}/{n}）"
                title_zh = title_base + page_lbl
                # Show explanation only on the last part
                part_exp_zh = exp_zh if i == n - 1 else ""
                block = build_example_slide_js(
                    ch, ex_num, palette, part_dims,
                    title_zh, subtitle_en, exp_title, part_exp_zh,
                    img_suffix=suffix
                )
                blocks.append(block)
        else:
            block = build_example_slide_js(
                ch, ex_num, palette, dims,
                title_base, subtitle_en, exp_title, exp_zh
            )
            blocks.append(block)

    insert_text = '\n'.join(blocks) + '\n\n'
    num_added = len(blocks)

    # Insert at the insertion point
    new_js = js[:insertion_idx] + insert_text + js[insertion_idx:]

    # Update TOC slide numbers
    new_js = update_toc(new_js, first_example_slide, num_added)

    if dry_run:
        print(f"  [dry-run] would insert {num_added} example slides before slide {first_example_slide}")
        print(f"  [dry-run] TOC entries >= {first_example_slide} incremented by {num_added}")
        return True

    js_path.write_text(new_js, encoding='utf-8')
    print(f"  Patched {js_path.name}: +{num_added} example slides (TOC updated)")
    return True


# ── Phase 4: Explanation Text ─────────────────────────────────────────────────

def extract_explanation_en(ch: int, ex_num: str) -> str:
    """
    Extract 1–2 English sentences describing Example ch.ex_num from the textbook.
    Returns cleaned English text, or empty string if not found.
    """
    p_start, p_end = CH_PAGES[ch]
    pdf_start = p_start + PDF_OFFSET
    pdf_end   = p_end   + PDF_OFFSET

    result = subprocess.run(
        ["pdftotext", "-layout",
         "-f", str(pdf_start), "-l", str(pdf_end), PDF, "-"],
        capture_output=True, text=True
    )
    # Clean: strip running heads, page numbers, figure captions
    lines = []
    for line in result.stdout.split('\n'):
        s = line.strip()
        if not s:
            continue
        if re.match(r'Grout10e_', s):
            continue
        if re.match(r'^[0-9]+\s+C\s*H\s*A\s*P', s):
            continue
        if re.match(r'^\d+\s*$', s):
            continue
        lines.append(s)

    full = ' '.join(lines)

    # Find sentence(s) that mention "Example ch.ex_num" inline
    pattern = rf'[^.!?]*\bExample {ch}\.{ex_num}\b[^.!?]*[.!?]'
    matches = re.findall(pattern, full)
    if matches:
        return ' '.join(matches[:2]).strip()

    return ''


def load_translations() -> dict:
    """Load pre-translated explanations from scripts/translations.json if it exists."""
    tr_path = Path(__file__).parent / "translations.json"
    if tr_path.exists():
        import json
        with open(tr_path, encoding='utf-8') as f:
            return json.load(f)
    return {}


def build_explanations(ch: int, pages: dict) -> dict:
    """
    Build explanations dict for all examples in a chapter.
    Returns {ex_num: {title_zh, subtitle_en, explanation_title, explanation_zh}}
    Priority: translations.json > auto-extracted English.
    """
    tr = load_translations()
    ch_tr = tr.get(str(ch), {})

    explanations = {}
    for ex_num in sorted(pages.keys(), key=int):
        # Use pre-translated entry if available
        if ex_num in ch_tr:
            explanations[ex_num] = ch_tr[ex_num]
            print(f"  ex{ch}.{ex_num}: [translated] {ch_tr[ex_num]['title_zh']}")
            continue

        # Fallback: auto-extract English
        en_text = extract_explanation_en(ch, ex_num)
        if en_text:
            explanation_title = f"譜例說明  教科書第 {CH_PAGES[ch][0]}–{CH_PAGES[ch][1]} 頁"
            explanation_zh = f"• {en_text[:300]}"
        else:
            explanation_title = f"譜例說明  教科書第 {CH_PAGES[ch][0]}–{CH_PAGES[ch][1]} 頁"
            explanation_zh = f"• 見教科書第 {CH_PAGES[ch][0]}–{CH_PAGES[ch][1]} 頁。"

        explanations[ex_num] = {
            'title_zh':          f"譜例 {ch}.{ex_num}",
            'subtitle_en':       f"Example {ch}.{ex_num}",
            'explanation_title': explanation_title,
            'explanation_zh':    explanation_zh,
        }
        snippet = explanation_zh[:60].replace('\n', ' ')
        print(f"  ex{ch}.{ex_num}: [auto] {snippet}…")
    return explanations


# ── Phase 5: Rebuild ──────────────────────────────────────────────────────────

def get_pptx_name(js_path: Path) -> str:
    """Extract the PPTX filename from pres.writeFile({ fileName: "..." }) in the JS."""
    js = js_path.read_text(encoding='utf-8')
    m = re.search(r'writeFile\s*\(\s*\{[^}]*fileName\s*:\s*["\']([^"\']+)["\']', js)
    if m:
        return m.group(1)
    # Fallback: derive from JS filename (e.g. ch04_song_dance.js → Ch04_Song_Dance.pptx)
    name = js_path.stem.replace('_', ' ').title().replace(' ', '_')
    return f"{name}.pptx"


def rebuild_chapter(ch: int) -> bool:
    """Run node + soffice to regenerate PPTX and PDF for a chapter."""
    js_path = find_js_file(ch)
    pptx_name = get_pptx_name(js_path)
    print(f"  Building {js_path.name} → {pptx_name}…", end=" ", flush=True)

    r1 = subprocess.run(["node", str(js_path)], capture_output=True, cwd=str(ROOT))
    if r1.returncode != 0:
        print(f"ERROR (node): {r1.stderr.decode()[:200]}")
        return False

    pptx_path = ROOT / pptx_name
    if not pptx_path.exists():
        print(f"ERROR: {pptx_name} not found after node run")
        return False

    r2 = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", str(pptx_path)],
        capture_output=True, cwd=str(ROOT)
    )
    if r2.returncode != 0:
        print(f"ERROR (soffice): {r2.stderr.decode()[:200]}")
        return False

    print("OK")
    return True


# ── Main ──────────────────────────────────────────────────────────────────────

def process_chapter(ch: int, dry_run: bool = False) -> bool:
    if ch in NO_EXAMPLES:
        print(f"Ch{ch:02d}: no examples — skip")
        return True

    print(f"\n{'='*60}")
    print(f"Ch{ch:02d}  (printed pp.{CH_PAGES[ch][0]}–{CH_PAGES[ch][1]})")
    print(f"{'='*60}")

    # Phase 1: Discover
    print("Phase 1: Discovering example pages…")
    pages = discover_examples(ch)
    if not pages:
        print(f"  WARNING: no EX AMPLE captions found for Ch{ch:02d}")
        return False
    print(f"  Found {len(pages)} example(s): {sorted(pages.keys(), key=int)}")

    # Phase 2: Extract
    print("Phase 2: Cropping score images…")
    dims = extract_all_examples(ch, pages, dry_run=dry_run)

    # Phase 4: Explanation text (before Phase 3 so we have text to embed)
    print("Phase 4: Extracting explanation text…")
    explanations = build_explanations(ch, pages)
    for num, info in sorted(explanations.items(), key=lambda x: int(x[0])):
        snippet = info['explanation_zh'][:80].replace('\n', ' ')
        print(f"  ex{ch}.{num}: {snippet}…")

    # Phase 3: Patch JS
    print("Phase 3: Patching JS…")
    try:
        js_path = find_js_file(ch)
    except FileNotFoundError as e:
        print(f"  ERROR: {e}")
        return False
    patch_js(ch, js_path, dims, explanations, dry_run=dry_run)

    # Phase 5: Rebuild
    if not dry_run:
        print("Phase 5: Rebuilding…")
        return rebuild_chapter(ch)

    return True


def main():
    parser = argparse.ArgumentParser(description="Add textbook examples to chapter slides")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--chapters', nargs='+', type=int, metavar='N',
                       help="Chapter numbers to process")
    group.add_argument('--all', action='store_true',
                       help="Process all 36 chapters with examples")
    parser.add_argument('--dry-run', action='store_true',
                        help="Show what would be done without writing files")
    args = parser.parse_args()

    chapters = list(range(1, 40)) if args.all else args.chapters
    chapters = [c for c in chapters if c not in NO_EXAMPLES]

    results = {}
    for i, ch in enumerate(chapters, 1):
        print(f"\n[{i:02d}/{len(chapters):02d}] Processing Ch{ch:02d}…")
        ok = process_chapter(ch, dry_run=args.dry_run)
        results[ch] = ok

    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    for ch, ok in sorted(results.items()):
        status = "OK" if ok else "FAILED"
        print(f"  Ch{ch:02d}: {status}")
    failed = [c for c, ok in results.items() if not ok]
    if failed:
        print(f"\nFailed chapters: {failed}")
        sys.exit(1)
    else:
        print(f"\nAll {len(results)} chapter(s) processed successfully.")


if __name__ == "__main__":
    main()
