"""
Nutrition Label Extraction Pipeline
====================================

Extracts Nutrition Facts panels and ingredient statements from marketing /
product-specification PDFs and exports them to an Excel workbook for
integration with FPED and FNDDS.

The PDFs come in two flavours:

  * Vector text  - the panel is real (selectable) text, but PDF reading order
                   usually scrambles the labels away from their values. We fix
                   this by rebuilding visual rows from word coordinates.
  * Image panels - the panel is a raster image (or vector paths drawn with a
                   broken font). These are handled with Tesseract OCR.

Because these are 2-column spec sheets, neighbouring text frequently lands on
the same visual row as a panel value, so the parser only reads a value within
a short span of its nutrient label and a coordinate fallback recovers the
large Calories figure when it prints on its own line.

Micronutrients that a label only reports as a % Daily Value (common for the
B-vitamins on breakfast cereals) are back-calculated from the FDA Daily Value
table and flagged. Every nutrient is also converted to a per-100 g basis,
which is the basis FNDDS uses.

Run:  python nutrition_extract.py
Out:  output/nutrition_extraction.xlsx   (Extraction + FNDDS_per_100g sheets)
"""

from __future__ import annotations

import io
import math
import os
import re
import shutil
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import argparse

import pymupdf
import pandas as pd
from PIL import Image, ImageOps, ImageFilter
import pytesseract


# ==============================================================
# TESSERACT LOCATION  (no PATH changes required)
# ==============================================================
# Order of preference: explicit env var -> tesseract on PATH -> the typical
# Windows install location used on the analyst's machine.
_TESS_ENV = os.environ.get("TESSERACT_CMD")
_TESS_PATH = shutil.which("tesseract")
_TESS_WIN = r"C:\Users\Conor.McGovern\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# ==============================================================
# SUPPORTED IMAGE EXTENSIONS
# ==============================================================
#: File suffixes (lowercase) treated as standalone image inputs.
IMAGE_EXTENSIONS: frozenset = frozenset(
    {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".webp", ".bmp"}
)
if _TESS_ENV:
    pytesseract.pytesseract.tesseract_cmd = _TESS_ENV
elif _TESS_PATH:
    pytesseract.pytesseract.tesseract_cmd = _TESS_PATH
elif Path(_TESS_WIN).exists():
    pytesseract.pytesseract.tesseract_cmd = _TESS_WIN


# ==============================================================
# NUTRIENT MASTER TABLE
# ==============================================================
# Each entry:
#   code        - internal column name
#   unit        - canonical reporting unit (FDA panel units are fixed)
#   dv          - FDA Daily Value (adults / children >=4 yr), or None
#   fndds_no    - FNDDS nutrient number, or None
#   fndds_name  - FNDDS nutrient description
#   label       - regex matching the nutrient's label on a panel row
#
# FDA Nutrition Facts panels use fixed units per nutrient, so the parser
# trusts the canonical unit rather than whatever unit OCR happens to read.

NUTRIENTS: List[dict] = [
    dict(code="energy_kcal",       unit="kcal", dv=None,  fndds_no=208, fndds_name="Energy (kcal)",
         label=r"calories"),
    dict(code="protein_g",         unit="g",    dv=50,    fndds_no=203, fndds_name="Protein (g)",
         label=r"protein"),
    dict(code="total_fat_g",       unit="g",    dv=78,    fndds_no=204, fndds_name="Total Fat (g)",
         label=r"total\s*fat"),
    dict(code="sat_fat_g",         unit="g",    dv=20,    fndds_no=606, fndds_name="Fatty acids, total saturated (g)",
         label=r"sat(?:urated|\.)?\s*fat"),
    dict(code="trans_fat_g",       unit="g",    dv=None,  fndds_no=605, fndds_name="Fatty acids, total trans (g)",
         label=r"trans\s*fat"),
    dict(code="mono_fat_g",        unit="g",    dv=None,  fndds_no=645, fndds_name="Fatty acids, total monounsaturated (g)",
         label=r"monoun?saturated\s*fat"),
    dict(code="poly_fat_g",        unit="g",    dv=None,  fndds_no=646, fndds_name="Fatty acids, total polyunsaturated (g)",
         label=r"polyun?saturated\s*fat"),
    dict(code="cholesterol_mg",    unit="mg",   dv=300,   fndds_no=601, fndds_name="Cholesterol (mg)",
         label=r"cholesterol"),
    dict(code="carb_g",            unit="g",    dv=275,   fndds_no=205, fndds_name="Carbohydrate (g)",
         label=r"(?:total\s*)?carb(?:ohydrates?|s)?\.?\b"),
    dict(code="fiber_g",           unit="g",    dv=28,    fndds_no=291, fndds_name="Fiber, total dietary (g)",
         label=r"(?:dietary\s*)?fib(?:er|re)"),
    dict(code="total_sugars_g",    unit="g",    dv=None,  fndds_no=269, fndds_name="Sugars, total (g)",
         label=r"sugars?"),
    dict(code="added_sugars_g",    unit="g",    dv=50,    fndds_no=None, fndds_name="Added sugars (g)",
         label=r"added\s*sugars?"),
    dict(code="sodium_mg",         unit="mg",   dv=2300,  fndds_no=307, fndds_name="Sodium (mg)",
         label=r"sodium"),
    dict(code="potassium_mg",      unit="mg",   dv=4700,  fndds_no=306, fndds_name="Potassium (mg)",
         label=r"potassium"),
    dict(code="calcium_mg",        unit="mg",   dv=1300,  fndds_no=301, fndds_name="Calcium (mg)",
         label=r"calcium"),
    dict(code="iron_mg",           unit="mg",   dv=18,    fndds_no=303, fndds_name="Iron (mg)",
         label=r"iron"),
    dict(code="magnesium_mg",      unit="mg",   dv=420,   fndds_no=304, fndds_name="Magnesium (mg)",
         label=r"magnesium"),
    dict(code="phosphorus_mg",     unit="mg",   dv=1250,  fndds_no=305, fndds_name="Phosphorus (mg)",
         label=r"phosphorus"),
    dict(code="zinc_mg",           unit="mg",   dv=11,    fndds_no=309, fndds_name="Zinc (mg)",
         label=r"zinc"),
    dict(code="vitamin_a_mcg_rae", unit="mcg",  dv=900,   fndds_no=320, fndds_name="Vitamin A, RAE (mcg_RAE)",
         label=r"vitamin\s*a"),
    dict(code="vitamin_c_mg",      unit="mg",   dv=90,    fndds_no=401, fndds_name="Vitamin C (mg)",
         label=r"vitamin\s*c"),
    dict(code="vitamin_d_mcg",     unit="mcg",  dv=20,    fndds_no=328, fndds_name="Vitamin D (D2+D3) (mcg)",
         label=r"vitamin\s*d"),
    dict(code="vitamin_e_mg",      unit="mg",   dv=15,    fndds_no=323, fndds_name="Vitamin E (alpha-tocopherol) (mg)",
         label=r"vitamin\s*e"),
    dict(code="thiamin_mg",        unit="mg",   dv=1.2,   fndds_no=404, fndds_name="Thiamin (mg)",
         label=r"thiamin"),
    dict(code="riboflavin_mg",     unit="mg",   dv=1.3,   fndds_no=405, fndds_name="Riboflavin (mg)",
         label=r"riboflavin"),
    dict(code="niacin_mg",         unit="mg",   dv=16,    fndds_no=406, fndds_name="Niacin (mg)",
         label=r"niacin"),
    dict(code="vitamin_b6_mg",     unit="mg",   dv=1.7,   fndds_no=415, fndds_name="Vitamin B-6 (mg)",
         label=r"vitamin\s*b\s*-?\s*6"),
    dict(code="folate_mcg_dfe",    unit="mcg",  dv=400,   fndds_no=435, fndds_name="Folate, DFE (mcg_DFE)",
         label=r"folate"),
    dict(code="folic_acid_mcg",    unit="mcg",  dv=None,  fndds_no=431, fndds_name="Folic acid (mcg)",
         label=r"folic\s*acid"),
    dict(code="vitamin_b12_mcg",   unit="mcg",  dv=2.4,   fndds_no=418, fndds_name="Vitamin B-12 (mcg)",
         label=r"vitamin\s*b\s*-?\s*12"),
]

NUTRIENT_BY_CODE: Dict[str, dict] = {n["code"]: n for n in NUTRIENTS}
NUTRIENT_CODES: List[str] = [n["code"] for n in NUTRIENTS]

# How far past a nutrient label a value is allowed to sit. Keeps a panel value
# attached to its own label instead of grabbing a number that bled in from a
# neighbouring column on the same visual row.
VALUE_SPAN = 34

# Matches percentage-like phrases found in *ingredient* lists, not nutrition panels
# (e.g. "contains 2% or less", "minimum of 280 calories"). Used to skip false
# label matches when marketing or ingredient text bleeds into the panel window.
_INGR_PCT_RE = re.compile(
    r"\b(?:or\s+less|no\s+more|at\s+least|minimum\s+of|not\s+more"
    r"|less\s+than|more\s+than|up\s+to)\b", re.I
)

# Nutrient keywords that signal a line is a nutrition panel row, not a serving
# size description. Used to avoid extracting gram amounts from nutrient rows
# when searching adjacent lines for a serving size. The trailing \w* matches
# plural / variant spellings on these labels (Sugars, Carbohydrates, Fibres).
_NUTRIENT_KW = re.compile(
    r"\b(?:cholesterol|sodium|potassium|calcium|iron|vitamin|fat|carb|fib(?:er|re)|sugar|protein)\w*",
    re.I,
)

# Ingredient / grain-credit line keywords. Lines containing these are not
# serving size sources: e.g. "GRAIN CREDIT: 2 oz" or "GRAMS OF FLOUR: 33.8g".
_INGR_LINE_KW = re.compile(
    r"\b(?:grain\s+credit|grams?\s+of\s+flour|creditable|meal\s+equivalent)\b", re.I,
)


# ==============================================================
# TEXT NORMALISATION
# ==============================================================

def normalize_text(s: str) -> str:
    """Collapse odd encodings / whitespace and drop control characters."""
    s = unicodedata.normalize("NFKC", s)
    s = "".join(ch for ch in s if ord(ch) >= 32 or ch in "\r\n\t")
    s = s.replace("**", "")
    s = re.sub(r"[^\S\r\n]+", " ", s)
    return s


def clean_row(s: str) -> str:
    """Repair the most common OCR / subset-font confusions for a panel row."""
    s = normalize_text(s)
    # Micro sign -> u (so 'ug' == mcg). NFKC folds U+00B5 to Greek mu U+03BC,
    # so both forms must be handled.
    s = s.replace("µ", "u").replace("μ", "u")
    # A letter O or Q standing in for a zero next to a unit/digit. Small panel
    # text rounds a zero into either letter, so "Q.2mg" is a misread "0.2mg";
    # the third rule also repairs the leading zero of a decimal ("O.2"/"Q.2"),
    # which the others miss because a '.' - not a digit - follows the letter.
    s = re.sub(r"(?i)\b[OQ](?=(?:mcg|mg|ug|g)\b)", "0", s)
    s = re.sub(r"(?i)(?<=\d)[OQ](?=\d)", "0", s)
    s = re.sub(r"(?i)\b[OQ](?=\.?\d)", "0", s)
    # Trailing O/Q (one or more) after a digit: "27O"→"270", "3QQ"→"300".
    s = re.sub(r"(?i)(?<=\d)[OQ]+(?=\s|[^a-zA-Z0-9]|$)",
               lambda m: "0" * len(m.group(0)), s)
    # "X0z" → "Xoz": letter O misread as digit 0 before 'z' in "oz" unit.
    # Covers "2.90z" → "2.9oz" and "30z" → "3oz".
    s = re.sub(r"(?i)(\d)0z\b", r"\g<1>oz", s)
    return s


# ==============================================================
# VECTOR-TEXT EXTRACTION  (coordinate-based row reconstruction)
# ==============================================================

def cluster_rows(words: List[tuple]) -> List[List[tuple]]:
    """Group words into visual rows by their vertical centre."""
    if not words:
        return []
    heights = sorted(w[3] - w[1] for w in words if w[3] > w[1])
    median_h = heights[len(heights) // 2] if heights else 8.0
    ytol = max(2.5, 0.45 * median_h)

    words = sorted(words, key=lambda w: ((w[1] + w[3]) / 2.0, w[0]))
    rows: List[List[tuple]] = []
    current: List[tuple] = []
    centre: Optional[float] = None
    for w in words:
        ymid = (w[1] + w[3]) / 2.0
        if centre is None or abs(ymid - centre) <= ytol:
            current.append(w)
            centre = ymid if centre is None else (0.6 * centre + 0.4 * ymid)
        else:
            rows.append(current)
            current = [w]
            centre = ymid
    if current:
        rows.append(current)
    return rows


def reconstruct_rows(words: List[tuple]) -> List[str]:
    """
    Rebuild visual rows from word coordinates.

    PyMuPDF's plain text extraction follows the PDF content stream, which on
    these spec sheets interleaves a product-info column with the Nutrition
    Facts column. Grouping words by their vertical centre and sorting by x
    restores each printed row, keeping every label next to its value.
    """
    out: List[str] = []
    for row in cluster_rows(words):
        row = sorted(row, key=lambda w: w[0])
        text = " ".join(w[4] for w in row).strip()
        if text:
            out.append(clean_row(text))
    return out


def extract_text(pdf_path: Path) -> Tuple[List[str], str, List[List[tuple]]]:
    """Return (reconstructed rows, raw page text, per-page word lists)."""
    rows: List[str] = []
    raw_parts: List[str] = []
    pages_words: List[List[tuple]] = []
    with pymupdf.open(str(pdf_path)) as doc:
        for page in doc:
            words = [tuple(w[:5]) for w in page.get_text("words")]
            pages_words.append(words)
            rows.extend(reconstruct_rows(words))
            raw_parts.append(page.get_text("text"))
    return rows, normalize_text("\n".join(raw_parts)), pages_words


# ==============================================================
# OCR  (image / broken-font panels)
# ==============================================================

def _to_rgb_on_white(img: Image.Image) -> Image.Image:
    """Flatten any transparency onto a white background."""
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
        img = Image.alpha_composite(bg, img)
    return img.convert("RGB")


def _ocr(img: Image.Image, psm: int, threshold: Optional[int] = None,
         scale: int = 2) -> str:
    """Pre-process an image and run Tesseract."""
    g = ImageOps.grayscale(img)
    g = ImageOps.autocontrast(g)
    if threshold is not None:
        g = g.point(lambda p: 255 if p > threshold else 0)
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.6, percent=150))
    if scale > 1:
        g = g.resize((g.width * scale, g.height * scale), Image.LANCZOS)
    return pytesseract.image_to_string(
        g, config=f"--oem 1 --psm {psm} -c preserve_interword_spaces=1"
    )


def ocr_collect(pdf_path: Path) -> str:
    """
    OCR every plausible source of a Nutrition Facts panel and concatenate the
    results. Embedded panel images are emitted first because they are the
    cleanest source; full-page renders follow as a fallback. Several
    preprocessing variants are produced on purpose - the parser keeps the
    first usable (unit-bearing) value for each nutrient, so a clean read in
    one variant repairs a garbled read in another.
    """
    image_blocks: List[str] = []
    page_blocks: List[str] = []
    with pymupdf.open(str(pdf_path)) as doc:
        for page in doc:
            seen = set()
            for img in page.get_images(full=True):
                xref, _, w, h = img[0], img[1], img[2], img[3]
                if xref in seen or min(w, h) < 150 or max(w, h) < 300:
                    continue
                seen.add(xref)
                try:
                    raw = doc.extract_image(xref)
                    panel = _to_rgb_on_white(Image.open(io.BytesIO(raw["image"])))
                except Exception:
                    continue
                image_blocks.append(_ocr(panel, psm=6, scale=3))
                image_blocks.append(_ocr(panel, psm=6, threshold=185, scale=3))

            pix = page.get_pixmap(dpi=400, alpha=False)
            page_img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            page_blocks.append(_ocr(page_img, psm=3, scale=1))
            page_blocks.append(_ocr(page_img, psm=6, threshold=190, scale=1))

    blocks = image_blocks + page_blocks
    return "\n".join(clean_row(line) for b in blocks for line in b.splitlines())


def ocr_from_image_file(img_path: Path) -> str:
    """
    OCR a standalone image file (JPG, PNG, TIFF, WebP, BMP, …).

    Several preprocessing variants are produced for the same reasons as in
    ``ocr_collect``: clean reads in one variant repair garbled reads in
    another, and the parser keeps the first usable value per nutrient.

    Scale is chosen adaptively so that phone photos (already large) are not
    inflated further, while small/thumbnail images are upscaled to give
    Tesseract enough resolution to read fine nutrition-label text:

      * max side ≥ 1 500 px  →  scale = 1  (high-res photo, no upscaling)
      * 600 – 1 499 px       →  scale = 2
      * < 600 px             →  scale = 3  (small/embedded thumbnail)
    """
    try:
        img = _to_rgb_on_white(Image.open(str(img_path)))
    except Exception as exc:
        raise ValueError(f"Cannot open image {img_path.name}: {exc}") from exc

    max_dim = max(img.width, img.height)
    scale = 1 if max_dim >= 1500 else (2 if max_dim >= 600 else 3)

    blocks: List[str] = []
    # Block-layout mode — best for a cropped, upright label
    blocks.append(_ocr(img, psm=6, scale=scale))
    blocks.append(_ocr(img, psm=6, threshold=185, scale=scale))
    # Full-auto page segmentation — handles tilted / cluttered photos
    blocks.append(_ocr(img, psm=3, scale=scale))
    blocks.append(_ocr(img, psm=3, threshold=190, scale=scale))

    return "\n".join(clean_row(line) for b in blocks for line in b.splitlines())


# ==============================================================
# PANEL PARSING
# ==============================================================

_AMOUNT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:mcg|mg|ug|g)\b", re.I)
_PCT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%")
_INT_RE = re.compile(r"(?<![\d.])(\d{1,4})(?![\d.])")
_PURE_INT_RE = re.compile(r"^\d{2,4}$")


_PANEL_START = re.compile(r"^\s*nutrition\w*\s*(?:facts?|information\s*:)", re.I)


def detect_three_column(text: str) -> bool:
    """True for panels that print a dedicated per-100 g column."""
    return (
        re.search(r"100\s*g\s*\(\s*100\s*g\s*\)", text, re.I) is not None
        or re.search(r"per\s*100\s*g(?:rams?)?\b", text, re.I) is not None
    )


def panel_window(lines: List[str]) -> List[str]:
    """
    Drop everything before the 'Nutrition Facts' header.

    This excludes marketing copy (e.g. "...a good source of fiber, 13g whole
    grains...") that would otherwise be mistaken for a panel value. No end
    anchor is used - the parser keeps the first usable value per nutrient, so
    the footnote/ingredient text after the panel is harmless.
    """
    for i, line in enumerate(lines):
        if _PANEL_START.search(line):
            return lines[i:]
    return lines


def _amounts(text: str) -> List[float]:
    return [float(m.group(1)) for m in _AMOUNT_RE.finditer(text)]


def _first_pct(text: str) -> Optional[float]:
    m = _PCT_RE.search(text)
    return float(m.group(1)) if m else None


def parse_panel(lines: List[str], three_col: bool) -> Dict[str, dict]:
    """
    Parse reconstructed/OCR rows into per-nutrient readings.

    Returns {code: {"amount": float|None, "dv": float|None,
                     "per100g": float|None, "lt": bool}}. Amounts are taken in
    the nutrient's canonical unit (panel units are fixed by the FDA).
    """
    result: Dict[str, dict] = {}

    for spec in NUTRIENTS:
        code = spec["code"]
        label = re.compile(r"\b(?:" + spec["label"] + r")", re.I)

        for i, line in enumerate(lines):
            low = line.lower()

            # ---- added sugars: value can sit just before or after label ----
            if code == "added_sugars_g":
                m = re.search(r"added\s*sugars?", low) or re.search(r"incl\.?\s*added", low)
                if not m:
                    continue
                seg_start = max(0, m.start() - 28)
                seg = line[seg_start: m.end() + 22]
                kw_pos = m.start() - seg_start  # keyword position within seg

                # Prefer the amount that is closest to the "Added Sugars" keyword:
                # first after it (standard "Added Sugars 9g"), otherwise the last
                # before it ("Includes 13g Added Sugars"). This avoids picking up
                # amounts from a preceding nutrient on the same row.
                all_amt_m = list(_AMOUNT_RE.finditer(seg))
                post_amts = [float(x.group(1)) for x in all_amt_m if x.start() >= kw_pos]
                pre_amts  = [float(x.group(1)) for x in all_amt_m if x.start() <  kw_pos]
                if post_amts:
                    amts = [post_amts[0]]
                elif pre_amts:
                    amts = [pre_amts[-1]]
                else:
                    amts = []

                # Similarly, prefer the %DV that follows the keyword over one
                # that belongs to a preceding nutrient on the same row.
                all_pct_m = list(re.finditer(r"(\d+(?:\.\d+)?)\s*%", seg))
                post_pcts = [float(x.group(1)) for x in all_pct_m if x.start() >= kw_pos]
                pre_pcts  = [float(x.group(1)) for x in all_pct_m if x.start() <  kw_pos]
                pct = post_pcts[0] if post_pcts else (pre_pcts[-1] if pre_pcts else None)

                # Tabular format fallback: bare number before %DV (no unit suffix)
                if not amts and pct is not None:
                    bare = re.search(r"(\d+(?:\.\d+)?)\s+\d+\.?\d*\s*%", seg[kw_pos:])
                    if bare:
                        amts = [float(bare.group(1))]
                if not amts and pct is None:
                    continue
                result[code] = dict(
                    amount=amts[0] if amts else None,
                    dv=pct,
                    per100g=amts[1] if (three_col and len(amts) > 1) else None,
                    lt=bool(re.search(r"<\s*\d", seg)),
                )
                break

            # ---- total sugars: skip the "added/includes" line ----
            if code == "total_sugars_g":
                m = re.search(r"sugars?", low)
                if not m:
                    continue
                before = low[max(0, m.start() - 16):m.start()]
                if "added" in before or "incl" in before:
                    continue
            elif code == "total_fat_g":
                # JTM-style spec sheets print "Fat (g) 5 5" without the "Total"
                # prefix, so match a bare "Fat" too — but skip the saturated /
                # trans / mono / poly sub-rows that also contain "fat".
                m = re.search(r"\bfat\b", low)
                if not m:
                    continue
                before = low[max(0, m.start() - 20):m.start()]
                if re.search(r"satur|trans|mono|poly", before):
                    continue
            else:
                m = label.search(low)
                if not m:
                    continue

            tail = line[m.end(): m.end() + VALUE_SPAN]

            if code == "energy_kcal":
                # "Calories per gram - Fat 9 ..." is a footnote, not a value.
                # Marketing text like "minimum of 280 calories" is also skipped.
                if re.match(r"\s*per\b", tail, re.I) or _INGR_PCT_RE.search(tail):
                    continue
                # Skip the standard "2,000 calories a day" daily-value footnote:
                # the label regex can match the word "calories" in that sentence.
                before = low[max(0, m.start() - 12): m.start()]
                if re.search(r"\b2[,.]000\s*$", before, re.I):
                    continue
                # Normalize OCR space-split digit groups: "1 60" → "160".
                tail = re.sub(r"\b(\d)\s+(\d{2,3})\b",
                              lambda m: m.group(1) + m.group(2), tail)
                # Require ≥2 digits to avoid matching single-digit zeros embedded
                # in unit strings on the same row (e.g. "0mcg", "0%").
                ints = [int(x) for x in re.findall(r"(?<![\d.])(\d{2,4})(?![\d.])", tail)]
                plausible = [v for v in ints if 10 <= v <= 2000]
                # Some panels print "Calories\nper serving 200" on two rows.
                # Limit to first 30 chars, normalize space-split digits, and only
                # scan up to the first nutrient keyword (e.g. "Vitamin") so that
                # a calorie value before "Vitamin D 0%" is found but a Calcium
                # value later on the same line is not.
                # Also exclude bare percentages (e.g. "51% WHOLE GRAIN").
                if not plausible and i + 1 < len(lines):
                    nxt30 = lines[i + 1][:30]
                    nxt30 = re.sub(r"\b(\d)\s+(\d{2,3})\b",
                                   lambda m: m.group(1) + m.group(2), nxt30)
                    kw_m = _NUTRIENT_KW.search(nxt30)
                    nxt_scan = nxt30[:kw_m.start()] if kw_m else nxt30
                    ints = [int(x) for x in re.findall(
                        r"(?<![\d.])(\d{2,4})(?![\d.%])", nxt_scan)]
                    plausible = [v for v in ints if 10 <= v <= 2000]
                if not plausible:
                    continue
                result[code] = dict(
                    amount=float(plausible[0]),
                    dv=_first_pct(tail),
                    per100g=float(plausible[1]) if (three_col and len(plausible) > 1) else None,
                    lt=False,
                )
                break

            amts = _amounts(tail)
            dv = _first_pct(tail)

            # Skip ingredient-list percentage phrases that bleed into the panel
            # window (e.g. "contains 2% or less", "no more than 15 fat grams").
            if _INGR_PCT_RE.search(tail):
                amts, dv = [], None

            # Tabular format fallback: some spec sheets print nutrients without a
            # unit suffix ("Total Fat: 13 17%" instead of "Total Fat 13g 17%").
            # When %DV is present, the number immediately before it is the amount.
            if not amts and dv is not None:
                bare = re.search(r"(\d+(?:\.\d+)?)\s+\d+\.?\d*\s*%", tail)
                if bare:
                    val_str = bare.group(1)
                    # g→9 OCR misread: "3g 10%" can OCR as "39 10%". Any 2+-digit
                    # integer ending in 9 is suspect; strip the trailing '9' only
                    # when the stripped value round-trips to the printed %DV
                    # (within 25% slack for label rounding).
                    daily = spec.get("dv")
                    stripped = None
                    if re.fullmatch(r"\d+9", val_str) and len(val_str) >= 2:
                        cand = float(val_str[:-1])
                        if daily and dv > 0:
                            implied = cand / daily * 100.0
                            if abs(implied - dv) <= max(2.0, dv * 0.25):
                                stripped = cand
                    if stripped is not None:
                        amts = [stripped]
                    else:
                        amts = [float(val_str)]

            # Same tabular format, nutrients with no %DV ("Protein: 17 -").
            if not amts and dv is None:
                bare = re.search(r"(\d+(?:\.\d+)?)\s*[-–]\s*$", tail.strip())
                if bare:
                    val_str = bare.group(1)
                    if re.fullmatch(r"\d{2,}9", val_str):
                        amts = [float(val_str[:-1])]
                    else:
                        amts = [float(val_str)]

            # Some PDFs split the amount onto the next line when ingredients text
            # occupies the same row ("Total Carbohydrate 8% INGREDIENTS...\n22g ...").
            if not amts and dv is not None and i + 1 < len(lines):
                next_amts = _amounts(lines[i + 1][:15])
                if next_amts:
                    amts = next_amts

            # Non-standard format: unit in parens before the number
            # ("Total Fat (g) 6", or two-column "Protein (g) 13 14" on spec
            # sheets with a dedicated per-100 g column).
            if not amts:
                m2 = re.search(
                    r"\(\s*(?:mcg|mg|ug|g)\s*\)\s*"
                    r"(\d+(?:\.\d+)?)(?:\s+(\d+(?:\.\d+)?))?\b",
                    tail,
                )
                if m2:
                    amts = [float(m2.group(1))]
                    if m2.group(2):
                        amts.append(float(m2.group(2)))

            if not amts and dv is None:
                continue
            result[code] = dict(
                amount=amts[0] if amts else None,
                dv=dv,
                per100g=amts[1] if (three_col and len(amts) > 1) else None,
                lt=bool(re.search(r"<\s*\d", tail)),
            )
            break

    # Fallback: some OCR panels garble the "Calories per serving" label (e.g.
    # "crow perservies—300"). Scan all lines for "per serv" immediately before
    # a plausible 3-4 digit calorie value.
    if "energy_kcal" not in result:
        for ln in lines:
            mp = re.search(r"per\s*serv\w*[-—\s]{0,3}(\d{3,4})\b", ln, re.I)
            if mp:
                val = int(mp.group(1))
                if 10 <= val <= 2000:
                    result["energy_kcal"] = dict(
                        amount=float(val), dv=None, per100g=None, lt=False)
                    break

    # Stronger fallback: OCR sometimes drops the "Calories" word entirely or
    # garbles it beyond recognition ("Ca lo re SS"), but the surrounding
    # "Amount Per Serving" header survives. When the label-based scan finds
    # nothing, anchor on that header and take the next plausible number on
    # the same row or one of the next two rows. Digit groups split by OCR
    # spaces ("9 0" → "90") are joined first.
    if "energy_kcal" not in result:
        def _digit_glue(s: str) -> str:
            # Collapse "9 0" → "90", "2 50" → "250". Run twice for "1 2 0".
            for _ in range(2):
                s = re.sub(r"(?<![\d.])(\d)\s+(\d)(?![\d.])", r"\1\2", s)
            return s

        anchor = re.compile(r"amount\s*per\s*serving", re.I)
        for i, ln in enumerate(lines):
            if not anchor.search(ln):
                continue
            for j in (i, i + 1, i + 2):
                if j >= len(lines):
                    break
                glued = _digit_glue(lines[j])
                # Strip the anchor itself so we don't pull a digit from words
                # like "Serving" or year numbers on the same line.
                glued = anchor.sub(" ", glued)
                for m in re.finditer(r"(?<![\d.])(\d{2,4})(?![\d.%])", glued):
                    val = int(m.group(1))
                    if 30 <= val <= 1500:
                        result["energy_kcal"] = dict(
                            amount=float(val), dv=None, per100g=None, lt=False)
                        break
                if "energy_kcal" in result:
                    break
            if "energy_kcal" in result:
                break

    # Atwater rescue: OCR often produces the calorie line twice with different
    # digit reads ("Calories 200" in one pass, "Calories 900" in another) and
    # the parser snaps to whichever came first. When the printed macros imply
    # a very different calorie value, scan every line for "Calories <number>"
    # and pick the candidate that's closest to 9*fat + 4*carb + 4*protein. The
    # same check rescues straight digit confusions ("2"→"9") that the existing
    # power-of-ten rescaler doesn't cover.
    fat = result.get("total_fat_g", {}).get("amount")
    carb = result.get("carb_g", {}).get("amount")
    prot = result.get("protein_g", {}).get("amount")
    cur = result.get("energy_kcal", {}).get("amount") if "energy_kcal" in result else None
    if cur is not None and None not in (fat, carb, prot):
        expected = 9 * fat + 4 * carb + 4 * prot
        if abs(expected - cur) > 0.5 * max(cur, expected) + 25:
            cands: List[int] = []
            for ln in lines:
                for mc in re.finditer(
                    r"calor(?:ie|y|ies)\w*\s*(?:per\s*serv\w*\s*)?(\d{2,4})",
                    ln,
                    re.I,
                ):
                    v = int(mc.group(1))
                    # Skip the standard "2,000 calorie diet" footnote.
                    if 30 <= v <= 1500 and v != 2000:
                        cands.append(v)
            if cands:
                best = min(cands, key=lambda v: abs(v - expected))
                if abs(best - expected) < abs(cur - expected):
                    result["energy_kcal"]["amount"] = float(best)

    return result


def calories_by_coordinates(pages_words: List[List[tuple]]) -> Optional[float]:
    """
    Recover the Calories figure when it prints on its own line.

    On several panels the large Calories number sits on a slightly different
    baseline than the word "Calories", so row reconstruction separates them.
    Here we find the word "Calories" and the nearest standalone 2-4 digit
    number on roughly the same line.
    """
    for words in pages_words:
        anchors = [w for w in words if w[4].strip().lower() in ("calories", "calorie")]
        for a in anchors:
            ay = (a[1] + a[3]) / 2.0
            best = None
            for w in words:
                if not _PURE_INT_RE.match(w[4].strip()):
                    continue
                value = int(w[4].strip())
                if not (10 <= value <= 1500):
                    continue
                if abs((w[1] + w[3]) / 2.0 - ay) > 18:
                    continue
                if w[0] < a[0] - 15:
                    continue
                height = w[3] - w[1]
                if best is None or height > best[0]:
                    best = (height, value)
            if best:
                return float(best[1])
    return None


def panel_score(panel: Dict[str, dict]) -> int:
    """Count nutrients with a usable amount or %DV reading."""
    return sum(1 for v in panel.values()
               if v.get("amount") is not None or v.get("dv") is not None)


# ==============================================================
# SERVING SIZE & INGREDIENTS
# ==============================================================

def parse_serving_size(lines: List[str], raw: str) -> Tuple[Optional[float], Optional[str]]:
    """
    Find the serving size and its gram weight.

    Checks a ±2-line window around each "Serving Size" label because some PDFs
    print the gram weight on its own row immediately before or after the label.
    Also handles split labels like "Serving\\nsize" by matching lines that start
    with "Serving" followed by a digit.

    Within the window, all candidate lines are first scanned for the most
    reliable form (parenthesised gram weight) before falling back to bare grams
    or ounces. Otherwise a stray "1g" on a "Polyunsaturated 1g" row could be
    mistaken for the serving size when the real "(27g)" sits two lines away.
    """
    desc: Optional[str] = None

    # Locate candidate line indices: explicit "Serving Size" label, or a split
    # label where the line starts "Serving <number>" (e.g. "Serving 1 bag (68g)").
    sv_idxs: List[int] = []
    for idx, ln in enumerate(lines):
        if re.search(r"serving\s*size", ln, re.I):
            sv_idxs.append(idx)
        elif re.match(r"\s*serving\s+\d", ln, re.I):
            sv_idxs.append(idx)

    # Build a window of nearby lines sorted by distance from the serving label.
    # Lines before the label are preferred over equally-distant lines after it.
    window: List[Tuple[int, int, int]] = []
    seen: set = set()
    for idx in sv_idxs:
        for j in range(max(0, idx - 2), min(len(lines), idx + 3)):
            if j not in seen:
                seen.add(j)
                window.append((abs(j - idx), j - idx, j))
    window.sort()

    cands: List[str] = []
    for _, _, j in window:
        line = lines[j]
        # Skip ingredient/grain-credit lines (e.g. "GRAIN CREDIT: 2 oz").
        if _INGR_LINE_KW.search(line):
            continue
        m = re.search(r"serving\s*size[:\s]*([^\n]*)", line, re.I)
        cand = m.group(1).strip() if m else line.strip()
        if cand and desc is None:
            desc = cand
        cands.append(cand)

    # Pass 1: parenthesised gram weight — the most reliable form, e.g. "(57g)".
    # Skip 100g (three-column-format column header, not serving size).
    for cand in cands:
        for gm in re.finditer(r"\(\s*(\d+(?:\.\d+)?)\s*g\s*\)", cand, re.I):
            val = float(gm.group(1))
            if val != 100.0:
                return val, cand

    # Pass 1b: g→9 OCR misread: "(66g)" can OCR as "(669)".
    for cand in cands:
        for gm in re.finditer(r"\(\s*(\d{2,3})\s*9\)", cand, re.I):
            val = float(gm.group(1))
            if val != 100.0 and 20.0 <= val <= 400.0:
                return val, cand

    # Pass 1c: tabular "Serving Size (g) 89.90" / "Serving Size (oz.) 3.17"
    # format used by some spec sheets where the unit sits inside the parens
    # before the number, not after. The "Serving Size" prefix has already
    # been stripped from the cand by the caller, so we match the parenthesised
    # unit at the start of the line. Sub-pass on (g) first, then (oz.), so a
    # nearby oz row doesn't win when a more precise gram row is also present.
    for cand in cands:
        m = re.match(r"\s*\(\s*g\s*\)\s*(\d+(?:\.\d+)?)\b", cand, re.I)
        if m:
            val = float(m.group(1))
            if val != 100.0:
                return val, cand
    for cand in cands:
        m = re.match(r"\s*\(\s*oz\.?\s*\)\s*(\d+(?:\.\d+)?)\b", cand, re.I)
        if m:
            return round(float(m.group(1)) * 28.3495, 2), cand

    # Pass 2: bare gram weight, only on lines with no nutrient keywords, to
    # avoid pulling the gram from "Includes 13g Added Sugars" on a mixed row.
    # Also skip 100g (per-100g column header).
    for cand in cands:
        if _NUTRIENT_KW.search(cand):
            continue
        g = re.search(r"(\d+(?:\.\d+)?)\s*g\b", cand, re.I)
        if g:
            val = float(g.group(1))
            if val != 100.0:
                return val, cand

    # Pass 3: ounces (converted to grams).
    for cand in cands:
        oz = re.search(r"(\d+(?:\.\d+)?)\s*oz\b", cand, re.I)
        if oz:
            return round(float(oz.group(1)) * 28.3495, 2), cand

    # Fallback: raw page text (catches serving sizes in PDF stream text).
    for line in raw.splitlines():
        if re.search(r"serving\s*size", line, re.I):
            m = re.search(r"serving\s*size[:\s]*([^\n]*)", line, re.I)
            cand = m.group(1).strip() if m else line.strip()
            if cand and desc is None:
                desc = cand
            g = re.search(r"\(\s*(\d+(?:\.\d+)?)\s*g\s*\)", cand, re.I)
            if g:
                return float(g.group(1)), cand
            if not _NUTRIENT_KW.search(cand):
                g = re.search(r"(\d+(?:\.\d+)?)\s*g\b", cand, re.I)
                if g:
                    return float(g.group(1)), cand
            oz = re.search(r"(\d+(?:\.\d+)?)\s*oz\b", cand, re.I)
            if oz:
                return round(float(oz.group(1)) * 28.3495, 2), cand

    # Fallback: "Unit Size" / "Net Wt" / "Unit Wt" field in product spec header.
    # Many school-food spec sheets list the individual item weight separately from
    # the Nutrition Facts panel (e.g. "Unit Size: 2.00 oz"). Search both the
    # reconstructed rows and the raw page text.
    # Also check the line immediately before/after the label, because some spec
    # sheets print "Unit Size" and the oz value on consecutive lines.
    for source in (lines, raw.splitlines()):
        src_list = list(source)
        for i_src, line in enumerate(src_list):
            if re.search(r"\b(?:unit\s*(?:size|wt\.?)|net\s*wt\.?)\b", line, re.I):
                for adj in range(max(0, i_src - 1), min(len(src_list), i_src + 2)):
                    oz = re.search(r"(\d+(?:\.\d+)?)\s*oz\b", src_list[adj], re.I)
                    if oz:
                        return round(float(oz.group(1)) * 28.3495, 2), desc or line.strip()

    # Fallback: scan the first lines of the panel window for oz amounts and
    # parenthesised gram weights. Handles label-stripped OCR panels where
    # "Serving Size" text is absent but the serving description still appears
    # near the panel header, e.g. "1waffle (2.9oz)" or "(91g)" / "(919)" with
    # the common g→9 OCR misread. Parenthesised grams are checked before the
    # ingredient-line skip so that lines containing both a serving weight and
    # a GRAIN CREDIT / FLOUR note still yield the correct serving size.
    panel_idx = next((i for i, ln in enumerate(lines) if _PANEL_START.search(ln)), None)
    if panel_idx is not None:
        for j in range(panel_idx, min(len(lines), panel_idx + 12)):
            ln = lines[j]
            if _NUTRIENT_KW.search(ln):
                continue
            # Normal parenthesised gram weight
            for gm in re.finditer(r"\(\s*(\d+(?:\.\d+)?)\s*g\s*\)", ln, re.I):
                val = float(gm.group(1))
                if val != 100.0 and 20.0 <= val <= 400.0:
                    return val, desc or ln.strip()
            # g→9 OCR: "(91g)" can appear as "(919)"
            for gm in re.finditer(r"\(\s*(\d{2,3})\s*9\)", ln, re.I):
                val = float(gm.group(1))
                if val != 100.0 and 20.0 <= val <= 400.0:
                    return val, desc or ln.strip()
            if _INGR_LINE_KW.search(ln):
                continue
            oz = re.search(r"(\d+(?:\.\d+)?)\s*oz\b", ln, re.I)
            if oz:
                return round(float(oz.group(1)) * 28.3495, 2), desc or ln.strip()

    return None, desc


_INGR_STOP = re.compile(
    r"\bcontains\s*:|\bcontains\s+(?:wheat|milk|soy|egg|peanut|tree|almond"
    r"|sesame|fish|shellfish|coconut|cashew|walnut|pecan)|\ballergen"
    r"|\bbioengineered|\bnutrition\s*facts|\bcn\s*statement|\btotal\s*creditable"
    r"|\bpreparation\b|\bshelf\s*life|\bdate\s*code|\bheating\s*instruction",
    re.I,
)


def parse_ingredients(raw: str) -> Optional[str]:
    """
    Pull the ingredient statement out of the readable page text.

    Every "ingredient(s)" mention is considered (these spec sheets also say
    things like "verify ingredients and allergens"); the candidate that most
    looks like a real comma-separated list wins.
    """
    best: Optional[str] = None
    best_commas = -1
    for m in re.finditer(r"ingredients?\b", raw, re.I):
        seg = raw[m.end(): m.end() + 2200]
        seg = re.sub(r"^\s*(?:&?\s*allergens?\b)?\s*(?:statement\b)?\s*[:.\-]?\s*",
                     "", seg, flags=re.I)
        stop = _INGR_STOP.search(seg)
        if stop:
            seg = seg[:stop.start()]
        seg = re.sub(r"\s*\n\s*", " ", seg).strip(" :.-")
        seg = re.sub(r"\s{2,}", " ", seg)
        commas = seg.count(",")
        # Reject broken-font garbage (scrambled subset fonts decode to
        # non-ASCII letters); real ingredient lists are plain text.
        weird = sum(1 for c in seg if ord(c) > 127 and c.isalpha())
        if weird > 8:
            continue
        if commas >= 3 and len(seg) >= 40:
            if commas > best_commas or (commas == best_commas
                                        and (best is None or len(seg) < len(best))):
                best, best_commas = seg[:2500], commas
    return best


# ==============================================================
# DERIVED VALUES
# ==============================================================

def backfill_from_dv(amount: Optional[float], dv: Optional[float],
                     code: str) -> Tuple[Optional[float], bool]:
    """Back-calculate an amount from %DV when no amount was printed."""
    if amount is not None or dv is None:
        return amount, False
    daily = NUTRIENT_BY_CODE[code]["dv"]
    if not daily:
        return None, False
    return round(daily * dv / 100.0, 4), True


def reconcile_amount_with_dv(amount: Optional[float], dv: Optional[float],
                             code: str) -> Tuple[Optional[float], bool]:
    """
    Repair a printed amount that disagrees with the printed %DV by a power
    of ten.

    Small panel text loses the decimal point under OCR ("0.2 mg" reads as
    "2 mg"), which inflates the amount tenfold. When the amount is far larger
    than the printed %DV implies and dividing it by an exact power of ten
    restores agreement, a dropped decimal is the cause and the amount is
    rescaled to the %DV. The %DV itself is only used to pick the power of ten -
    the rescaled amount keeps the label's own precision.
    """
    if amount is None or dv is None or amount <= 0 or dv <= 0:
        return amount, False
    daily = NUTRIENT_BY_CODE[code]["dv"]
    if not daily:
        return amount, False
    expected = daily * dv / 100.0
    ratio = amount / expected
    if ratio <= 3.0:                       # already consistent with the %DV
        return amount, False
    power = 10.0 ** round(math.log10(ratio))
    if power <= 1.0:                       # off, but not by a clean 10x
        return amount, False
    corrected = amount / power
    # Accept only when the rescaled amount round-trips to the printed %DV.
    # A genuine dropped decimal restores that consistency exactly; a spurious
    # extra digit ("1g" read as "19g") does not, and is left untouched.
    if abs(corrected / daily * 100.0 - dv) <= 1.0:
        return round(corrected, 4), True
    return amount, False


def per_100g(amount: Optional[float], serving_g: Optional[float]) -> Optional[float]:
    if amount is None or not serving_g:
        return None
    return round(amount * 100.0 / serving_g, 3)


# ==============================================================
# VALIDATION
# ==============================================================

def validate(row: dict, serving_g: Optional[float], method: str,
             estimated: List[str], rescaled: List[str]) -> str:
    notes: List[str] = []

    if row.get("energy_kcal") is None:
        notes.append("calories not found")
    if serving_g is None:
        notes.append("serving size (g) not found - per-100g not computed")

    fat, carb, prot = row.get("total_fat_g"), row.get("carb_g"), row.get("protein_g")
    kcal = row.get("energy_kcal")
    if None not in (fat, carb, prot, kcal):
        atwater = 9 * fat + 4 * carb + 4 * prot
        if abs(atwater - kcal) > 0.25 * kcal + 25:
            notes.append(f"Atwater check: label {kcal:g} kcal vs macros "
                         f"{atwater:g} kcal")

    add, tot = row.get("added_sugars_g"), row.get("total_sugars_g")
    if add is not None and tot is not None and add > tot + 0.5:
        notes.append("added sugars exceed total sugars")

    for code in ("total_fat_g", "carb_g", "fiber_g", "protein_g",
                 "sat_fat_g", "total_sugars_g", "added_sugars_g"):
        v = row.get(f"{code}_per100g")
        if v is not None and v > 100:
            notes.append(f"{code} per-100g implausible ({v:g})")

    if rescaled:
        nice = ", ".join(NUTRIENT_BY_CODE[c]["fndds_name"].split(" (")[0]
                         for c in rescaled)
        notes.append(f"amount rescaled to printed %DV (suspected OCR "
                     f"dropped decimal): {nice}")

    if estimated:
        nice = ", ".join(NUTRIENT_BY_CODE[c]["fndds_name"].split(" (")[0]
                         for c in estimated)
        notes.append(f"estimated from %DV: {nice}")

    if method.startswith("ocr"):
        if "image file" in method:
            notes.append("panel read by OCR from image file - verify against source image")
        else:
            notes.append("panel read by OCR - verify against source PDF")

    return "; ".join(notes)


# ==============================================================
# CONFIDENCE RATING
# ==============================================================
# A 0-100 heuristic so quality checkers know where to focus.  Penalties stack
# from the same signals the validator already records (calories/serving missing,
# Atwater mismatch, OCR usage, %DV estimation, OCR rescaling, implausible
# per-100g values, etc.).  The bucket label is a coarser version of the score:
#
#   High   (>=85) - panel parsed cleanly, no notable issues
#   Medium (65-84) - usable but worth a spot-check
#   Low    (1-64)  - several issues; review the source PDF before trusting
#   Failed (0)     - extraction errored out

def compute_confidence(row: dict, method: str, found_count: int,
                       estimated: List[str], rescaled: List[str],
                       approx: List[str], serving_g: Optional[float],
                       ingredients: Optional[str]) -> Tuple[int, str]:
    if method == "error":
        return 0, "Failed"

    score = 100

    if method.startswith("ocr"):
        score -= 15

    kcal = row.get("energy_kcal")
    if kcal is None:
        score -= 25

    if serving_g is None:
        score -= 20
    elif serving_g < 5 or serving_g > 800:
        # Most real serving sizes sit between ~10g and ~500g.  A handful of
        # condiments fall outside that band, but a value outside 5-800 g is
        # almost always a digit-grab mistake worth flagging.
        score -= 15

    fat = row.get("total_fat_g")
    carb = row.get("carb_g")
    prot = row.get("protein_g")
    if None not in (fat, carb, prot, kcal):
        atwater = 9 * fat + 4 * carb + 4 * prot
        if abs(atwater - kcal) > 0.25 * kcal + 25:
            score -= 25

    add = row.get("added_sugars_g")
    tot = row.get("total_sugars_g")
    if add is not None and tot is not None and add > tot + 0.5:
        score -= 15

    impl = 0
    for code in ("total_fat_g", "carb_g", "fiber_g", "protein_g",
                 "sat_fat_g", "total_sugars_g", "added_sugars_g"):
        v = row.get(f"{code}_per100g")
        if v is not None and v > 100:
            impl += 1
    score -= min(30, impl * 15)

    score -= min(15, len(estimated) * 3)
    score -= min(10, len(rescaled) * 5)
    score -= min(10, len(approx) * 2)

    if not ingredients:
        score -= 5
    if found_count < 6:
        score -= 15

    score = max(0, min(100, score))
    if score >= 85:
        rating = "High"
    elif score >= 65:
        rating = "Medium"
    elif score > 0:
        rating = "Low"
    else:
        rating = "Failed"
    return score, rating


# ==============================================================
# PER-PDF EXTRACTION
# ==============================================================

def parse_filename(stem: str) -> Tuple[str, str, str, str]:
    """`<id>_<description>_<brand>_<code>` -> (id, description, brand, code)."""
    parts = stem.split("_")
    if len(parts) >= 4:
        return parts[0], " ".join(parts[1:-2]).strip(), parts[-2].strip(), parts[-1].strip()
    if len(parts) == 3:
        return parts[0], parts[1].strip(), parts[2].strip(), ""
    return stem, stem, "", ""


def _panel_ok(panel: Dict[str, dict]) -> bool:
    """A panel is trusted when calories plus a handful of nutrients are read."""
    return panel.get("energy_kcal") is not None and panel_score(panel) >= 6


def extract_from_pdf(pdf_path: Path) -> dict:
    item_id, description, brand, code = parse_filename(pdf_path.stem)

    text_rows, raw, pages_words = extract_text(pdf_path)
    three_col = detect_three_column(raw)
    text_panel = parse_panel(panel_window(text_rows), three_col)

    # The big Calories figure sometimes prints on its own baseline.
    if "energy_kcal" not in text_panel:
        cal = calories_by_coordinates(pages_words)
        if cal is not None:
            text_panel["energy_kcal"] = dict(amount=cal, dv=None,
                                             per100g=None, lt=False)

    if _panel_ok(text_panel):
        panel, method, src_lines = text_panel, "text (layout-reconstructed)", text_rows
    else:
        try:
            ocr_text = ocr_collect(pdf_path)
            # Don't apply `panel_window` to OCR output. `ocr_collect` emits
            # cleanly-cropped panel images first and full-page renders last,
            # which can put panel rows *before* the "Nutrition Facts" anchor
            # found later in the full-page render. Trimming on that anchor
            # would discard the cleanest reads. The parser already tolerates
            # the marketing/footnote noise that anchor was meant to filter.
            ocr_lines = ocr_text.splitlines()
            three_col = three_col or detect_three_column(ocr_text)
            ocr_panel = parse_panel(ocr_lines, three_col)
            if panel_score(ocr_panel) > panel_score(text_panel):
                panel, method, src_lines = ocr_panel, "ocr", ocr_lines
            else:
                panel, method, src_lines = text_panel, "text (layout-reconstructed)", text_rows
        except Exception:
            # OCR unavailable (Tesseract not installed) or failed; fall back to
            # whatever text extraction found, even if the panel is incomplete.
            panel, method, src_lines = text_panel, "text (layout-reconstructed)", text_rows

    serving_g, serving_desc = parse_serving_size(src_lines, raw)
    ingredients = parse_ingredients(raw)

    row: dict = {
        "item_id": item_id,
        "source_file": pdf_path.name,
        "product_description": description,
        "brand": brand,
        "product_code": code,
        "serving_desc": serving_desc,
        "serving_size_g": serving_g,
        "extraction_method": method,
    }

    estimated: List[str] = []
    rescaled: List[str] = []
    approx: List[str] = []
    for spec in NUTRIENTS:
        c = spec["code"]
        reading = panel.get(c, {})
        amount = reading.get("amount")
        dv = reading.get("dv")

        amount, was_rescaled = reconcile_amount_with_dv(amount, dv, c)
        if was_rescaled:
            rescaled.append(c)
        amount, was_est = backfill_from_dv(amount, dv, c)
        if was_est:
            estimated.append(c)
        if reading.get("lt") and amount is not None:
            approx.append(f"{c} label-stated as <{amount:g} g")

        direct100 = reading.get("per100g")
        p100 = direct100 if direct100 is not None else per_100g(amount, serving_g)

        row[c] = amount
        row[f"{c}_per100g"] = p100

    row["ingredients"] = ingredients
    notes = validate(row, serving_g, method, estimated, rescaled)
    if approx:
        notes = "; ".join(filter(None, [notes, "; ".join(approx)]))
    row["validation_notes"] = notes

    found_count = sum(1 for c in NUTRIENT_CODES if row.get(c) is not None)
    score, rating = compute_confidence(
        row, method, found_count, estimated, rescaled, approx, serving_g, ingredients
    )
    row["confidence"] = score
    row["confidence_rating"] = rating
    return row


def extract_from_image(img_path: Path) -> dict:
    """
    Extract a Nutrition Facts panel and ingredient statement from a standalone
    image file (photo of a nutrition label, product-spec screenshot, etc.).

    Because there is no embedded vector text to fall back on, the pipeline goes
    straight to OCR using :func:`ocr_from_image_file`.  Everything downstream
    (panel parsing, serving-size recovery, per-100 g conversion, validation, and
    confidence scoring) is identical to the PDF path.
    """
    item_id, description, brand, code = parse_filename(img_path.stem)

    ocr_text = ocr_from_image_file(img_path)
    ocr_lines = ocr_text.splitlines()

    three_col = detect_three_column(ocr_text)
    panel = parse_panel(ocr_lines, three_col)
    method = "ocr (image file)"

    # Use the OCR'd text as the raw source for serving-size and ingredient
    # extraction (plays the role of the PDF's page text in extract_from_pdf).
    serving_g, serving_desc = parse_serving_size(ocr_lines, ocr_text)
    ingredients = parse_ingredients(ocr_text)

    row: dict = {
        "item_id": item_id,
        "source_file": img_path.name,
        "product_description": description,
        "brand": brand,
        "product_code": code,
        "serving_desc": serving_desc,
        "serving_size_g": serving_g,
        "extraction_method": method,
    }

    estimated: List[str] = []
    rescaled: List[str] = []
    approx: List[str] = []
    for spec in NUTRIENTS:
        c = spec["code"]
        reading = panel.get(c, {})
        amount = reading.get("amount")
        dv = reading.get("dv")

        amount, was_rescaled = reconcile_amount_with_dv(amount, dv, c)
        if was_rescaled:
            rescaled.append(c)
        amount, was_est = backfill_from_dv(amount, dv, c)
        if was_est:
            estimated.append(c)
        if reading.get("lt") and amount is not None:
            approx.append(f"{c} label-stated as <{amount:g} g")

        direct100 = reading.get("per100g")
        p100 = direct100 if direct100 is not None else per_100g(amount, serving_g)

        row[c] = amount
        row[f"{c}_per100g"] = p100

    row["ingredients"] = ingredients
    notes = validate(row, serving_g, method, estimated, rescaled)
    if approx:
        notes = "; ".join(filter(None, [notes, "; ".join(approx)]))
    row["validation_notes"] = notes

    found_count = sum(1 for c in NUTRIENT_CODES if row.get(c) is not None)
    score, rating = compute_confidence(
        row, method, found_count, estimated, rescaled, approx, serving_g, ingredients
    )
    row["confidence"] = score
    row["confidence_rating"] = rating
    return row


def batch_extract(folder: Path) -> List[dict]:
    rows: List[dict] = []
    files = sorted(folder.iterdir(), key=lambda p: p.name.lower())
    for file_path in files:
        suffix = file_path.suffix.lower()
        is_pdf = suffix == ".pdf"
        is_image = suffix in IMAGE_EXTENSIONS
        if not (is_pdf or is_image):
            continue
        try:
            if is_pdf:
                print(f"  extracting (pdf):   {file_path.name}")
                rows.append(extract_from_pdf(file_path))
            else:
                print(f"  extracting (image): {file_path.name}")
                rows.append(extract_from_image(file_path))
        except Exception as exc:  # keep the batch alive
            item_id, description, brand, code = parse_filename(file_path.stem)
            err = {
                "item_id": item_id, "source_file": file_path.name,
                "product_description": description, "brand": brand,
                "product_code": code, "serving_desc": None,
                "serving_size_g": None, "extraction_method": "error",
                "ingredients": None, "validation_notes": f"ERROR: {exc}",
                "confidence": 0, "confidence_rating": "Failed",
            }
            for c in NUTRIENT_CODES:
                err[c] = None
                err[f"{c}_per100g"] = None
            rows.append(err)
    return rows


# ==============================================================
# OUTPUT
# ==============================================================

def build_extraction_frame(rows: List[dict]) -> pd.DataFrame:
    meta = ["item_id", "source_file", "product_description", "brand",
            "product_code", "serving_desc", "serving_size_g", "extraction_method",
            "confidence", "confidence_rating"]
    amount_cols = NUTRIENT_CODES
    per100_cols = [f"{c}_per100g" for c in NUTRIENT_CODES]
    ordered = meta + amount_cols + per100_cols + ["validation_notes", "ingredients"]
    df = pd.DataFrame(rows)
    for col in ordered:
        if col not in df.columns:
            df[col] = None
    return df[ordered]


def build_fndds_frame(rows: List[dict]) -> pd.DataFrame:
    """
    Per-100 g view laid out for FNDDS. Each nutrient column is headed by its
    FNDDS nutrient number and description; values are per 100 g edible portion,
    the basis FNDDS uses.
    """
    records: List[dict] = []
    for r in rows:
        rec = {
            "item_id": r.get("item_id"),
            "product_description": r.get("product_description"),
            "brand": r.get("brand"),
            "product_code": r.get("product_code"),
            "serving_size_g": r.get("serving_size_g"),
            "confidence": r.get("confidence"),
            "confidence_rating": r.get("confidence_rating"),
        }
        for spec in NUTRIENTS:
            tag = (f"{spec['fndds_no']} - {spec['fndds_name']}"
                   if spec["fndds_no"] else spec["fndds_name"])
            rec[tag] = r.get(f"{spec['code']}_per100g")
        rec["validation_notes"] = r.get("validation_notes")
        records.append(rec)
    return pd.DataFrame(records)


def write_workbook(rows: List[dict], output_path: Path) -> None:
    extraction = build_extraction_frame(rows)
    fndds = build_fndds_frame(rows)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        extraction.to_excel(writer, sheet_name="Extraction", index=False)
        fndds.to_excel(writer, sheet_name="FNDDS_per_100g", index=False)

        for sheet, frame in (("Extraction", extraction),
                             ("FNDDS_per_100g", fndds)):
            ws = writer.sheets[sheet]
            ws.freeze_panes = "A2"
            for idx, col in enumerate(frame.columns, start=1):
                letter = ws.cell(row=1, column=idx).column_letter
                width = max(len(str(col)), 12)
                if col in ("ingredients", "validation_notes",
                           "serving_desc", "product_description"):
                    width = 48
                ws.column_dimensions[letter].width = min(width + 2, 52)


# ==============================================================
# MAIN
# ==============================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Extract nutrition facts from PDFs and/or image files "
            "(JPG, PNG, TIFF, WebP, BMP) to Excel."
        )
    )
    parser.add_argument(
        "-i", "--input",
        default="pdfs",
        metavar="FOLDER",
        help=(
            "Input folder containing PDFs and/or image files "
            "(JPG, PNG, TIFF, WebP, BMP) to process (default: pdfs)"
        ),
    )
    parser.add_argument(
        "-o", "--output",
        default="output/nutrition_extraction.xlsx",
        metavar="FILE",
        help="Output Excel filename (default: output/nutrition_extraction.xlsx)",
    )
    args = parser.parse_args()

    input_folder = Path(args.input)
    output_path = Path(args.output)

    print(f"Scanning PDFs and images in: {input_folder.resolve()}")
    rows = batch_extract(input_folder)
    write_workbook(rows, output_path)

    print(f"\nProcessed {len(rows)} file(s).")
    for r in rows:
        found = sum(1 for c in NUTRIENT_CODES if r.get(c) is not None)
        print(f"  - {r['source_file']}: {found}/{len(NUTRIENT_CODES)} nutrients"
              f" | serving={r.get('serving_size_g')} g"
              f" | {r.get('extraction_method')}")
        if r.get("validation_notes"):
            print(f"      notes: {r['validation_notes']}")
    print(f"\nSaved workbook: {output_path.resolve()}")


if __name__ == "__main__":
    main()
