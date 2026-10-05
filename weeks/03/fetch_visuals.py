# -*- coding: utf-8 -*-
"""Download ready figures for L3 (Commons, arXiv 2110.11255, textbook page renders)."""
from __future__ import annotations

import io
import ssl
import time
import urllib.request
from pathlib import Path

import pymupdf
from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
FIG = Path(__file__).resolve().parent / "figures"
PHOTO = FIG / "photo"
PAPER = FIG / "paper"
BOOK = FIG / "book"

UA = "IITP26-course-slides/1.0 (educational; mailto:none)"
CTX = ssl.create_default_context()

COMMONS = {
    # sources — real lamps / sky / spectra
    "src_bulb.jpg": "Incandescent_light_bulb.jpg",
    "src_cfl.jpg": "Compact_fluorescent_lamp.jpg",
    "src_led.jpg": "LED_lamp.jpg",
    "src_tube.jpg": "Fluorescent_light_fixture.jpg",
    "src_sun.jpg": "The_Sun_by_SDO_(2010).jpg",
    "src_overcast.jpg": "Overcast.jpg",
    "src_window.jpg": "Window.jpg",
    "src_spectra_photo.jpg": "Various_lighting_spectrums_-_Flurescent_incandescent_diode_and_candle.jpg",
    "src_fluor_cd.jpg": "Fluorescent_lamp_spectrum.jpg",
    # surfaces
    "surf_velvet.jpg": "Red_velvet.jpg",
    "surf_silk.jpg": "Silk.jpg",
    "surf_chrome.jpg": "Chromium.jpg",
    "surf_cd.jpg": "Compact_Disc.jpg",
    "surf_paper.jpg": "Paper.jpg",
    "surf_car.jpg": "Car_paint.jpg",
    "surf_peacock.jpg": "Indian_Peafowl.jpg",
    "surf_apple.jpg": "Red_Apple.jpg",
    # eye / retina
    "eye_fundus.jpg": "Fundus_photograph-normal_retina_EDA06.JPG",
    "eye_mosaic.jpg": "ConeMosaics.jpg",
    "eye_rods_cones.jpg": "1414_Rods_and_Cones.jpg",
    "eye_opponent.jpg": "1416_Color_Processing.jpg",
    "eye_anatomy.jpg": "Schematic_diagram_of_the_human_eye_en.svg",
    "eye_dist.png": "Distribution_of_Cones_and_Rods_on_Human_Retina.png",
    # theory (ready diagrams)
    "cie_xy.png": "CIExy1931.png",
    "cie_xy_blank.svg": "CIE1931xy_blank.svg",
    "cie_macadam.png": "CIExy1931_MacAdam.png",
    "cie_gamut.jpg": "CIE_1931_Chromaticy_Diagram_and_Gamut.jpg",
}


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, context=CTX, timeout=60) as r:
        return r.read()


def commons(name: str, dest: Path) -> bool:
    dest.parent.mkdir(parents=True, exist_ok=True)
    url = "https://commons.wikimedia.org/wiki/Special:FilePath/" + urllib.request.quote(name)
    try:
        data = fetch(url)
        if len(data) < 2000:
            print("tiny", name, len(data))
            return False
        dest.write_bytes(data)
        print("ok", dest.name, len(data))
        return True
    except Exception as e:
        print("fail", name, type(e).__name__, e)
        return False


def save_pix(pix, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)
    if pix.n >= 4 and pix.alpha:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    elif pix.n == 4:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    pix.save(str(dest))


def extract_pdf_images(pdf: Path, out: Path, prefix: str, min_w=80):
    doc = pymupdf.open(str(pdf))
    n = 0
    for pi, page in enumerate(doc):
        for j, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            try:
                pix = pymupdf.Pixmap(doc, xref)
            except Exception:
                continue
            if pix.width < min_w or pix.height < min_w:
                continue
            n += 1
            save_pix(pix, out / f"{prefix}_p{pi+1:03d}_{j:02d}_{pix.width}x{pix.height}.png")
    print("extracted", n, "from", pdf.name)
    return n


def render_pages(pdf: Path, out: Path, pages: list[int], prefix: str, zoom=2.0):
    doc = pymupdf.open(str(pdf))
    mat = pymupdf.Matrix(zoom, zoom)
    out.mkdir(parents=True, exist_ok=True)
    for p in pages:
        if p < 1 or p > doc.page_count:
            continue
        pix = doc[p - 1].get_pixmap(matrix=mat, alpha=False)
        dest = out / f"{prefix}_page{p:03d}.png"
        pix.save(str(dest))
        print("render", dest.name)


def find_pages(pdf: Path, needles: list[str], limit=40):
    doc = pymupdf.open(str(pdf))
    hits = []
    for i in range(doc.page_count):
        t = doc[i].get_text().lower()
        if any(n.lower() in t for n in needles):
            hits.append(i + 1)
            if len(hits) >= limit:
                break
    return hits


def main():
    PHOTO.mkdir(parents=True, exist_ok=True)
    PAPER.mkdir(parents=True, exist_ok=True)
    BOOK.mkdir(parents=True, exist_ok=True)

    for dest, wiki in COMMONS.items():
        path = PHOTO / dest
        if path.exists() and path.stat().st_size > 2000:
            print("have", dest)
            continue
        commons(wiki, path)
        time.sleep(0.4)

    arxiv = FIG / "2110.11255.pdf"
    if not arxiv.exists():
        print("arxiv pdf...")
        arxiv.write_bytes(fetch("https://arxiv.org/pdf/2110.11255.pdf"))
    extract_pdf_images(arxiv, PAPER, "josaa")
    # also full-page renders of likely figure pages
    render_pages(arxiv, PAPER, [1, 2, 3, 4, 5, 6, 7, 8], "josaa", zoom=2.2)

    books = ROOT / "books"
    maximov = books / "Maximov_Transformation_of_color_space.pdf"
    if maximov.exists():
        pages = find_pages(
            maximov,
            ["цветовой конус", "цветовое тело", "ступенчат", "дополнител", "окраск"],
            limit=25,
        )
        (BOOK / "maximov_pages.txt").write_text(
            "\n".join(str(p) for p in pages), encoding="utf-8"
        )
        print("maximov hits", pages)
        extract_pdf_images(maximov, BOOK, "maximov", min_w=120)
        render_pages(maximov, BOOK, pages[:18], "maximov", zoom=1.8)

    ohta = books / "Ohta_Robertson_Colorimetry.pdf"
    if ohta.exists():
        pages = find_pages(ohta, ["Wright", "Guild", "color-matching function"], limit=12)
        print("ohta hits", pages)
        render_pages(ohta, BOOK, pages[:8], "ohta", zoom=1.8)

    shevell = books / "Shevell_The_Science_of_Color.pdf"
    if shevell.exists():
        pages = find_pages(shevell, ["retina", "photoreceptor", "ganglion"], limit=10)
        print("shevell hits", pages)
        render_pages(shevell, BOOK, pages[:6], "shevell", zoom=1.7)


if __name__ == "__main__":
    main()
