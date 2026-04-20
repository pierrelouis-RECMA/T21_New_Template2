"""
pdf_nav.py — Ajoute navigation cliquable au PDF
  1. Bookmarks sidebar (panneau gauche du lecteur PDF)
  2. TOC page 1 cliquable
  3. Boutons nav bas de page

Usage: python3 pdf_nav.py input.pdf  →  input_nav.pdf
"""
import os, sys
from pypdf import PdfReader, PdfWriter

SLIDES = [
    {"title": "Content",                   "page": 0},
    {"title": "Key Findings",              "page": 1},
    {"title": "TOP moves 2025",            "page": 2},
    {"title": "T21a – Agencies overview",  "page": 3},
    {"title": "T21b – Group overview",     "page": 4},
    {"title": "T21c – Retentions ranking", "page": 5},
    {"title": "T21d – Details by agency",  "page": 6},
]

# Positions des lignes TOC (% de hauteur depuis le bas, page 1)
TOC_ROWS = [
    {"slide_idx": 1, "y_min_pct": 0.565, "y_max_pct": 0.615},
    {"slide_idx": 2, "y_min_pct": 0.510, "y_max_pct": 0.560},
    {"slide_idx": 3, "y_min_pct": 0.455, "y_max_pct": 0.505},
    {"slide_idx": 4, "y_min_pct": 0.400, "y_max_pct": 0.450},
    {"slide_idx": 5, "y_min_pct": 0.345, "y_max_pct": 0.395},
    {"slide_idx": 6, "y_min_pct": 0.290, "y_max_pct": 0.340},
]

def make_link(x0, y0, x1, y1, target_page):
    """Annotation de lien PDF (format dict pypdf)."""
    from pypdf.generic import (
        ArrayObject, DictionaryObject, NameObject, NumberObject
    )
    return DictionaryObject({
        NameObject("/Type"):    NameObject("/Annot"),
        NameObject("/Subtype"): NameObject("/Link"),
        NameObject("/Rect"):    ArrayObject([
            NumberObject(round(x0, 2)), NumberObject(round(y0, 2)),
            NumberObject(round(x1, 2)), NumberObject(round(y1, 2)),
        ]),
        NameObject("/Border"): ArrayObject([
            NumberObject(0), NumberObject(0), NumberObject(0)
        ]),
        NameObject("/A"): DictionaryObject({
            NameObject("/Type"): NameObject("/Action"),
            NameObject("/S"):    NameObject("/GoTo"),
            NameObject("/D"):    ArrayObject([
                NumberObject(target_page),
                NameObject("/Fit"),
            ]),
        }),
    })

def add_pdf_navigation(input_pdf: str, output_pdf: str) -> str:
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    n = len(writer.pages)
    print(f"  {n} pages")

    # ── 1. Bookmarks sidebar ──────────────────────────────────────────────────
    print("  Bookmarks sidebar...")
    for s in SLIDES:
        if s["page"] < n:
            writer.add_outline_item(s["title"], s["page"])

    # ── 2. Liens TOC page 1 ───────────────────────────────────────────────────
    print("  Liens TOC page 1...")
    if n > 0:
        p0 = writer.pages[0]
        pw = float(p0.mediabox.width)
        ph = float(p0.mediabox.height)

        from pypdf.generic import ArrayObject, NameObject
        if "/Annots" not in p0:
            p0[NameObject("/Annots")] = ArrayObject()

        x0, x1 = pw * 0.15, pw * 0.90
        for row in TOC_ROWS:
            if row["slide_idx"] < n:
                y0 = ph * row["y_min_pct"]
                y1 = ph * row["y_max_pct"]
                annot = make_link(x0, y0, x1, y1, row["slide_idx"])
                writer._add_object(annot)
                p0["/Annots"].append(annot.indirect_reference or annot)

    # ── 3. Nav bas de chaque page ─────────────────────────────────────────────
    print("  Navigation inter-pages...")
    from pypdf.generic import ArrayObject, NameObject

    for pi in range(n):
        page = writer.pages[pi]
        pw   = float(page.mediabox.width)
        ph   = float(page.mediabox.height)

        if "/Annots" not in page:
            page[NameObject("/Annots")] = ArrayObject()

        nav_y0 = ph * 0.005
        nav_y1 = ph * 0.045

        # Boutons numérotés centrés
        margin  = pw * 0.25
        btn_w   = (pw * 0.50) / n
        for i, s in enumerate(SLIDES):
            if s["page"] >= n: continue
            x0 = margin + i * btn_w
            x1 = x0 + btn_w * 0.85
            annot = make_link(x0, nav_y0, x1, nav_y1, s["page"])
            writer._add_object(annot)
            page["/Annots"].append(annot)

        # ◀ Précédent
        if pi > 0:
            a = make_link(pw*0.01, nav_y0, pw*0.12, nav_y1, pi-1)
            writer._add_object(a); page["/Annots"].append(a)

        # Suivant ▶
        if pi < n-1:
            a = make_link(pw*0.88, nav_y0, pw*0.99, nav_y1, pi+1)
            writer._add_object(a); page["/Annots"].append(a)

    # ── Sauvegarde ────────────────────────────────────────────────────────────
    with open(output_pdf, "wb") as f:
        writer.write(f)

    size = os.path.getsize(output_pdf)/1024
    print(f"  ✅ {output_pdf} ({size:.0f} KB)")
    return output_pdf

if __name__ == "__main__":
    inp = sys.argv[1] if len(sys.argv) > 1 else "/home/claude/HK_NBB_Report_2025_AUTO.pdf"
    out = inp.replace(".pdf", "_nav.pdf")
    add_pdf_navigation(inp, out)
