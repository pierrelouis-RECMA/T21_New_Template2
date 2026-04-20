import os, re, shutil, subprocess, zipfile, tempfile
import pandas as pd
import xml.etree.ElementTree as ET
from pptx import Presentation

# ── CONFIGURATION & NAMESPACES ──
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
ET.register_namespace('a', A); ET.register_namespace('p', P)
ta = lambda l: f'{{{A}}}{l}'
tp = lambda l: f'{{{P}}}{l}'

# Ta géométrie A4 Portrait
SW, SH = 7559675, 10439400
NAV_H, MARG_X, MARG_TOP = 370000, 220000, 100000
FOOTER_Y = SH - 170000 - 30000
CARD_W, CARD_GAP = SW - 2 * MARG_X, 100000

# Ta Palette
C_HEADER_BG, C_HEADER_TXT = "0F172A", "FFFFFF"
C_WIN_TXT, C_DEP_TXT, C_RET_TXT = "059669", "BE123C", "B45309"
C_CARD_BG, C_COL_BG, C_ITEM_TXT = "FFFFFF", "F8FAFC", "334155"
C_BORDER, C_NAV_BG, C_NAV_TXT = "E2E8F0", "0F172A", "64748B"
FONT = "Segoe UI"

_id = [1000]
def nid():
    _id[0] += 1
    return _id[0]

# ── TES FONCTIONS DE DESSIN (Stricte copie de ton v11) ──
def mk_sp(id_, name, x, y, w, h, geom="rect", fill=None, corner=None, border_c=None, border_w=4762):
    sp = ET.Element(tp('sp'))
    nv = ET.SubElement(sp, tp('nvSpPr'))
    ET.SubElement(nv, tp('cNvPr'), id=str(id_), name=name)
    ET.SubElement(nv, tp('cNvSpPr')); ET.SubElement(nv, tp('nvPr'))
    pr = ET.SubElement(sp, tp('spPr'))
    xf = ET.SubElement(pr, ta('xfrm'))
    ET.SubElement(xf, ta('off'), x=str(x), y=str(y))
    ET.SubElement(xf, ta('ext'), cx=str(w), cy=str(h))
    pg = ET.SubElement(pr, ta('prstGeom'), prst=geom)
    if fill:
        sf = ET.SubElement(pr, ta('solidFill'))
        ET.SubElement(sf, ta('srgbClr')).set('val', fill)
    return sp

def mk_tx(id_, name, x, y, w, h, runs, algn='l', anchor='t'):
    sp = mk_sp(id_, name, x, y, w, h)
    tx = ET.SubElement(sp, tp('txBody'))
    ET.SubElement(tx, ta('bodyPr'), anchor=anchor, lIns="91440", rIns="91440")
    ET.SubElement(tx, ta('lstStyle'))
    p = ET.SubElement(tx, ta('p'))
    ET.SubElement(p, ta('pPr'), algn=algn)
    for text, sz, bold, color in runs:
        r = ET.SubElement(p, ta('r'))
        rpr = ET.SubElement(r, ta('rPr'), sz=str(sz), b=('1' if bold else '0'))
        sf = ET.SubElement(rpr, ta('solidFill'))
        ET.SubElement(sf, ta('srgbClr')).set('val', color)
        ET.SubElement(rpr, ta('latin'), typeface=FONT)
        ET.SubElement(r, ta('t')).text = text
    return sp

# ── LOGIQUE DE DONNÉES ──
def load_data(file_path):
    df = pd.read_csv(file_path) if file_path.endswith('.csv') else pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    df['Agency'] = df['Agency'].astype(str).str.upper()
    
    agencies = []
    for ag in df['Agency'].unique():
        if ag in ['NAN', '']: continue
        sub = df[df['Agency'] == ag]
        w = sub[sub['NewBiz']=='WIN']['Integrated Spends'].sum()
        d = sub[sub['NewBiz']=='DEPARTURE']['Integrated Spends'].sum()
        
        agencies.append({
            "name": ag, "group": "MEXICO", "nbb": f"{w+d:+.1f}m$",
            "wins": [(r['Advertiser'], f"+{r['Integrated Spends']:.1f}m") for _, r in sub[sub['NewBiz']=='WIN'].iterrows()],
            "deps": [(r['Advertiser'], f"{r['Integrated Spends']:.1f}m") for _, r in sub[sub['NewBiz']=='DEPARTURE'].iterrows()],
            "rets": [(r['Advertiser'], "") for _, r in sub[sub['NewBiz']=='RETENTION'].iterrows()],
            "val": w+d
        })
    agencies.sort(key=lambda x: x['val'], reverse=True)
    return {i//4 + 22: agencies[i:i+4] for i in range(0, len(agencies), 4)}

# ── ENTRÉE RENDER ──
def generate_report(input_excel_path):
    # Charge les données
    data_by_slide = load_data(input_excel_path)
    
    # Utilise le template existant
    template = os.path.join(os.path.dirname(__file__), 'T21_HK_Agencies_Glass_v12.pptx')
    prs = Presentation(template)
    
    # On met à jour les slides simples (Pays)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        r.text = r.text.replace("Hong Kong", "MEXICO").replace("HONG KONG", "MEXICO")

    out_path = os.path.join(tempfile.gettempdir(), "NBB_Report_Mexico.pptx")
    prs.save(out_path)
    return out_path
