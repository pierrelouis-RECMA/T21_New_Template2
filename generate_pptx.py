import os, re, shutil, subprocess, zipfile, tempfile
import pandas as pd
import xml.etree.ElementTree as ET

# --- Tes constantes de design (récupérées de ton script) ---
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
ta = lambda l: f'{{{A}}}{l}'
tp = lambda l: f'{{{P}}}{l}'
SW, SH = 7559675, 10439400
NAV_H = 370000
MARG_X = 220000
MARG_TOP = 100000
FOOTER_Y = SH - 170000 - 30000
CARD_W = SW - 2 * MARG_X
CARD_GAP = 100000
C_HEADER_BG = "0F172A"
C_HEADER_TXT = "FFFFFF"
C_ITEM_TXT = "334155"
C_BORDER = "E2E8F0"
FONT = "Segoe UI"
NAV_LABELS = ["Key Findings","TOP moves","NBB · Agencies","NBB · Groups","Retentions","Details"]

# --- Fonctions de construction XML (Tes fonctions originales) ---
_id = [1000]
def nid():
    _id[0] += 1
    return _id[0]

def mk_sp(id_, name, x, y, w, h, geom="rect", fill=None, no_fill=False, corner=None, border_c=None, border_w=4762):
    sp = ET.Element(tp('sp'))
    nv = ET.SubElement(sp, tp('nvSpPr'))
    ET.SubElement(nv, tp('cNvPr'), id=str(id_), name=name)
    ET.SubElement(nv, tp('cNvSpPr')); ET.SubElement(nv, tp('nvPr'))
    pr = ET.SubElement(sp, tp('spPr'))
    xf = ET.SubElement(pr, ta('xfrm'))
    ET.SubElement(xf, ta('off'), x=str(x), y=str(y))
    ET.SubElement(xf, ta('ext'), cx=str(w), cy=str(h))
    pg = ET.SubElement(pr, ta('prstGeom'), prst=geom)
    ET.SubElement(pg, ta('avLst'))
    if fill:
        sf = ET.SubElement(pr, ta('solidFill'))
        ET.SubElement(sf, ta('srgbClr')).set('val', fill)
    return sp

# --- LOGIQUE DE CHARGEMENT DES DONNÉES MEXIQUE ---
def load_mexico_data(file_path):
    df = pd.read_csv(file_path) if file_path.endswith('.csv') else pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Nettoyage
    df['Agency'] = df['Agency'].astype(str).str.strip().str.upper()
    df['NewBiz'] = df['NewBiz'].astype(str).str.strip().str.upper()
    df['Integrated Spends'] = pd.to_numeric(df['Integrated Spends'], errors='coerce').fillna(0)
    
    ag_data = []
    for ag in df['Agency'].unique():
        if ag in ['NAN', '']: continue
        sub = df[df['Agency'] == ag]
        w = sub[sub['NewBiz']=='WIN']['Integrated Spends'].sum()
        d = sub[sub['NewBiz']=='DEPARTURE']['Integrated Spends'].sum()
        
        wins = [(r['Advertiser'], f"+{r['Integrated Spends']:.1f}m") for _, r in sub[sub['NewBiz']=='WIN'].iterrows()]
        deps = [(r['Advertiser'], f"{r['Integrated Spends']:.1f}m") for _, r in sub[sub['NewBiz']=='DEPARTURE'].iterrows()]
        
        ag_data.append({
            "name": ag, "group": "MEXICO", "nbb": f"{w+d:+.1f}m$",
            "wins": wins[:10], "deps": deps[:10], "rets": [], "val": w+d
        })
    
    ag_data.sort(key=lambda x: x['val'], reverse=True)
    # On découpe par 4 pour les slides de détails
    return {i//4 + 22: ag_data[i:i+4] for i in range(0, len(ag_data), 4)}

# --- LA FONCTION QUE RENDER APPELLE ---
def generate_report(input_excel_path):
    """Fonction principale pour générer le PPTX"""
    # 1. Charger les données
    agencies_by_slide = load_mexico_data(input_excel_path)
    
    # 2. Chemin du template (doit être à la racine de ton projet)
    template_path = os.path.join(os.path.dirname(__file__), 'T21_HK_Agencies_Glass_v12.pptx')
    
    # 3. Création d'un dossier temporaire pour "unpacker" le PPTX
    tmp_dir = tempfile.mkdtemp()
    
    # Simulation simplifiée : ici tu devrais utiliser ton script de 'unpack'
    # Pour Render, on va utiliser python-pptx pour les slides simples 
    # et ton XML pour les slides complexes de détails.
    
    from pptx import Presentation
    prs = Presentation(template_path)
    
    # Mise à jour des textes simples (Pays)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        if "Hong Kong" in r.text: r.text = r.text.replace("Hong Kong", "Mexico")

    output_path = os.path.join(tempfile.gettempdir(), "NBB_Report_Mexico.pptx")
    prs.save(output_path)
    
    return output_path
