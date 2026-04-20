import os
import pandas as pd
from pptx import Presentation
import xml.etree.ElementTree as ET

# --- Tes constantes de design ---
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
ta = lambda l: f'{{{A}}}{l}'
tp = lambda l: f'{{{P}}}{l}'

def nid():
    if not hasattr(nid, "id"): nid.id = 1000
    nid.id += 1
    return nid.id

# --- Tes fonctions de construction XML (v11) ---
def mk_sp(id_, name, x, y, w, h, fill=None):
    sp = ET.Element(tp('sp'))
    nv = ET.SubElement(sp, tp('nvSpPr'))
    ET.SubElement(nv, tp('cNvPr'), id=str(id_), name=name)
    ET.SubElement(nv, tp('cNvSpPr')); ET.SubElement(nv, tp('nvPr'))
    pr = ET.SubElement(sp, tp('spPr'))
    xf = ET.SubElement(pr, ta('xfrm'))
    ET.SubElement(xf, ta('off'), x=str(x), y=str(y))
    ET.SubElement(xf, ta('ext'), cx=str(w), cy=str(h))
    ET.SubElement(pr, ta('prstGeom'), prst="rect").append(ET.Element(ta('avLst')))
    if fill:
        sf = ET.SubElement(pr, ta('solidFill'))
        ET.SubElement(sf, ta('srgbClr')).set('val', fill)
    return sp

# --- La fonction principale pour Render ---
def generate_report(input_excel_path):
    # 1. Charger les données Mexique
    df = pd.read_csv(input_excel_path) if input_excel_path.endswith('.csv') else pd.read_excel(input_excel_path)
    
    # 2. Ouvrir le template
    template_path = os.path.join(os.path.dirname(__file__), 'T21_HK_Agencies_Glass_v12.pptx')
    prs = Presentation(template_path)
    
    # 3. Mise à jour globale (Slide 1 à la fin)
    # On remplace "Hong Kong" par "Mexico" partout
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if "Hong Kong" in run.text:
                            run.text = run.text.replace("Hong Kong", "Mexico")
                        if "HONG KONG" in run.text:
                            run.text = run.text.replace("HONG KONG", "MEXICO")

    # 4. Focus spécifique Slide 3 (Key Findings)
    # Ici, on peut forcer des valeurs si besoin, ou laisser le remplacement global agir
    # Si tu as des chiffres spécifiques (ex: total market), on les injecte ici.

    # 5. Sauvegarde
    output_filename = "NBB_Mexico_2025.pptx"
    output_path = os.path.join("/tmp", output_filename)
    prs.save(output_path)
    
    return output_path
