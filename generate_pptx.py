import copy, io, os, shutil, warnings, sys
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib import rcParams
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.oxml.ns import qn

# Configuration pour éviter les avertissements inutiles
warnings.filterwarnings('ignore')
rcParams['font.family'] = 'DejaVu Sans'

# --- CONFIGURATION ET PALETTE ---
C_HEADER, C_WIN, C_DEP, C_RET = '#2D5C54', '#4CAF50', '#E53935', '#FF9800'
C_ALT, C_TOTAL, C_NBB_POS, C_NBB_NEG = '#F2F6F5', '#E8F5E9', '#1B5E20', '#B71C1C'

GROUP_COLORS = {
    'Publicis Media': '#FFEADD', 'Omnicom Media': '#FFE699', 
    'Dentsu': '#E2F0D9', 'Havas Media Network': '#DCB9FF', 'WPP Media': '#FFE4FF'
}
GROUP_ORDER = ['Publicis Media','Omnicom Media','Dentsu','Havas Media Network','WPP Media']

AGENCY_GROUP = {
    'SPARK FOUNDRY':'Publicis Media','STARCOM':'Publicis Media','ZENITH':'Publicis Media',
    'PHD':'Omnicom Media','OMD':'Omnicom Media','UM/INITIATIVE':'Omnicom Media',
    'HEARTS & SCIENCE':'Omnicom Media','DENTSU X':'Dentsu','IPROSPECT':'Dentsu',
    'CARAT':'Dentsu','HAVAS MEDIA':'Havas Media Network','ESSENCEMEDIACOM':'WPP Media',
    'WAVEMAKER':'WPP Media','MINDSHARE':'WPP Media'
}

# Variable globale pour l'app Render
EXCEL = ""

# --- LOGIQUE DE CALCUL ---
def load_stats():
    # Utilise le fichier Excel défini globalement par l'upload
    df = pd.read_excel(EXCEL, sheet_name=0)
    df['Agency'] = df['Agency'].str.strip().str.upper()
    df['NewBiz'] = df['NewBiz'].str.strip().str.upper()
    df['Group'] = df['Agency'].map(AGENCY_GROUP).fillna('Autres')

    def fmt_date(v):
        try: return pd.to_datetime(v).strftime('%b-%y')
        except: return ''

    df['Date_str'] = df['Date of announcement'].apply(fmt_date)
    market_detected = str(df['Country of Decision'].iloc[0]) if 'Country of Decision' in df.columns else "Mexico"

    agencies = []
    for ag in df.Agency.unique():
        sub = df[df.Agency == ag]
        w = sub[sub.NewBiz == 'WIN']['Integrated Spends'].sum()
        d = sub[sub.NewBiz == 'DEPARTURE']['Integrated Spends'].sum()
        r = sub[sub.NewBiz == 'RETENTION']['Integrated Spends'].sum()
        agencies.append({
            'agency': ag, 'group': AGENCY_GROUP.get(ag,'Autres'),
            'nbb': w+d, 'wins': w, 'dep': d, 'ret': r,
            'wc': len(sub[sub.NewBiz == 'WIN']), 'dc': len(sub[sub.NewBiz == 'DEPARTURE']),
            'wins_rows': sub[sub.NewBiz == 'WIN'].to_dict('records'),
            'dep_rows': sub[sub.NewBiz == 'DEPARTURE'].to_dict('records'),
            'ret_rows': sub[sub.NewBiz == 'RETENTION'].to_dict('records'),
        })
    agencies.sort(key=lambda x: -x['nbb'])
    for i, a in enumerate(agencies): a['rank'] = i+1

    group_stats = {g: {'group':g, 'nbb': sum(a['nbb'] for a in agencies if a['group']==g),
                       'wins': sum(a['wins'] for a in agencies if a['group']==g),
                       'dep': sum(a['dep'] for a in agencies if a['group']==g),
                       'wc': sum(a['wc'] for a in agencies if a['group']==g),
                       'dc': sum(a['dc'] for a in agencies if a['group']==g)} for g in GROUP_ORDER}

    top_moves = df[df.NewBiz.isin(['WIN','RETENTION'])].copy()
    top_moves['IS_abs'] = top_moves['Integrated Spends'].abs()
    top_moves = top_moves.sort_values('IS_abs', ascending=False).head(24)

    return agencies, group_stats, top_moves, df, market_detected

# --- FONCTIONS DE REMPLACEMENT ---
def replace_text_in_pptx(prs, old_text, new_text):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if old_text in run.text: run.text = run.text.replace(old_text, new_text)
                        if old_text.upper() in run.text: run.text = run.text.replace(old_text.upper(), new_text.upper())

def replace_slide_image(prs, slide_index, img_buf, _unused):
    from pptx.parts.image import ImagePart
    slide = prs.slides[slide_index]
    img_buf.seek(0)
    img_bytes = img_buf.read()
    for shape in slide.shapes:
        if shape.shape_type == 13: # PICTURE
            rId = shape._element.find('.//' + qn('a:blip')).get(qn('r:embed'))
            slide.part.related_part(rId)._blob = img_bytes
            return True
    return False

# --- GÉNÉRATION DES GRAPHIQUES (STUBS) ---
def make_slide3_img(top_moves): return io.BytesIO() # Remplacé par design moderne
def make_slide4_img(agencies): return io.BytesIO() # Remplacé par design moderne

def make_slide6_chart(agencies):
    fig, ax = plt.subplots(figsize=(7, 3))
    # Logique simplifiée du graphique de rétention
    plt.close()
    buf = io.BytesIO(); buf.seek(0)
    return buf

# Fonctions vides pour la compatibilité
def update_slide2(prs, agencies, group_stats): pass
def make_slide5_img(agencies, group_stats): return io.BytesIO()
def make_slide7_img(agencies): return io.BytesIO()

# --- FONCTION PRINCIPALE (APPELÉE PAR RENDER) ---
def generate_report(input_excel_path):
    global EXCEL
    EXCEL = input_excel_path
    
    # 1. Calculs
    agencies, group_stats, top_moves, df, market_detected = load_stats()
    
    # 2. Template
    template_path = os.path.join(os.path.dirname(__file__), 'T21_HK_Agencies_Glass_v12.pptx')
    prs = Presentation(template_path)
    
    # 3. Remplacement Texte
    replace_text_in_pptx(prs, "Hong Kong", market_detected)
    
    # 4. Design Moderne (Slides 3 & 4)
    try:
        from modern_design import build_slide3_modern, build_slide4_modern
        build_slide3_modern(prs.slides[2], top_moves)
        build_slide4_modern(prs.slides[3], agencies)
    except Exception as e:
        print(f"Erreur design moderne: {e}")

    # 5. Sauvegarde
    output_path = os.path.join("/tmp", f"NBB_{market_detected}_2025.pptx")
    prs.save(output_path)
    return output_path
