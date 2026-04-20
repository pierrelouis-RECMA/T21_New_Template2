"""
NBB Report – PPTX Generator (Version Multi-Pays & Render)
Usage : Automatiquement appelé par render_app.py
"""

import copy, io, os, shutil, warnings, tempfile
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib import rcParams
rcParams['font.family'] = 'DejaVu Sans'

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

warnings.filterwarnings('ignore')

# ── Paths (Configuration pour Render) ────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Le template doit être à la racine de votre GitHub
TEMPLATE  = os.path.join(BASE_DIR, 'T21_HK_Agencies_Glass_v12.pptx')
# Chemins temporaires pour le traitement sur Render
IMG_DIR   = os.path.join(tempfile.gettempdir(), 'slide_imgs')
os.makedirs(IMG_DIR, exist_ok=True)

# ── Palette & Groupes ────────────────────────────────────────────────────────
C_HEADER   = '#2D5C54'
C_WIN      = '#4CAF50'
C_DEP      = '#E53935'
C_RET      = '#FF9800'
C_ALT      = '#F2F6F5'
C_TOTAL    = '#E8F5E9'
C_NBB_POS  = '#1B5E20'
C_NBB_NEG  = '#B71C1C'

GROUP_COLORS = {
    'Publicis Media':      '#FFEADD',
    'Omnicom Media':       '#FFE699',
    'Dentsu':              '#E2F0D9',
    'Havas Media Network': '#DCB9FF',
    'WPP Media':           '#FFE4FF',
    'IPG Mediabrands':     '#D9EAD3',
    'Independent':         '#FFFFFF', # Fond Blanc
}

AGENCY_GROUP = {
    'SPARK FOUNDRY': 'Publicis Media', 'STARCOM': 'Publicis Media', 'ZENITH': 'Publicis Media',
    'PHD': 'Omnicom Media', 'OMD': 'Omnicom Media', 'HEARTS & SCIENCE': 'Omnicom Media',
    'DENTSU X': 'Dentsu', 'IPROSPECT': 'Dentsu', 'CARAT': 'Dentsu',
    'HAVAS MEDIA': 'Havas Media Network', 'ARENA': 'Havas Media Network',
    'ESSENCEMEDIACOM': 'WPP Media', 'WAVEMAKER': 'WPP Media', 'MINDSHARE': 'WPP Media',
    'INITIATIVE': 'IPG Mediabrands', 'UM': 'IPG Mediabrands',
    'VALE MEDIA': 'Independent'
}

GROUP_ORDER = ['Publicis Media','Omnicom Media','Dentsu','Havas Media Network','WPP Media', 'IPG Mediabrands', 'Independent']

def get_agency_group(agency_name):
    clean_name = str(agency_name).upper().strip()
    return AGENCY_GROUP.get(clean_name, 'Independent')

# ── Fonctions de Traitement ──────────────────────────────────────────────────
def load_stats(file_path):
    xl = pd.ExcelFile(file_path)
    df = pd.read_excel(file_path, sheet_name=xl.sheet_names[0])
    
    df['Agency'] = df['Agency'].astype(str).str.strip().str.upper()
    df['NewBiz'] = df['NewBiz'].astype(str).str.strip().str.upper()
    df['Group']  = df['Agency'].apply(get_agency_group)

    # Détection Pays
    country_col = 'Country' if 'Country' in df.columns else 'Country of Decision'
    market_name = str(df[country_col].dropna().iloc[0]) if country_col in df.columns else "Market"

    def fmt_date(v):
        try: return pd.to_datetime(v).strftime('%b-%y')
        except: return ''
    df['Date_str'] = df['Date of announcement'].apply(fmt_date)

    agencies = []
    for ag in df.Agency.unique():
        if ag == 'NAN': continue
        sub = df[df.Agency == ag]
        w = sub[sub.NewBiz=='WIN']['Integrated Spends'].sum()
        d = sub[sub.NewBiz=='DEPARTURE']['Integrated Spends'].sum()
        r = sub[sub.NewBiz=='RETENTION']['Integrated Spends'].sum()
        agencies.append({
            'agency': ag, 'group': get_agency_group(ag),
            'nbb': w+d, 'wins': w, 'dep': d, 'ret': r,
            'wc': len(sub[sub.NewBiz=='WIN']),
            'dc': len(sub[sub.NewBiz=='DEPARTURE']),
            'wins_rows': sub[sub.NewBiz=='WIN'].to_dict('records'),
            'dep_rows':  sub[sub.NewBiz=='DEPARTURE'].to_dict('records'),
        })
    
    agencies.sort(key=lambda x: -x['nbb'])
    for i, a in enumerate(agencies): a['rank'] = i+1
    
    group_stats = {g: {'group':g, 'nbb':0, 'wins':0, 'dep':0, 'wc':0, 'dc':0, 'agencies':[]} for g in GROUP_ORDER}
    for a in agencies:
        g = a['group']
        if g in group_stats:
            group_stats[g]['nbb'] += a['nbb']
            group_stats[g]['wins'] += a['wins']
            group_stats[g]['dep'] += a['dep']
            group_stats[g]['wc'] += a['wc']
            group_stats[g]['dc'] += a['dc']
            group_stats[g]['agencies'].append(a)

    top_moves = df[df.NewBiz.isin(['WIN','RETENTION'])].copy()
    top_moves['IS_abs'] = top_moves['Integrated Spends'].abs()
    top_moves = top_moves.sort_values('IS_abs', ascending=False).head(20)

    return agencies, group_stats, top_moves, market_name

def replace_text_in_pptx(prs, old_text, new_text):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)

# ── Helpers Graphiques (simplifiés) ──────────────────────────────────────────
def hex2rgb(h): h=h.lstrip('#'); return tuple(int(h[i:i+2],16)/255 for i in (0,2,4))

def make_table_img(col_labels, col_widths, rows, row_styles, title=''):
    n_rows = len(rows) + 1
    fig_h = 0.4 + n_rows * 0.25
    fig, ax = plt.subplots(figsize=(10, fig_h))
    ax.set_axis_off()
    
    # Simple table logic pour les slides 5, 6, 7
    table = ax.table(cellText=rows, colLabels=col_labels, loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close()
    buf.seek(0)
    return buf

# ── Main Function appelée par l'App ──────────────────────────────────────────
def generate_report(input_excel_path):
    agencies, group_stats, top_moves, market_name = load_stats(input_excel_path)
    
    if not os.path.exists(TEMPLATE):
        raise FileNotFoundError(f"Template non trouvé : {TEMPLATE}")
        
    prs = Presentation(TEMPLATE)
    replace_text_in_pptx(prs, "Hong Kong", market_name)
    
    # Import des designs modernes pour Slides 3 & 4
    try:
        from modern_design import build_slide3_modern, build_slide4_modern
        build_slide3_modern(prs.slides[2], top_moves, market=market_name)
        build_slide4_modern(prs.slides[3], agencies, market=market_name)
    except Exception as e:
        print(f"Erreur design moderne: {e}")

    output_path = os.path.join(tempfile.gettempdir(), f"NBB_Report_{market_name}.pptx")
    prs.save(output_path)
    return output_path
