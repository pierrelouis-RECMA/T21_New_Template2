import os, io, warnings, tempfile
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from modern_design import build_slide3_modern, build_slide4_modern

warnings.filterwarnings('ignore')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE_DIR, 'T21_HK_Agencies_Glass_v12.pptx')

# Dictionnaire de correspondance Agence -> Groupe (indispensable pour modern_design)
AGENCY_GROUP_MAP = {
    'SPARK FOUNDRY': 'Publicis Media', 'STARCOM': 'Publicis Media', 'ZENITH': 'Publicis Media',
    'PHD': 'Omnicom Media', 'OMD': 'Omnicom Media', 'HEARTS & SCIENCE': 'Omnicom Media',
    'DENTSU X': 'Dentsu', 'IPROSPECT': 'Dentsu', 'CARAT': 'Dentsu',
    'HAVAS MEDIA': 'Havas Media Network', 'ARENA': 'Havas Media Network',
    'ESSENCEMEDIACOM': 'WPP Media', 'WAVEMAKER': 'WPP Media', 'MINDSHARE': 'WPP Media',
    'INITIATIVE': 'IPG Mediabrands', 'UM': 'IPG Mediabrands',
    'VALE MEDIA': 'Independent'
}

def get_group(agency_name):
    return AGENCY_GROUP_MAP.get(agency_name.upper(), 'Independent')

def clean_slide(slide):
    for shape in list(slide.shapes):
        if shape.has_table or shape.has_text_frame:
            if shape.top > Inches(1.1): 
                sp = shape._element
                sp.getparent().remove(sp)

def replace_text_globally(prs, old_text, new_text):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)

def load_stats(file_path):
    df = pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    
    country_col = 'Country' if 'Country' in df.columns else 'Country of Decision'
    market_name = str(df[country_col].dropna().iloc[0]).upper() if country_col in df.columns else "MARKET"
    
    df['Agency'] = df['Agency'].astype(str).str.strip().str.upper()
    df['NewBiz'] = df['NewBiz'].astype(str).str.strip().str.upper()

    agencies = []
    unique_agencies = [a for a in df['Agency'].unique() if a not in ['NAN', '', 'NONE']]
    
    for ag in unique_agencies:
        sub = df[df['Agency'] == ag]
        w = pd.to_numeric(sub[sub['NewBiz']=='WIN']['Integrated Spends'], errors='coerce').sum()
        d = pd.to_numeric(sub[sub['NewBiz']=='DEPARTURE']['Integrated Spends'], errors='coerce').sum()
        
        agencies.append({
            'agency': ag,
            'group': get_group(ag), # AJOUT DE LA CLÉ 'group' MANQUANTE
            'nbb': w + d, 
            'wins': w, 
            'dep': d,
            'wc': len(sub[sub['NewBiz']=='WIN']),
            'dc': len(sub[sub['NewBiz']=='DEPARTURE']),
            'wins_rows': sub[sub['NewBiz']=='WIN'].to_dict('records'),
            'dep_rows': sub[sub['NewBiz']=='DEPARTURE'].to_dict('records')
        })
    
    agencies.sort(key=lambda x: x['nbb'], reverse=True)
    for i, a in enumerate(agencies): a['rank'] = i+1

    top_moves = df[df['NewBiz'].isin(['WIN','RETENTION'])].copy()
    top_moves['IS_abs'] = pd.to_numeric(top_moves['Integrated Spends'], errors='coerce').abs()
    top_moves = top_moves.sort_values('IS_abs', ascending=False).head(15)

    return agencies, top_moves, market_name

def generate_report(input_excel_path):
    agencies, top_moves, market_name = load_stats(input_excel_path)
    
    if not os.path.exists(TEMPLATE):
        raise FileNotFoundError(f"Template non trouvé : {TEMPLATE}")
        
    prs = Presentation(TEMPLATE)
    replace_text_globally(prs, "Hong Kong", market_name.capitalize())
    replace_text_globally(prs, "HONG KONG", market_name)

    # Nettoyage des slides pour éviter les superpositions
    for idx in [1, 2, 3, 4, 5]: 
        if idx < len(prs.slides):
            clean_slide(prs.slides[idx])

    # Injection du design (Slide 3 & 4)
    build_slide3_modern(prs.slides[2], top_moves, market=market_name)
    build_slide4_modern(prs.slides[3], agencies, market=market_name)

    out_name = f"NBB_Report_{market_name}.pptx"
    output_path = os.path.join(tempfile.gettempdir(), out_name)
    prs.save(output_path)
    return output_path
