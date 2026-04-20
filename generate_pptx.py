import os, warnings, tempfile
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from modern_design import build_slide3_modern, build_slide4_modern

warnings.filterwarnings('ignore')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE_DIR, 'T21_HK_Agencies_Glass_v12.pptx')

# Mapping des Groupes indispensable pour le design
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

def clean_data_only(slide):
    """
    Supprime uniquement les tableaux et les zones de texte de données,
    mais GARDE les cadres, le header et les liens de navigation.
    """
    for shape in list(slide.shapes):
        # On ne touche pas aux éléments du header (souvent en haut de la slide)
        # On ne touche pas aux rectangles/formes qui servent de "cadres" (shape_type 1 ou 6)
        if shape.top > Inches(1.2): # Tout ce qui est en dessous du header
            if shape.has_table: # On supprime les anciens tableaux de données
                sp = shape._element
                sp.getparent().remove(sp)
            elif shape.has_text_frame:
                # On ne supprime le texte que s'il contient des données variables 
                # (on évite de supprimer les titres de colonnes s'ils sont fixes)
                if any(word in shape.text for word in ['Nestlé', 'Disneyland', 'HK', 'Disney']):
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

def load_data(file_path):
    df = pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    country_col = 'Country' if 'Country' in df.columns else 'Country of Decision'
    market = str(df[country_col].dropna().iloc[0]).upper() if country_col in df.columns else "MEXICO"
    
    df['Agency'] = df['Agency'].astype(str).str.strip().str.upper()
    df['NewBiz'] = df['NewBiz'].astype(str).str.strip().str.upper()
    df['Integrated Spends'] = pd.to_numeric(df['Integrated Spends'], errors='coerce').fillna(0)

    agencies = []
    for ag in df['Agency'].unique():
        if ag in ['NAN', '', 'NONE']: continue
        sub = df[df['Agency'] == ag]
        w = sub[sub['NewBiz']=='WIN']['Integrated Spends'].sum()
        d = sub[sub['NewBiz']=='DEPARTURE']['Integrated Spends'].sum()
        agencies.append({
            'agency': ag, 'group': get_group(ag),
            'nbb': w + d, 'wins': w, 'dep': d,
            'wc': len(sub[sub['NewBiz']=='WIN']),
            'dc': len(sub[sub['NewBiz']=='DEPARTURE']),
            'wins_rows': sub[sub['NewBiz']=='WIN'].to_dict('records'),
            'dep_rows': sub[sub['NewBiz']=='DEPARTURE'].to_dict('records')
        })
    agencies.sort(key=lambda x: x['nbb'], reverse=True)
    for i, a in enumerate(agencies): a['rank'] = i+1
    
    top_moves = df[df['NewBiz'].isin(['WIN','RETENTION'])].copy()
    top_moves['IS_abs'] = top_moves['Integrated Spends'].abs()
    top_moves = top_moves.sort_values('IS_abs', ascending=False).head(15)
    return agencies, top_moves, market

def generate_report(input_excel_path):
    agencies, top_moves, market_name = load_data(input_excel_path)
    prs = Presentation(TEMPLATE)
    
    # 1. Remplacement des textes (Hong Kong -> Mexico) dans les titres
    replace_text_globally(prs, "Hong Kong", market_name.capitalize())
    replace_text_globally(prs, "HONG KONG", market_name)

    # 2. Mise à jour des slides
    for i, slide in enumerate(prs.slides):
        # On nettoie uniquement les données sur les slides 3, 4, 5, 6
        if 2 <= i <= 5:
            clean_data_only(slide)
            
        if i == 2: # Slide TOP MOVES
            build_slide3_modern(slide, top_moves, market=market_name)
        if i == 3: # Slide AGENCIES OVERVIEW
            build_slide4_modern(slide, agencies, market=market_name)

    out_name = f"NBB_Report_{market_name}.pptx"
    output_path = os.path.join(tempfile.gettempdir(), out_name)
    prs.save(output_path)
    return output_path
