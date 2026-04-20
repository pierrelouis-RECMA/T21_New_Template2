import os, warnings, tempfile
import pandas as pd
from pptx import Presentation
from pptx.util import Inches

warnings.filterwarnings('ignore')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE_DIR, 'T21_HK_Agencies_Glass_v12.pptx')

# Mapping des Groupes
AGENCY_GROUP_MAP = {
    'SPARK FOUNDRY': 'Publicis Media', 'STARCOM': 'Publicis Media', 'ZENITH': 'Publicis Media',
    'PHD': 'Omnicom Media', 'OMD': 'Omnicom Media', 'HEARTS & SCIENCE': 'Omnicom Media',
    'DENTSU X': 'Dentsu', 'IPROSPECT': 'Dentsu', 'CARAT': 'Dentsu',
    'HAVAS MEDIA': 'Havas Media Network', 'ARENA': 'Havas Media Network',
    'ESSENCEMEDIACOM': 'WPP Media', 'WAVEMAKER': 'WPP Media', 'MINDSHARE': 'WPP Media',
    'INITIATIVE': 'IPG Mediabrands', 'UM': 'IPG Mediabrands',
    'VALE MEDIA': 'Independent'
}

def get_group(ag_name):
    return AGENCY_GROUP_MAP.get(str(ag_name).upper(), 'Independent')

def fill_table_safely(slide, data_rows, mapping_keys):
    """Cherche un tableau et force l'écriture des données"""
    found = False
    for shape in slide.shapes:
        if shape.has_table:
            found = True
            table = shape.table
            print(f"--> Tableau trouvé sur la slide. Lignes: {len(table.rows)}")
            for i, row_data in enumerate(data_rows):
                if (i + 1) < len(table.rows):
                    cells = table.rows[i + 1].cells
                    for col_idx, key in enumerate(mapping_keys):
                        if col_idx < len(cells):
                            # On vide la cellule avant d'écrire
                            cells[col_idx].text_frame.clear()
                            val = row_data.get(key, "")
                            
                            # Formatage des nombres
                            if isinstance(val, (int, float)) and key in ['nbb', 'wins', 'dep', 'Integrated Spends']:
                                text_val = f"{val:+.1f}m$" if val != 0 else "0m$"
                            else:
                                text_val = str(val)
                            
                            # On écrit la nouvelle valeur
                            p = cells[col_idx].text_frame.paragraphs[0]
                            p.text = text_val
    if not found:
        print(" /!\\ AUCUN TABLEAU TROUVÉ SUR CETTE SLIDE")

def replace_text_globally(prs, old, new):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        if old in r.text:
                            r.text = r.text.replace(old, new)

def load_all_data(file_path):
    # Lecture multi-format (CSV ou Excel)
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)
    
    df.columns = [str(c).strip() for c in df.columns]
    df['Agency'] = df['Agency'].astype(str).str.strip().str.upper()
    df['NewBiz'] = df['NewBiz'].astype(str).str.strip().str.upper()
    df['Integrated Spends'] = pd.to_numeric(df['Integrated Spends'], errors='coerce').fillna(0)
    
    # 1. TOP Moves (Slide 3)
    top_moves = df[df['NewBiz'].isin(['WIN','RETENTION'])].sort_values('Integrated Spends', ascending=False).head(20).to_dict('records')
    for i, r in enumerate(top_moves): r['rank'] = i+1

    # 2. Agencies Stats (Slide 4)
    ag_list = []
    for ag in df['Agency'].unique():
        if ag in ['NAN', '', 'NONE']: continue
        sub = df[df['Agency'] == ag]
        w = sub[sub['NewBiz']=='WIN']['Integrated Spends'].sum()
        d = sub[sub['NewBiz']=='DEPARTURE']['Integrated Spends'].sum()
        ag_list.append({'rank': 0, 'agency': ag, 'nbb': w+d, 'wins': w, 'dep': d, 'group': get_group(ag)})
    
    ag_list.sort(key=lambda x: x['nbb'], reverse=True)
    for i, a in enumerate(ag_list): a['rank'] = i+1

    # 3. Group Stats (Slide 5)
    if ag_list:
        grp_df = pd.DataFrame(ag_list).groupby('group')['nbb'].sum().reset_index()
        grp_stats = grp_df.sort_values('nbb', ascending=False).to_dict('records')
        for i, g in enumerate(grp_stats): g['rank'] = i+1
    else:
        grp_stats = []

    market = "MEXICO"
    if 'Country' in df.columns:
        market = str(df['Country'].dropna().iloc[0]).upper()

    return top_moves, ag_list, grp_stats, market

def generate_report(input_excel_path):
    print(f"Début de la génération pour: {input_excel_path}")
    top_moves, ag_stats, grp_stats, market = load_all_data(input_excel_path)
    
    if not os.path.exists(TEMPLATE):
        raise FileNotFoundError(f"Template introuvable: {TEMPLATE}")
        
    prs = Presentation(TEMPLATE)
    
    # Update global
    replace_text_globally(prs, "Hong Kong", market.capitalize())
    replace_text_globally(prs, "HONG KONG", market)

    print("Mise à jour Slide 3...")
    fill_table_safely(prs.slides[2], top_moves, ['rank', 'Agency', 'Advertiser', 'Integrated Spends'])
    
    print("Mise à jour Slide 4...")
    fill_table_safely(prs.slides[3], ag_stats, ['rank', 'agency', 'nbb', 'wins', 'dep'])
    
    print("Mise à jour Slide 5...")
    fill_table_safely(prs.slides[4], grp_stats, ['rank', 'group', 'nbb'])

    out_name = f"NBB_Report_{market}.pptx"
    output_path = os.path.join(tempfile.gettempdir(), out_name)
    prs.save(output_path)
    print(f"Fichier sauvegardé: {output_path}")
    return output_path
