import os, warnings, tempfile
import pandas as pd
from pptx import Presentation
from pptx.util import Inches

warnings.filterwarnings('ignore')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE_DIR, 'T21_HK_Agencies_Glass_v12.pptx')

# Mapping des Groupes RECMA
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
    return AGENCY_GROUP_MAP.get(ag_name.upper(), 'Independent')

def fill_table(slide, data_rows, mapping_cols, start_row=1):
    """Remplit un tableau existant sur une slide"""
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            for i, row_data in enumerate(data_rows):
                if (i + start_row) < len(table.rows):
                    cells = table.rows[i + start_row].cells
                    for col_idx, key in enumerate(mapping_cols):
                        if col_idx < len(cells):
                            val = row_data.get(key, "")
                            if isinstance(val, (int, float)) and key in ['nbb', 'wins', 'dep', 'Integrated Spends', 'value']:
                                cells[col_idx].text = f"{val:+.1f}m$" if val < 0 or key == 'nbb' else f"{val:.1f}m$"
                            else:
                                cells[col_idx].text = str(val)
            return True
    return False

def replace_text_globally(prs, old, new):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        if old in r.text: r.text = r.text.replace(old, new)

def load_all_data(file_path):
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
        if ag in ['NAN', '']: continue
        sub = df[df['Agency']==ag]
        w = sub[sub['NewBiz']=='WIN']['Integrated Spends'].sum()
        d = sub[sub['NewBiz']=='DEPARTURE']['Integrated Spends'].sum()
        ag_list.append({'agency': ag, 'nbb': w+d, 'wins': w, 'dep': d, 'group': get_group(ag)})
    ag_list.sort(key=lambda x: x['nbb'], reverse=True)
    for i, a in enumerate(ag_list): a['rank'] = i+1

    # 3. Group Stats (Slide 5)
    grp_df = pd.DataFrame(ag_list).groupby('group')['nbb'].sum().reset_index()
    grp_stats = grp_df.sort_values('nbb', ascending=False).to_dict('records')
    for i, g in enumerate(grp_stats): g['rank'] = i+1

    market = str(df['Country'].iloc[0]).upper() if 'Country' in df.columns else "MEXICO"
    return top_moves, ag_list, grp_stats, market

def generate_report(input_excel_path):
    top_moves, ag_stats, grp_stats, market = load_all_data(input_excel_path)
    prs = Presentation(TEMPLATE)
    
    # Mise à jour globale
    replace_text_globally(prs, "Hong Kong", market.capitalize())
    replace_text_globally(prs, "HONG KONG", market)

    # Remplissage des tableaux
    fill_table(prs.slides[2], top_moves, ['rank', 'Agency', 'Advertiser', 'Integrated Spends']) # Slide 3
    fill_table(prs.slides[3], ag_stats, ['rank', 'agency', 'nbb', 'wins', 'dep'])              # Slide 4
    fill_table(prs.slides[4], grp_stats, ['rank', 'group', 'nbb'])                            # Slide 5

    out_name = f"NBB_Report_{market}.pptx"
    output_path = os.path.join(tempfile.gettempdir(), out_name)
    prs.save(output_path)
    return output_path
