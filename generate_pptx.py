import os, warnings, tempfile
import pandas as pd
from pptx import Presentation
from pptx.util import Inches

warnings.filterwarnings('ignore')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE_DIR, 'T21_HK_Agencies_Glass_v12.pptx')

def fill_existing_table(slide, data_rows, start_row=1):
    """Trouve le premier tableau de la slide et le remplit"""
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            # On parcourt les données
            for i, row_data in enumerate(data_rows):
                if (i + start_row) < len(table.rows):
                    # Exemple pour la slide TOP Moves :
                    # Col 0: Rank, Col 1: Agency, Col 2: Advertiser, Col 3: Spend
                    cells = table.rows[i + start_row].cells
                    cells[0].text = str(i + 1)
                    cells[1].text = str(row_data.get('Agency', ''))
                    cells[2].text = str(row_data.get('Advertiser', ''))
                    val = row_data.get('Integrated Spends', 0)
                    cells[3].text = f"{val:+.1f}m$"
            return True
    return False

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
    
    # Top Moves pour Slide 3
    top_moves = df[df['NewBiz'].isin(['WIN','RETENTION'])].copy()
    top_moves = top_moves.sort_values('Integrated Spends', ascending=False).head(15).to_dict('records')
    
    return top_moves, market

def generate_report(input_excel_path):
    top_moves, market_name = load_data(input_excel_path)
    prs = Presentation(TEMPLATE)
    
    # 1. Update Pays
    replace_text_globally(prs, "Hong Kong", market_name.capitalize())
    replace_text_globally(prs, "HONG KONG", market_name)

    # 2. Remplissage Slide 3 (TOP moves)
    # On n'appelle PLUS modern_design. On remplit le tableau du template.
    fill_existing_table(prs.slides[2], top_moves, start_row=1)

    out_name = f"NBB_Report_{market_name}.pptx"
    output_path = os.path.join(tempfile.gettempdir(), out_name)
    prs.save(output_path)
    return output_path
