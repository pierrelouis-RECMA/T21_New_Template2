import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor

def generate_ppt(excel_path, output_ppt_path):
    # 1. Lire les données Excel (adapte le nom de la feuille si nécessaire)
    df = pd.read_excel(excel_path, sheet_name="Feuil1")

    # 2. Filtrer et trier les données pour les groupes (slide 3)
    groups_df = df[df["Type"] == "Group"].sort_values(by="New Biz Balance (€m)", ascending=False)

    # 3. Charger le template PowerPoint
    prs = Presentation("templates/T21_HK_Agencies_Glass_v12.pptx")
    slide = prs.slides[2]  # Slide 3: "New Biz Balance overview by Group"
    table = slide.shapes[0].table  # Supposons que le tableau est la première forme

    # 4. Remplir les en-têtes
    headers = ["Rank", "Group", "New Biz Balance 2025 (€m)", "Wins (€m)", "Departures (€m)"]
    for col, header in enumerate(headers):
        table.cell(0, col).text = header

    # 5. Remplir les lignes avec les données
    for i, row in groups_df.iterrows():
        table.cell(i+1, 0).text = str(i+1)  # Rank
        table.cell(i+1, 1).text = row["Group"]
        table.cell(i+1, 2).text = str(row["New Biz Balance (€m)"])
        table.cell(i+1, 3).text = str(row["Wins (€m)"])
        table.cell(i+1, 4).text = str(row["Departures (€m)"])

        # Mise en forme conditionnelle
        cell = table.cell(i+1, 2)
        if row["New Biz Balance (€m)"] > 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0, 255, 0)  # Vert
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 0, 0)  # Rouge

    # 6. Sauvegarder le PPT généré
    prs.save(output_ppt_path)
