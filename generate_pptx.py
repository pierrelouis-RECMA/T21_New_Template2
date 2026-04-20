import os, re, pandas as pd
import xml.etree.ElementTree as ET
# On importe les fonctions de dessin de ton script original (ou on les inclut ici)

def load_mexico_data(file_path):
    """Transforme l'Excel Mexique en dictionnaire AGENCIES pour ton script"""
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Nettoyage
    df['Agency'] = df['Agency'].astype(str).str.strip().str.upper()
    df['NewBiz'] = df['NewBiz'].astype(str).str.strip().str.upper()
    df['Integrated Spends'] = pd.to_numeric(df['Integrated Spends'], errors='coerce').fillna(0)
    df['Advertiser'] = df['Advertiser'].astype(str).fillna("Unknown")

    # Groupement par Agence
    mexico_agencies = []
    for ag_name in df['Agency'].unique():
        if ag_name in ['NAN', '']: continue
        sub = df[df['Agency'] == ag_name]
        
        # Calcul NBB
        w_sum = sub[sub['NewBiz']=='WIN']['Integrated Spends'].sum()
        d_sum = sub[sub['NewBiz']=='DEPARTURE']['Integrated Spends'].sum()
        total_nbb = w_sum + d_sum
        
        # Extraction des listes (Wins, Deps, Rets)
        wins = [(r['Advertiser'], f"+{r['Integrated Spends']:.1f}m") 
                for _, r in sub[sub['NewBiz']=='WIN'].iterrows()]
        deps = [(r['Advertiser'], f"{r['Integrated Spends']:.1f}m") 
                for _, r in sub[sub['NewBiz']=='DEPARTURE'].iterrows()]
        rets = [(r['Advertiser'], "") 
                for _, r in sub[sub['NewBiz']=='RETENTION'].iterrows()]

        mexico_agencies.append({
            "name": ag_name,
            "group": "MEXICO", # On peut mapper les vrais groupes ici
            "nbb": f"{total_nbb:+.1f}m$",
            "wins": wins,
            "deps": deps,
            "rets": rets,
            "val_num": total_nbb
        })

    # Tri par NBB et répartition par Slide (4 agences max par slide comme ton script)
    mexico_agencies.sort(key=lambda x: x['val_num'], reverse=True)
    
    # Découpage par paquets de 4 pour les slides 22, 23, 24...
    final_dict = {}
    for i in range(0, len(mexico_agencies), 4):
        slide_idx = 22 + (i // 4)
        final_dict[slide_idx] = mexico_agencies[i:i+4]
    
    return final_dict

# Ensuite, dans ta fonction principale :
# AGENCIES = load_mexico_data("ton_fichier_mexique.xlsx")
# Et le reste de ton code générera les slides XML parfaitement.
