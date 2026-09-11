


from find_component import find_by_row
import pandas as pd


def clean_BOM(bom_path):
    """ Fonction qui nettoie le BOM en supprimant les lignes vides et en renommant les colonnes. """
    bom = pd.read_excel(bom_path, header=7).dropna(how="all").drop(columns=['Row'])
    return bom


def build_report_rows(BOM_name, library_name):
    """Construit un DataFrame résumant le résultat de matching pour chaque ligne du BOM."""
    BOM = clean_BOM(BOM_name)
    library = pd.read_excel(library_name)
    lignes = []
    for index, row in BOM.iterrows():
        resultat = find_by_row(library, row)

        if resultat is None:
            statut = "aucune correspondance"
            composants = ""
            score = ""
        elif isinstance(resultat, pd.DataFrame):
            statut = "plusieurs candidats"
            composants = ", ".join(resultat["nom du composant"].tolist())
            score = ", ".join(f"{s:.2f}" for s in resultat["__score"].tolist())
        else:  # pd.Series -> match unique
            statut = "trouvé"
            composants = resultat["nom du composant"]
            score = f"{resultat['__score']:.2f}"

        lignes.append({
            "Ligne BOM": index + 1,
            "References": row.get("References", ""),
            "Description": row.get("Description", ""),
            "Value": row.get("Value", ""),
            "Manufacturer Ref": row.get("Manufacturer Ref", ""),
            "Statut": statut,
            "Composant(s) bibliothèque": composants,
            "Score": score,
        })

    return pd.DataFrame(lignes)



