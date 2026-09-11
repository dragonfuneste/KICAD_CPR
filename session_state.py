import streamlit as st
import pandas as pd
from component_checker import build_report_rows, clean_BOM
from find_component import find_by_row, find_by_keyword
from Pricing import Update_Price_Stock
from Pricing import build_pricing_report  # ou le nom du fichier où tu as mis cette fonction

import streamlit as st
import pandas as pd

from find_component import find_by_row


def afficher_matching_tableau(BOM, library):
    """
    Affiche le matching du BOM sous forme de tableau éditable :
    une ligne par composant du BOM, les candidats proposés en texte,
    et une colonne déroulante pour choisir/corriger le composant final.

    Retourne le DataFrame édité (avec la colonne 'Composant choisi' à jour).
    """
    tous_les_composants = sorted(library["nom du composant"].dropna().unique().tolist())

    lignes = []
    for index, row in BOM.iterrows():
        resultat = find_by_row(library, row)

        if isinstance(resultat, pd.Series):
            statut = "trouvé"
            candidats_txt = resultat["nom du composant"]
            choix_par_defaut = resultat["nom du composant"]
        elif isinstance(resultat, pd.DataFrame):
            statut = "plusieurs candidats"
            candidats_txt = " | ".join(resultat["nom du composant"].tolist())
            choix_par_defaut = resultat.iloc[0]["nom du composant"]
        else:
            statut = "aucune correspondance"
            candidats_txt = ""
            choix_par_defaut = None

        lignes.append({
            "Ligne": index + 1,
            "Manufacturer Ref": row.get("Manufacturer Ref", ""),
            "Description": row.get("Description", ""),
            "Statut": statut,
            "Candidats proposés": candidats_txt,
            "Composant choisi": choix_par_defaut,
        })

    df = pd.DataFrame(lignes)

    df_edite = st.data_editor(
        df,
        column_config={
            "Composant choisi": st.column_config.SelectboxColumn(
                "Composant choisi",
                help="Tape pour rechercher dans toute la bibliothèque",
                options=tous_les_composants,
                required=False,
            ),
            "Statut": st.column_config.TextColumn("Statut", disabled=True),
        },
        disabled=["Ligne", "Manufacturer Ref", "Description", "Candidats proposés"],
        use_container_width=True,
        hide_index=True,
        key="editeur_matching",
    )

    return df_edite