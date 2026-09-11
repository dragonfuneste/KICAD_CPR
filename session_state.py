import streamlit as st
import pandas as pd
from component_checker import build_report_rows, clean_BOM
from find_component import find_by_row, find_by_keyword
from Pricing import Update_Price_Stock
from Pricing import build_pricing_report  # ou le nom du fichier où tu as mis cette fonction

def afficher_matching_interactif(BOM, library):
    """
    Affiche le BOM ligne par ligne. Pour chaque ligne sans match fiable,
    propose un menu déroulant des candidats + une recherche manuelle.
    Retourne un dict {index_ligne_BOM: nom_du_composant_choisi}.
    """
    if "choix_matching" not in st.session_state:
        st.session_state.choix_matching = {}

    for index, row in BOM.iterrows():
        resultat = find_by_row(library, row)
        libelle = row.get("Manufacturer Ref") or row.get("Description") or f"Ligne {index + 1}"

        # cas 1 : match unique et fiable -> affiché en lecture seule, pas d'intervention
        if isinstance(resultat, pd.Series):
            st.session_state.choix_matching[index] = resultat["nom du composant"]
            st.markdown(f"✅ **Ligne {index + 1}** ({libelle}) → `{resultat['nom du composant']}`")
            continue

        # cas 2 : plusieurs candidats ou aucun -> on demande à l'utilisateur
        candidats = resultat["nom du composant"].tolist() if isinstance(resultat, pd.DataFrame) else []

        with st.expander(f"⚠️ Ligne {index + 1} — {libelle} (à valider)", expanded=True):
            options = ["-- à définir --"] + candidats + ["🔍 Recherche manuelle..."]
            choix = st.selectbox(
                "Composant à associer",
                options,
                key=f"select_{index}",
            )

            if choix == "🔍 Recherche manuelle...":
                terme = st.text_input("Rechercher dans la bibliothèque", key=f"recherche_{index}")
                if terme:
                    resultat_recherche = find_by_keyword(library, terme)

                    if isinstance(resultat_recherche, pd.DataFrame):
                        choix_manuel = st.selectbox(
                            "Résultats trouvés",
                            resultat_recherche["nom du composant"].tolist(),
                            key=f"manuel_{index}",
                        )
                        st.session_state.choix_matching[index] = choix_manuel
                    elif isinstance(resultat_recherche, pd.Series):
                        st.session_state.choix_matching[index] = resultat_recherche["nom du composant"]
                        st.info(f"Trouvé : {resultat_recherche['nom du composant']}")
                    else:
                        st.warning("Rien trouvé pour ce terme.")
                        st.session_state.choix_matching.pop(index, None)

            elif choix != "-- à définir --":
                st.session_state.choix_matching[index] = choix
            else:
                st.session_state.choix_matching.pop(index, None)

    return st.session_state.choix_matching