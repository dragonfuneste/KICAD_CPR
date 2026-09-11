import streamlit as st
import pandas as pd

from component_checker import build_report_rows, clean_BOM
from Pricing import Update_Price_Stock
from Pricing import build_pricing_report  # ou le nom du fichier où tu as mis cette fonction
from session_state import afficher_matching_tableau
st.set_page_config(page_title="BOM Checker", layout="wide")
st.title("Vérification BOM / Bibliothèque de composants")

# --------------------------------------------------------
# Sidebar : chemins des fichiers (ou upload)
# --------------------------------------------------------
with st.sidebar:
    st.header("Fichiers")
    chemin_lib = st.text_input("Chemin bibliothèque", "lib/Liste des composants.xlsx")
    nom_feuille = st.text_input("Nom de la feuille", "Feuille 1")
    fichier_bom = st.file_uploader("BOM à analyser (.xlsx)", type=["xlsx"])
    n_cartes = st.number_input("Nombre de cartes à produire", min_value=1, value=5)

onglet_maj, onglet_matching, onglet_pricing = st.tabs([
    "Mise à jour prix/stock", "Matching BOM", "Pricing & disponibilité"
])

with onglet_maj:
    st.write("Interroge LCSC pour mettre à jour Price/Stock dans la bibliothèque.")
    if st.button("Lancer la mise à jour"):
        with st.spinner("Interrogation de LCSC en cours..."):
            Update_Price_Stock(chemin_lib, nom_feuille)
        st.success("Bibliothèque mise à jour.")



with onglet_pricing:
    if fichier_bom is not None and st.button("Calculer le pricing"):
        with st.spinner("Calcul en cours..."):
            report, cout_total, manquants, repartition = build_pricing_report(
                fichier_bom, chemin_lib, n=n_cartes, sheet_name_lib=nom_feuille
            )

        st.metric(f"Coût total pour {n_cartes} carte(s)", f"{cout_total:.2f} EUR")

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Répartition par famille")
            st.dataframe(repartition, use_container_width=True)
            st.bar_chart(repartition.set_index("Famille")["% du coût total"])
        with col2:
            st.subheader("Composants manquants")
            if manquants.empty:
                st.success("Tout est disponible.")
            else:
                st.warning(f"{len(manquants)} ligne(s) à problème")
                st.dataframe(manquants, use_container_width=True)

        st.subheader("Détail complet")
        st.dataframe(report, use_container_width=True)


with onglet_matching:
    if fichier_bom is not None:
        BOM = clean_BOM(fichier_bom)
        library = pd.read_excel(chemin_lib, sheet_name=nom_feuille)

        df_matching = afficher_matching_tableau(BOM, library)

        if st.button("Valider les choix et passer au pricing"):
            st.session_state.bom_valide = BOM
            st.session_state.choix_valides = dict(zip(df_matching["Ligne"] - 1, df_matching["Composant choisi"]))
            st.success(f"{df_matching['Composant choisi'].notna().sum()}/{len(BOM)} lignes associées.")