import streamlit as st
import pandas as pd

from find_component import find_by_row, extraire_empreinte  # si extraire_empreinte existe encore dans ton module


def _fusion_bom(row):
    """Fusionne Description + Value + empreinte d'une ligne BOM."""
    morceaux = []
    if pd.notna(row.get("Description")):
        morceaux.append(str(row["Description"]))
    if pd.notna(row.get("Value")):
        morceaux.append(str(row["Value"]))
    if pd.notna(row.get("empreinte")):
        morceaux.append(str(row["empreinte"]))
    return " | ".join(morceaux)

def _fusion_lib(lib_row):
    """Fusionne nom du composant + type + specification + empreinte d'une ligne bibliothèque."""
    morceaux = []
    for champ in ["nom du composant", "type", "specification", "empreinte"]:
        val = lib_row.get(champ)
        if pd.notna(val):
            morceaux.append(str(val))
    return " | ".join(morceaux)


def afficher_matching_tableau(BOM, library):
    """
    Tableau de matching : colonne 'Infos BOM' (Description+Value+empreinte fusionnés),
    et colonne 'Composant choisi' avec les candidats affichés en version fusionnée
    (nom+type+specification+empreinte) pour être lisible sans deviner.
    """
    label_vers_nom = {}
    for _, lib_row in library.iterrows():
        if pd.isna(lib_row.get("nom du composant")):
            continue
        label = _fusion_lib(lib_row)
        label_vers_nom[label] = lib_row["nom du composant"]

    tous_les_labels = sorted(label_vers_nom.keys())

    lignes = []
    for index, row in BOM.iterrows():
        resultat = find_by_row(library, row)

        if isinstance(resultat, pd.Series):
            statut = "trouvé"
            label_defaut = _fusion_lib(resultat)
        elif isinstance(resultat, pd.DataFrame):
            statut = "plusieurs candidats"
            label_defaut = _fusion_lib(resultat.iloc[0])
        else:
            statut = "aucune correspondance"
            label_defaut = None

        lignes.append({
            "Ligne": index + 1,
            "📋 Infos BOM (Description | Value | Empreinte)": _fusion_bom(row),
            "Statut": statut,
            "🎯 Composant choisi (nom | type | spec | empreinte)": label_defaut,
        })

    df = pd.DataFrame(lignes)

    df_edite = st.data_editor(
        df,
        column_config={
            "🎯 Composant choisi (nom | type | spec | empreinte)": st.column_config.SelectboxColumn(
                "🎯 Composant choisi (nom | type | spec | empreinte)",
                help="Tape pour rechercher dans toute la bibliothèque",
                options=tous_les_labels,
                required=False,
                width="large",
            ),
            "📋 Infos BOM (Description | Value | Empreinte)": st.column_config.TextColumn(
                "📋 Infos BOM (Description | Value | Empreinte)", disabled=True, width="large"
            ),
            "Statut": st.column_config.TextColumn("Statut", disabled=True),
        },
        disabled=["Ligne"],
        use_container_width=True,
        hide_index=True,
        key="editeur_matching",
    )

    # on reconvertit les libellés choisis en vrais noms de composant pour la suite (pricing, etc.)
    df_edite["Composant (nom réel)"] = df_edite[
        "🎯 Composant choisi (nom | type | spec | empreinte)"
    ].map(label_vers_nom)

    return df_edite