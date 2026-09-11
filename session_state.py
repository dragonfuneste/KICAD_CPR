import streamlit as st
import pandas as pd
from find_component import find_by_row

def _fusion_bom(row):
    morceaux = []
    if pd.notna(row.get("Description")):
        morceaux.append(str(row["Description"]))
    if pd.notna(row.get("Value")):
        morceaux.append(str(row["Value"]))
    if pd.notna(row.get("empreinte")):
        morceaux.append(str(row["empreinte"]))
    return " | ".join(morceaux)

def _fusion_lib(lib_row):
    morceaux = []
    for champ in ["nom du composant", "type", "specification", "empreinte"]:
        val = lib_row.get(champ)
        if pd.notna(val):
            morceaux.append(str(val))
    return " | ".join(morceaux)

def afficher_matching_tableau(BOM, library):
    label_vers_nom = {}
    for _, lib_row in library.iterrows():
        if pd.isna(lib_row.get("nom du composant")):
            continue
        label = _fusion_lib(lib_row)
        label_vers_nom[label] = lib_row["nom du composant"]

    tous_les_labels = sorted(label_vers_nom.keys())

    # Initialisation de l'état persistant dans st.session_state si besoin
    if "df_matching_etat" not in st.session_state:
        lignes = []
        for index, row in BOM.iterrows():
            resultat = find_by_row(library, row)
            if isinstance(resultat, pd.Series):
                label_defaut = _fusion_lib(resultat)
            elif isinstance(resultat, pd.DataFrame):
                label_defaut = _fusion_lib(resultat.iloc[0])
            else:
                label_defaut = None

            lignes.append({
                "Ligne": index + 1,
                "🔒 Verrouillé": False,
                "📋 Infos BOM (Description | Value | Empreinte)": _fusion_bom(row),
                "Statut": "trouvé" if label_defaut else "aucune correspondance",
                "🎯 Composant choisi": label_defaut,
            })
        st.session_state.df_matching_etat = pd.DataFrame(lignes)

    df_actuel = st.session_state.df_matching_etat

    # Construction dynamique de la configuration des colonnes selon l'état de la case "Verrouillé"
    column_config = {
        "🔒 Verrouillé": st.column_config.CheckboxColumn(
            "🔒 Verrouillé",
            help="Coche pour figer et bloquer le choix de ce composant",
            default=False,
        ),
        "📋 Infos BOM (Description | Value | Empreinte)": st.column_config.TextColumn(
            disabled=True, width="large"
        ),
        "Statut": st.column_config.TextColumn(disabled=True),
    }

    # On définit dynamiquement si la colonne "Composant choisi" est un Selectbox ou un Texte grisé (désactivé) ligne par ligne
    # Note : st.data_editor applique une config par colonne entière. 
    # Pour simuler le blocage parfait : si la ligne est verrouillée, on bascule la valeur dans un format textuel figé, 
    # ou on empêche l'interaction via la logique de ré-attribution en arrière-plan.
    
    column_config["🎯 Composant choisi"] = st.column_config.SelectboxColumn(
        "🎯 Composant choisi",
        help="Tape pour rechercher dans la bibliothèque",
        options=tous_les_labels,
        required=False,
        width="large",
    )

    df_edite = st.data_editor(
        df_actuel,
        column_config=column_config,
        disabled=["Ligne", "📋 Infos BOM (Description | Value | Empreinte)", "Statut"],
        use_container_width=True,
        hide_index=True,
        key="editeur_matching_verrou",
    )

    # Sécurité anti-modification : si la case "Verrouillé" était et reste cochée, 
    # on force la valeur à ne pas bouger (bloque instantanément toute modification sur la ligne)
    for i in range(len(df_edite)):
        if df_actuel.loc[i, "🔒 Verrouillé"] and df_edite.loc[i, "🔒 Verrouillé"]:
            df_edite.loc[i, "🎯 Composant choisi"] = df_actuel.loc[i, "🎯 Composant choisi"]

    # Sauvegarde du nouvel état dans la session
    st.session_state.df_matching_etat = df_edite

    # Traduction en nom réel pour la suite
    df_edite["Composant (nom réel)"] = df_edite["🎯 Composant choisi"].map(label_vers_nom)

    return df_edite