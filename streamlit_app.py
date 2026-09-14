"""
Application Streamlit : reprise du workflow du notebook (Notebook_back.ipynb)

- Upload BOM + Librairie de composants (.xlsx)
- Recherche automatique dans la lib (Manufacturer Ref / Footprint / Value)
- Remplissage du LCSC_part_number + détection des écarts (mismatch, Mouser only...)
- Mise à jour prix/stock (scraping LCSC) sur la librairie
- Rapport final : prix total pour n PCB, composants à vérifier
- Téléchargement des fichiers résultats
"""

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from backend.find_component import search_in_lib
from backend.BOM_function import fill_bom_result, check_bom
from backend.Pricing_lcsc import Update_Price_Stock
from backend.BOM_report import build_bom_report, export_report_excel


st.set_page_config(page_title="BOM Checker", page_icon="🔧", layout="wide")


# ============================================================
# Utilitaires
# ============================================================

def _save_uploaded_file(uploaded_file, workdir: Path) -> Path:
    """Sauvegarde un fichier uploadé sur disque (openpyxl a besoin d'un chemin)."""
    path = workdir / uploaded_file.name
    path.write_bytes(uploaded_file.getbuffer())
    return path


def _df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="BOM")
    return buffer.getvalue()


def _status_for_display(row: pd.Series) -> str:
    """Statut simple pour un aperçu rapide ligne par ligne (avant le rapport final)."""
    lcsc = row.get("LCSC_part_number")
    comments = str(row.get("Comments", "")) if pd.notna(row.get("Comments")) else ""

    if comments:
        return "⚠️ " + comments
    if pd.isna(lcsc) or str(lcsc).strip() == "":
        return "❌ Non détecté"
    return "✅ OK"


# ============================================================
# Etat de session
# ============================================================

if "workdir" not in st.session_state:
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="bom_checker_"))
if "bom" not in st.session_state:
    st.session_state.bom = None
if "lib_path" not in st.session_state:
    st.session_state.lib_path = None
if "lib" not in st.session_state:
    st.session_state.lib = None
if "prices_updated" not in st.session_state:
    st.session_state.prices_updated = False


st.title("🔧 BOM Checker")
st.caption("Recherche des composants dans la librairie, vérification des références, prix pour n PCB.")


# ============================================================
# 1. Upload des fichiers
# ============================================================

st.header("1. Fichiers")

col1, col2 = st.columns(2)

with col1:
    bom_file = st.file_uploader("BOM (.xlsx)", type=["xlsx"])
    header_row = st.number_input(
        "Ligne d'en-tête du BOM (0 = première ligne)",
        min_value=0, max_value=50, value=7,
        help="Nombre de lignes de titre/logo avant les en-têtes de colonnes. "
             "7 pour un export KiCad classique (comme Carte_radar_V0.2--BoM.xlsx).",
    )

with col2:
    lib_file = st.file_uploader("Librairie de composants (.xlsx)", type=["xlsx"])

if bom_file is not None and lib_file is not None:

    lib_path = _save_uploaded_file(lib_file, st.session_state.workdir)
    st.session_state.lib_path = lib_path

    try:
        bom_raw = pd.read_excel(bom_file, header=header_row).dropna(how="all")
        if "Row" in bom_raw.columns:
            bom_raw = bom_raw.drop(columns=["Row"])
        if "LCSC_part_number" not in bom_raw.columns:
            bom_raw["LCSC_part_number"] = pd.NA
        if "Comments" not in bom_raw.columns:
            bom_raw["Comments"] = pd.NA
        bom_raw["LCSC_part_number"] = bom_raw["LCSC_part_number"].astype("object")
        bom_raw["Comments"] = bom_raw["Comments"].astype("object")

        lib_df = pd.read_excel(lib_path)

    except Exception as e:
        st.error(f"Erreur de lecture des fichiers : {e}")
        st.stop()

    if st.session_state.bom is None:
        st.session_state.bom = bom_raw
    st.session_state.lib = lib_df

    st.success(f"BOM : {len(bom_raw)} lignes  |  Librairie : {len(lib_df)} composants")

else:
    st.info("Charge un BOM et une librairie pour commencer.")
    st.stop()


# ============================================================
# 2. Recherche + remplissage du LCSC
# ============================================================

st.header("2. Recherche dans la librairie")

if st.button("🔍 Lancer la recherche et remplir le BOM", type="primary"):

    bom = st.session_state.bom.copy()
    lib = st.session_state.lib

    progress = st.progress(0, text="Recherche en cours...")
    nb = len(bom)

    for i, (index, row) in enumerate(bom.iterrows()):
        result = search_in_lib(lib, row)
        row = fill_bom_result(row, result)
        row = check_bom(row, result)
        bom.loc[index] = row
        progress.progress((i + 1) / nb, text=f"Ligne {i + 1}/{nb}")

    progress.empty()
    st.session_state.bom = bom
    st.success("Recherche terminée.")

if st.session_state.bom is not None and "LCSC_part_number" in st.session_state.bom.columns:

    bom = st.session_state.bom

    apercu = bom.copy()
    apercu["Statut"] = apercu.apply(_status_for_display, axis=1)

    nb_ok = (apercu["Statut"] == "✅ OK").sum()
    nb_warn = apercu["Statut"].str.startswith("⚠️").sum()
    nb_ko = (apercu["Statut"] == "❌ Non détecté").sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("OK", nb_ok)
    c2.metric("À vérifier", nb_warn)
    c3.metric("Non détecté", nb_ko)

    cols_a_montrer = [c for c in [
        "References", "Manufacturer Ref", "Value", "Footprint",
        "LCSC_part_number", "Comments", "Statut",
    ] if c in apercu.columns]

    st.dataframe(apercu[cols_a_montrer], use_container_width=True, height=400)

    st.download_button(
        "⬇️ Télécharger le BOM (état actuel)",
        data=_df_to_excel_bytes(bom.drop(columns=["Statut"], errors="ignore")),
        file_name="bom_verifie.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================
# 3. Mise à jour prix / stock (scraping LCSC)
# ============================================================

st.header("3. Prix & stock")

st.caption(
    "Interroge LCSC.com pour chaque composant de la librairie (1 requête/seconde). "
    "Peut prendre plusieurs minutes selon la taille de la librairie."
)

if st.button("💰 Mettre à jour les prix et stocks"):
    with st.spinner("Récupération des prix sur LCSC..."):
        try:
            Update_Price_Stock(str(st.session_state.lib_path))
            st.session_state.lib = pd.read_excel(st.session_state.lib_path)
            st.session_state.prices_updated = True
            st.success("Prix et stocks mis à jour.")
        except Exception as e:
            st.error(f"Erreur pendant la mise à jour : {e}")


# ============================================================
# 4. Rapport final
# ============================================================

st.header("4. Rapport")

n_pcb = st.number_input("Nombre de PCB", min_value=1, value=10, step=1)

if not st.session_state.prices_updated:
    st.warning("Lance la mise à jour des prix (étape 3) avant de générer le rapport, sinon les prix seront manquants.")

if st.button("📊 Générer le rapport"):

    bom = st.session_state.bom
    lib = st.session_state.lib

    report = build_bom_report(bom, lib, n_pcb=n_pcb)

    detail = report["detail"]
    missing = report["missing"]
    total_price = report["total_price"]

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Prix total pour {n_pcb} PCB", f"{total_price:.2f} €")
    c2.metric("Lignes OK", (detail["Status"] == "OK").sum())
    c3.metric("Lignes à vérifier", len(missing))

    if not missing.empty:
        st.subheader("Composants à vérifier")
        cols = [c for c in [
            "Manufacturer Ref", "Value", "Footprint",
            "LCSC_part_number", "Status", "Comments",
        ] if c in missing.columns]
        st.dataframe(missing[cols], use_container_width=True)
    else:
        st.success("Tous les composants sont correctement identifiés.")

    with st.expander("Voir le détail complet"):
        st.dataframe(detail, use_container_width=True)

    # Export Excel du rapport
    report_path = st.session_state.workdir / "rapport_bom.xlsx"
    export_report_excel(report, str(report_path), n_pcb=n_pcb)

    st.download_button(
        "⬇️ Télécharger le rapport (.xlsx)",
        data=report_path.read_bytes(),
        file_name="rapport_bom.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )