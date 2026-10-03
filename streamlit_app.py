"""
Application Streamlit : reprise du workflow du notebook (Notebook_back.ipynb)

- Upload BOM + Librairie de composants (.xlsx)
- Recherche automatique dans la lib (Manufacturer Ref / Footprint / Value)
- Remplissage du LCSC_part_number + détection des écarts (mismatch, Mouser only...)
- Affichage côte à côte BOM / composant retenu dans la lib
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
from backend.BOM_function import fill_bom_result, check_bom, _add_comment
from backend.Pricing_lcsc import Update_Price_Stock, get_lcsc_price_from_page
from backend.BOM_report import build_bom_report, export_report_excel, _get_quantity_column


st.set_page_config(page_title="BOM Checker", page_icon="🔧", layout="wide")

# Librairie par défaut embarquée dans le repo (à côté de streamlit_app.py).
# L'utilisateur peut la remplacer via l'upload ci-dessous.
DEFAULT_LIB_PATH = Path(__file__).parent / "lib/Component_library.xlsx"


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


def describe_match(result) -> dict:
    """Résume le(s) composant(s) trouvé(s) dans la lib pour une ligne de BOM."""
    vide = {
        "Nb_matches": 0,
        "Lib_Manufacturer_Ref": "",
        "Lib_Value": "",
        "Lib_Footprint": "",
        "Lib_LCSC": "",
    }
    if result is None or result.empty:
        return vide

    def _join(col, max_items=3):
        if col not in result.columns:
            return ""
        vals = result[col].dropna().astype(str).unique()
        s = " | ".join(vals[:max_items])
        if len(vals) > max_items:
            s += f" (+{len(vals) - max_items})"
        return s

    return {
        "Nb_matches": len(result),
        "Lib_Manufacturer_Ref": _join("Manufacturer Ref"),
        "Lib_Value": _join("Value"),
        "Lib_Footprint": _join("Footprint"),
        "Lib_LCSC": _join("reference_LCSC"),
    }


def _norm(x) -> str:
    return "" if pd.isna(x) else str(x).strip().lower()


def match_verdict(row) -> str:
    """Verdict simple sur la correspondance BOM <-> lib."""
    n = row.get("Nb_matches", 0)
    if n == 0:
        return "❌ Aucun"
    if n > 1:
        return f"❓ {n} candidats"
    if _norm(row.get("Manufacturer Ref")) == _norm(row.get("Lib_Manufacturer_Ref")):
        return "✅ Identique"
    return "🔶 Réf différente"


def _color_match(val):
    if isinstance(val, str):
        if val.startswith("✅"):
            return "background-color: #d4edda"
        if val.startswith(("🔶", "❓")):
            return "background-color: #fff3cd"
        if val.startswith("❌"):
            return "background-color: #f8d7da"
    return ""



# ============================================================
# Choix d'un composant quand plusieurs candidats identiques (pop-up)
# ============================================================

@st.cache_data(ttl=900, show_spinner=False)
def _live_lcsc(ref: str, qty: int):
    """Prix + stock LCSC en direct (mis en cache 15 min pour ne pas re-scraper)."""
    try:
        return get_lcsc_price_from_page(ref, qty)
    except Exception:
        return None


def _is_filled(x) -> bool:
    return pd.notna(x) and str(x).strip() != ""


def _candidate_table(cands: pd.DataFrame, qty_total: int) -> pd.DataFrame:
    """Un tableau prix/stock par candidat (live LCSC, sinon valeurs de la lib)."""
    rows = []
    for lib_idx, c in cands.iterrows():
        ref = c.get("reference_LCSC")
        price, stock, source = None, None, "Pas de réf. LCSC"

        if _is_filled(ref):
            live = _live_lcsc(str(ref).strip(), qty_total)
            if live and live.get("prix_unitaire_eur") is not None:
                price, stock, source = live["prix_unitaire_eur"], live["stock"], "LCSC (live)"
            else:
                p = pd.to_numeric(c.get("Price"), errors="coerce")
                s_ = pd.to_numeric(c.get("Stock website"), errors="coerce")
                price = None if pd.isna(p) else float(p)
                stock = None if pd.isna(s_) else int(s_)
                source = "Librairie (cache)" if (price is not None or stock is not None) else "Indisponible"

        rows.append({
            "_lib_idx": lib_idx,
            "Manufacturer Ref": c.get("Manufacturer Ref"),
            "LCSC": ref if _is_filled(ref) else "",
            "Mouser": c.get("reference_Mouser") if _is_filled(c.get("reference_Mouser")) else "",
            "Prix unitaire (€)": price,
            "Stock": stock,
            "Stock suffisant": "—" if stock is None else ("✅" if stock >= qty_total else "❌"),
            "Prix ligne (€)": None if price is None else round(price * qty_total, 3),
            "Source": source,
        })
    return pd.DataFrame(rows)


def _recommend(table: pd.DataFrame, qty_total: int) -> int:
    """Position conseillée : le moins cher parmi ceux en stock suffisant, sinon le plus gros stock."""
    ok = table[table["Stock"].notna() & (table["Stock"] >= qty_total) & table["Prix unitaire (€)"].notna()]
    if not ok.empty:
        return table.index.get_loc(ok["Prix unitaire (€)"].idxmin())
    in_stock = table[table["Stock"].notna()]
    if not in_stock.empty:
        return table.index.get_loc(in_stock["Stock"].idxmax())
    return 0


def _apply_choice(index, chosen: pd.Series):
    """Écrit le composant choisi dans le BOM (et met à jour l'aperçu BOM/lib)."""
    bom = st.session_state.bom
    ref = chosen.get("reference_LCSC")

    if _is_filled(ref):
        bom.loc[index, "LCSC_part_number"] = str(ref).strip()
    else:
        mouser = chosen.get("reference_Mouser")
        msg = f"NO LCSC REF - MOUSER: {mouser}" if _is_filled(mouser) else "NO LCSC REF"
        line = _add_comment(bom.loc[index].copy(), msg)
        bom.loc[index] = line

    if st.session_state.matches is not None:
        for k, v in describe_match(chosen.to_frame().T).items():
            st.session_state.matches.loc[index, k] = v


@st.dialog("Plusieurs composants possibles", width="large")
def choose_component_dialog():
    pending = st.session_state.pending
    if not pending:
        st.rerun()

    index = next(iter(pending))
    cands = pending[index]
    bom = st.session_state.bom
    line = bom.loc[index]

    qty_col = _get_quantity_column(bom)
    qty_line = pd.to_numeric(line.get(qty_col), errors="coerce") if qty_col else 1
    qty_line = 1 if pd.isna(qty_line) else qty_line
    qty_total = max(int(qty_line * st.session_state.n_pcb), 1)

    st.caption(f"{len(pending)} ligne(s) restante(s) à trancher")
    st.markdown(
        f"**{line.get('References', '')}**  \n"
        f"Valeur : `{line.get('Value', '')}` — Footprint : `{line.get('Footprint', '')}`  \n"
        f"Quantité nécessaire : **{qty_total}** ({int(qty_line)} × {st.session_state.n_pcb} PCB)"
    )

    with st.spinner("Vérification prix / stock sur LCSC..."):
        table = _candidate_table(cands, qty_total)

    st.dataframe(table.drop(columns=["_lib_idx"]), hide_index=True, use_container_width=True)

    def _label(i):
        r = table.iloc[i]
        prix = "prix ?" if pd.isna(r["Prix unitaire (€)"]) else f"{r['Prix unitaire (€)']:.4f} €"
        stock = "stock ?" if pd.isna(r["Stock"]) else f"stock {int(r['Stock'])}"
        return f"{r['Manufacturer Ref']}  |  {r['LCSC'] or 'sans LCSC'}  |  {prix}  |  {stock}  {r['Stock suffisant']}"

    reco = _recommend(table, qty_total)
    st.caption("Présélection : le moins cher parmi ceux dont le stock couvre la quantité.")
    choice = st.radio(
        "Composant retenu",
        options=list(range(len(table))),
        index=reco,
        format_func=_label,
        key=f"choice_{index}",
    )

    def _done():
        pending.pop(index, None)
        # on enchaîne sur la ligne suivante
        st.session_state.auto_open_dialog = bool(pending)
        st.rerun()

    c_ok, c_skip = st.columns(2)
    if c_ok.button("✅ Valider ce composant", type="primary", use_container_width=True):
        _apply_choice(index, cands.loc[table.iloc[choice]["_lib_idx"]])
        _done()
    if c_skip.button("⏭️ Ignorer (laisser vide)", use_container_width=True):
        _done()


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
if "matches" not in st.session_state:
    st.session_state.matches = None
if "prices_updated" not in st.session_state:
    st.session_state.prices_updated = False
if "report" not in st.session_state:
    st.session_state.report = None
if "pending" not in st.session_state:
    st.session_state.pending = {}


st.title("🔧 BOM Checker")
st.caption("Recherche des composants dans la librairie, vérification des références, prix pour n PCB.")


# ============================================================
# 1. Upload des fichiers
# ============================================================

st.header("1. Fichiers")

col1, col2 = st.columns(2)

with col1:
    bom_file = st.file_uploader("BOM (.xlsx)", type=["xlsx"])
    header_row = 7

with col2:
    lib_file = st.file_uploader(
        "Librairie de composants (.xlsx) — optionnel",
        type=["xlsx"],
        help="Par défaut, la librairie du projet (Component_library.xlsx) est utilisée. "
             "Charge un fichier ici pour la remplacer.",
    )

if bom_file is not None:

    if lib_file is not None:
        lib_path = _save_uploaded_file(lib_file, st.session_state.workdir)
        st.caption(f"📚 Librairie utilisée : **{lib_file.name}** (chargée)")
    elif DEFAULT_LIB_PATH.exists():
        # On copie la lib par défaut dans le workdir pour ne jamais modifier
        # le fichier du repo (Update_Price_Stock écrit dans le fichier).
        lib_path = st.session_state.workdir / DEFAULT_LIB_PATH.name
        if not lib_path.exists():
            lib_path.write_bytes(DEFAULT_LIB_PATH.read_bytes())
        st.caption(f"📚 Librairie utilisée : **{DEFAULT_LIB_PATH.name}** (par défaut, celle du projet)")
    else:
        st.error(
            f"Aucune librairie chargée et {DEFAULT_LIB_PATH.name} introuvable dans le projet. "
            "Charge un fichier librairie."
        )
        st.stop()

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

st.number_input("Nombre de PCB", min_value=1, value=10, step=1, key="n_pcb",
                help="Sert à calculer la quantité totale pour vérifier le stock / le palier de prix.")

if st.button("🔍 Lancer la recherche et remplir le BOM", type="primary"):

    bom = st.session_state.bom.copy()
    lib = st.session_state.lib

    progress = st.progress(0, text="Recherche en cours...")
    nb = len(bom)
    matches = {}
    pending = {}

    for i, (index, row) in enumerate(bom.iterrows()):
        result = search_in_lib(lib, row)
        matches[index] = describe_match(result)
        row = fill_bom_result(row, result)
        row = check_bom(row, result)
        bom.loc[index] = row

        # Plusieurs candidats (après dédoublonnage) et rien de rempli -> choix manuel
        if (
            result is not None and len(result) > 1
            and not _is_filled(row.get("LCSC_part_number"))
        ):
            pending[index] = result

        progress.progress((i + 1) / nb, text=f"Ligne {i + 1}/{nb}")

    progress.empty()
    st.session_state.bom = bom
    st.session_state.matches = pd.DataFrame.from_dict(matches, orient="index")
    st.session_state.pending = pending
    st.session_state.auto_open_dialog = bool(pending)
    st.success("Recherche terminée.")

if st.session_state.pending:
    st.warning(f"{len(st.session_state.pending)} ligne(s) avec plusieurs composants possibles : "
               "à choisir selon le prix et le stock.")
    if st.button("🧩 Choisir les composants") or st.session_state.pop("auto_open_dialog", False):
        choose_component_dialog()

if st.session_state.bom is not None and "LCSC_part_number" in st.session_state.bom.columns:

    bom = st.session_state.bom

    apercu = bom.copy()
    apercu["Statut"] = apercu.apply(_status_for_display, axis=1)

    matches = st.session_state.get("matches")
    if matches is not None:
        apercu = apercu.join(matches)
        apercu["Match"] = apercu.apply(match_verdict, axis=1)

    nb_ok = (apercu["Statut"] == "✅ OK").sum()
    nb_warn = apercu["Statut"].str.startswith("⚠️").sum()
    nb_ko = (apercu["Statut"] == "❌ Non détecté").sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("OK", nb_ok)
    c2.metric("À vérifier", nb_warn)
    c3.metric("Non détecté", nb_ko)

    only_check = st.checkbox("Afficher uniquement les lignes à vérifier")
    if only_check and "Match" in apercu.columns:
        apercu = apercu[apercu["Match"] != "✅ Identique"]

    # Colonnes BOM / lib côte à côte pour comparer facilement
    cols_a_montrer = [c for c in [
        "References",
        "Match",
        "Manufacturer Ref", "Lib_Manufacturer_Ref",
        "Value", "Lib_Value",
        "Footprint", "Lib_Footprint",
        "LCSC_part_number", "Lib_LCSC",
        "Nb_matches", "Comments", "Statut",
    ] if c in apercu.columns]

    tableau = apercu[cols_a_montrer]
    if "Match" in tableau.columns:
        # pandas < 2.1 : remplacer .map par .applymap
        tableau = tableau.style.map(_color_match, subset=["Match"])

    st.dataframe(tableau, use_container_width=True, height=400)

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

n_pcb = st.session_state.n_pcb
st.caption(f"Rapport calculé pour {n_pcb} PCB (réglable dans la section 2).")
top_n = st.number_input("Nombre de lignes dans les classements", min_value=3, max_value=30, value=10, step=1)

if st.button("📊 Générer le rapport"):
    bom = st.session_state.bom
    lib = st.session_state.lib
    st.session_state.report = build_bom_report(bom, lib, n_pcb=n_pcb, top_n=top_n)
    st.session_state.report_n_pcb = n_pcb

if st.session_state.get("report") is not None:

    report = st.session_state.report
    n_pcb_affiche = st.session_state.get("report_n_pcb", n_pcb)

    detail = report["detail"]
    missing = report["missing"]
    total_price = report["total_price"]

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Prix total pour {n_pcb_affiche} PCB", f"{total_price:.2f} €")
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

    st.divider()

    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("💸 Composants les plus chers")
        top_expensive = report.get("top_expensive")
        if top_expensive is not None and not top_expensive.empty:
            st.dataframe(top_expensive, use_container_width=True, hide_index=True)
        else:
            st.info("Pas assez de données de prix pour ce classement.")

    with col_b:
        st.subheader("🔁 Composants les plus récurrents")
        most_recurring = report.get("most_recurring")
        if most_recurring is not None and not most_recurring.empty:
            st.dataframe(most_recurring, use_container_width=True, hide_index=True)
        else:
            st.info("Pas de colonne Manufacturer Ref pour ce classement.")

    st.subheader("🏷️ Coût par famille de composant")
    cost_by_type = report.get("cost_by_type")
    if cost_by_type is not None and not cost_by_type.empty:
        c_chart, c_table = st.columns([2, 1])
        with c_chart:
            st.bar_chart(cost_by_type)
        with c_table:
            st.dataframe(
                cost_by_type.rename("Coût total (€)").to_frame(),
                use_container_width=True,
            )
    else:
        st.info("Colonne 'type' absente de la librairie : impossible de grouper par famille.")

    st.divider()

    with st.expander("Voir le détail complet"):
        st.dataframe(detail, use_container_width=True)

    # Export Excel du rapport
    report_path = st.session_state.workdir / "rapport_bom.xlsx"
    export_report_excel(report, str(report_path), n_pcb=n_pcb_affiche)

    st.download_button(
        "⬇️ Télécharger le rapport (.xlsx)",
        data=report_path.read_bytes(),
        file_name="rapport_bom.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )