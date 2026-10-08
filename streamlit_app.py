"""
BOM Checker

Entête : fichiers (BOM, librairie), noms / nombre de PCB, bouton de recherche.
Onglets (verrouillés tant que les prérequis manquent) :
  🔍 Résultats BOM        -> BOM vs librairie        (BOM + recherche faite)
  📦 Composants & stock   -> liste consolidée + stock (BOM + recherche faite)
  💰 Prix & rapport       -> prix pour n PCB          (BOM + recherche faite)
  🛒 Commande             -> paniers LCSC / Mouser de ce qui manque (BOM + recherche faite)
  🏭 Production           -> retire du stock les PCB fabriqués (BOM + recherche faite)
  📚 Librairie            -> ajouter / modifier un composant (librairie seule)
  📥 Stock & prix         -> import commandes / saisie manuelle / prix LCSC (librairie seule)
"""

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from backend.find_component import search_in_lib
from backend.BOM_function import fill_bom_result, check_bom, _add_comment
from backend.Pricing_lcsc import Update_Price_Stock, get_lcsc_price_from_page, get_lcsc_product_info
from backend.BOM_report import build_bom_report, export_report_excel, _get_quantity_column
from backend.BOM_components import build_component_list
from backend.Bom_validate import row_status, can_validate, validate_row, revert_row
from backend.Stock_personnel import (
    STOCK_COL, read_mouser_file, read_lcsc_file, plan_stock_import, apply_stock_import,
    set_stock_manual, already_imported, build_stock_check, export_stock_check_excel,
    production_check, max_buildable, production_deductions, apply_production,
)
from backend.Order_export import (
    PASSIVE_TYPES, build_order_lines, to_lcsc_df, to_mouser_df, df_to_csv_bytes, df_to_xlsx_bytes,
)
from backend.Library_edit import NUMERIC_COLUMNS, find_duplicates, update_component, add_component


st.set_page_config(page_title="BOM Checker", page_icon="🔧", layout="wide")

# Librairie par défaut embarquée dans le repo (à côté de streamlit_app.py).
DEFAULT_LIB_PATH = Path(__file__).parent / "lib/Component_library.xlsx"

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# `use_container_width` est déprécié dans les Streamlit récents (-> width="stretch")
_ST_VERSION = tuple(int(x) for x in st.__version__.split(".")[:2] if x.isdigit())
STRETCH = {"width": "stretch"} if _ST_VERSION >= (1, 50) else {"use_container_width": True}


# ============================================================
# Style
# ============================================================

st.markdown(
    """
    <style>
    .block-container {max-width: 1500px; padding-top: 1.2rem;}
    .app-hero {padding: 1rem 1.4rem; border-radius: 14px; margin-bottom: .9rem;
               background: linear-gradient(120deg, #1f6feb 0%, #7c3aed 100%);}
    .app-hero h1 {margin: 0; padding: 0; font-size: 1.7rem; color: #fff;}
    .app-hero p {margin: .25rem 0 0; color: rgba(255,255,255,.88); font-size: .95rem;}
    div[data-testid="stMetric"] {border: 1px solid rgba(128,128,128,.25); border-radius: 12px;
                                 padding: .7rem 1rem; background: rgba(128,128,128,.06);}
    .chip {display: inline-block; padding: .18rem .65rem; border-radius: 999px; font-size: .82rem;
           margin: .15rem .3rem .15rem 0; font-weight: 500;}
    .chip-ok {background: rgba(46,160,67,.18); color: #2ea043;}
    .chip-warn {background: rgba(210,153,34,.20); color: #d29922;}
    .chip-off {background: rgba(128,128,128,.18); color: #8b949e;}
    .lock-card {border: 1px dashed rgba(128,128,128,.5); border-radius: 14px; padding: 2rem 1.5rem;
                text-align: center; background: rgba(128,128,128,.05); margin-top: .5rem;}
    .lock-card h3 {margin: 0 0 .4rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def chip(text: str, kind: str = "off") -> str:
    return f'<span class="chip chip-{kind}">{text}</span>'


# ============================================================
# Utilitaires
# ============================================================

def _df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="BOM")
    return buffer.getvalue()


def _is_filled(x) -> bool:
    return pd.notna(x) and str(x).strip() != ""


def _norm(x) -> str:
    return "" if pd.isna(x) else str(x).strip().lower()


def describe_match(result) -> dict:
    """Résume le(s) composant(s) trouvé(s) dans la lib pour une ligne de BOM."""
    vide = {"Nb_matches": 0, "Lib_Manufacturer_Ref": "", "Lib_Value": "", "Lib_Footprint": "", "Lib_LCSC": "",
            "Lib_idx": None}
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
        "Lib_idx": result.index[0] if len(result) == 1 else None,   # sert à valider la proposition
    }


def _color_stock(val):
    if val in ("RIEN EN STOCK", "ABSENT DE LA LIB"):
        return "background-color: #f8d7da; color: #58151c"
    if val in ("MANQUE", "NON IDENTIFIÉ"):
        return "background-color: #fff3cd; color: #664d03"
    if val == "OK":
        return "background-color: #d4edda; color: #1b4332"
    return ""


def _prepare_bom_frame(raw: pd.DataFrame, i: int) -> pd.DataFrame:
    if "Row" in raw.columns:
        raw = raw.drop(columns=["Row"])
    if "LCSC_part_number" not in raw.columns:
        raw["LCSC_part_number"] = pd.NA
    if "Comments" not in raw.columns:
        raw["Comments"] = pd.NA
    for col in ("Manufacturer Ref", "Manufacturer"):
        if col not in raw.columns:
            raw[col] = pd.NA
    for col in ("LCSC_part_number", "Comments", "Manufacturer Ref", "Manufacturer"):
        raw[col] = raw[col].astype("object")
    raw["BOM_id"] = i
    return raw


# ============================================================
# Choix d'un composant quand plusieurs candidats (pop-up)
# ============================================================

@st.cache_data(ttl=900, show_spinner=False)
def _live_lcsc(ref: str, qty: int):
    """Prix + stock LCSC en direct (mis en cache 15 min pour ne pas re-scraper)."""
    try:
        return get_lcsc_price_from_page(ref, qty)
    except Exception:
        return None


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
            "Ligne lib (Excel)": int(lib_idx) + 2,
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
        bom.loc[index] = _add_comment(bom.loc[index].copy(), msg)

    if st.session_state.matches is not None:
        for k, v in describe_match(chosen.to_frame().T).items():
            st.session_state.matches.loc[index, k] = v


def _signature(line) -> tuple:
    """Deux lignes de BOM avec la même signature demandent le même composant."""
    return (_norm(line.get("Manufacturer Ref")), _norm(line.get("Value")), _norm(line.get("Footprint")))


def _qty_total_for(bom: pd.DataFrame, idx) -> float:
    qty_col = _get_quantity_column(bom)
    q = pd.to_numeric(bom.loc[idx, qty_col], errors="coerce") if qty_col else 1
    q = 1 if pd.isna(q) else q
    n = pd.to_numeric(bom.loc[idx, "N_PCB"], errors="coerce") if "N_PCB" in bom.columns else 1
    n = 1 if pd.isna(n) else n
    return q * n


@st.dialog("Plusieurs composants possibles", width="large")
def choose_component_dialog():
    pending = st.session_state.pending
    if not pending:
        st.rerun()

    index = next(iter(pending))
    cands = pending[index]
    bom = st.session_state.bom
    line = bom.loc[index]

    # Même composant demandé dans plusieurs BOM : une seule question, quantités cumulées
    sig = _signature(line)
    group = [k for k in pending if _signature(bom.loc[k]) == sig]
    qty_total = max(int(sum(_qty_total_for(bom, k) for k in group)), 1)
    boms_txt = ", ".join(dict.fromkeys(str(bom.loc[k, "BOM"]) for k in group)) if "BOM" in bom.columns else ""

    st.caption(f"{len(pending)} ligne(s) restante(s) à trancher")
    st.markdown(
        f"**{line.get('References', '')}**  \n"
        f"Valeur : `{line.get('Value', '')}` — Footprint : `{line.get('Footprint', '')}`  \n"
        f"BOM : **{boms_txt}** — quantité nécessaire : **{qty_total}**"
    )
    if len(group) > 1:
        st.info(f"Ce choix s'appliquera aux {len(group)} lignes identiques des BOM : {boms_txt}.")

    with st.spinner("Vérification prix / stock sur LCSC..."):
        table = _candidate_table(cands, qty_total)

    st.dataframe(table.drop(columns=["_lib_idx"]), hide_index=True, **STRETCH)

    refs = table["LCSC"].replace("", pd.NA).dropna()
    if refs.duplicated().any():
        st.info("ℹ️ Plusieurs lignes de la librairie ont la même réf LCSC (doublons) : "
                "le choix donnera le même résultat. Pense à supprimer le doublon dans la lib.")

    def _label(i):
        r = table.iloc[i]
        prix = "prix ?" if pd.isna(r["Prix unitaire (€)"]) else f"{r['Prix unitaire (€)']:.4f} €"
        stock = "stock ?" if pd.isna(r["Stock"]) else f"stock {int(r['Stock'])}"
        return (f"Lib l.{r['Ligne lib (Excel)']}  |  {r['Manufacturer Ref']}  |  "
                f"{r['LCSC'] or 'sans LCSC'}  |  {prix}  |  {stock}  {r['Stock suffisant']}")

    reco = _recommend(table, qty_total)
    st.caption("Présélection : le moins cher parmi ceux dont le stock couvre la quantité.")
    choice = st.radio("Composant retenu", options=list(range(len(table))), index=reco,
                      format_func=_label, key=f"choice_{index}")

    def _done():
        for k in group:
            pending.pop(k, None)
        st.session_state.auto_open_dialog = bool(pending)   # on enchaîne sur la ligne suivante
        st.rerun()

    c_ok, c_skip = st.columns(2)
    if c_ok.button("✅ Valider ce composant", type="primary", **STRETCH):
        chosen = cands.loc[table.iloc[choice]["_lib_idx"]]
        for k in group:
            _apply_choice(k, chosen)
        _done()
    if c_skip.button("⏭️ Ignorer (laisser vide)", **STRETCH):
        _done()


# ============================================================
# Etat de session
# ============================================================

_defaults = {
    "workdir": None, "bom_raw": None, "bom": None, "bom_signature": None,
    "lib_path": None, "lib_key": None, "lib": None,
    "matches": None, "pending": {}, "nav": "stock",
    "lib_version": 0, "search_lib_version": None, "validated": {}, "results_ver": 0,
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v
if st.session_state.workdir is None:
    st.session_state.workdir = Path(tempfile.mkdtemp(prefix="bom_checker_"))


def run_search():
    """Recherche chaque ligne du BOM dans la lib (repart toujours du BOM d'origine)."""
    bom = st.session_state.bom_raw.copy()
    lib = st.session_state.lib

    progress = st.progress(0, text="Recherche en cours...")
    nb = max(len(bom), 1)
    matches, pending = {}, {}

    for i, (index, row) in enumerate(bom.iterrows()):
        result = search_in_lib(lib, row)
        matches[index] = describe_match(result)
        row = fill_bom_result(row, result)
        row = check_bom(row, result)
        bom.loc[index] = row

        # Plusieurs candidats et rien de rempli -> choix manuel (pop-up)
        if result is not None and len(result) > 1 and not _is_filled(row.get("LCSC_part_number")):
            pending[index] = result

        progress.progress((i + 1) / nb, text=f"Ligne {i + 1}/{len(bom)}")

    progress.empty()
    st.session_state.bom = bom
    st.session_state.matches = pd.DataFrame.from_dict(matches, orient="index")
    st.session_state.pending = pending
    st.session_state.validated = {}
    st.session_state.auto_open_dialog = bool(pending)
    st.session_state.search_lib_version = st.session_state.lib_version
    st.session_state.nav = "results"


# ============================================================
# ENTÊTE : fichiers + recherche
# ============================================================

st.markdown(
    '<div class="app-hero"><h1>🔧 BOM Checker</h1>'
    "<p>Retrouve tes composants dans la librairie, calcule le prix et vérifie ton stock.</p></div>",
    unsafe_allow_html=True,
)

with st.container(border=True):

    c_bom, c_lib = st.columns(2)

    with c_bom:
        bom_files = st.file_uploader("BOM (.xlsx) — une ou plusieurs", type=["xlsx"], accept_multiple_files=True)

    with c_lib:
        lib_file = st.file_uploader(
            "Librairie de composants (.xlsx) — optionnel", type=["xlsx"],
            help="Par défaut, la librairie du projet est utilisée. Charge un fichier pour la remplacer.",
        )

    # ---- Librairie (chargée même sans BOM) ----
    if lib_file is not None:
        lib_path = st.session_state.workdir / lib_file.name
        lib_key = (lib_file.name, lib_file.size)
        if st.session_state.lib_key != lib_key:
            lib_path.write_bytes(lib_file.getbuffer())   # écrite une seule fois (sinon prix/stock écrasés)
            st.session_state.matches = None              # la recherche dépendait de l'ancienne lib
            st.session_state.pending = {}
            st.session_state.lib_key = lib_key
        lib_label = f"{lib_file.name} (chargée)"
    elif DEFAULT_LIB_PATH.exists():
        # copie de travail : Update_Price_Stock / import de stock écrivent dans le fichier
        lib_path = st.session_state.workdir / DEFAULT_LIB_PATH.name
        if not lib_path.exists():
            lib_path.write_bytes(DEFAULT_LIB_PATH.read_bytes())
        if st.session_state.lib_key is not None:         # retour à la lib par défaut
            st.session_state.lib_key = None
            st.session_state.matches = None
            st.session_state.pending = {}
        lib_label = f"{DEFAULT_LIB_PATH.name} (préinstallée)"
    else:
        lib_path, lib_label = None, None

    if lib_path is None:
        st.error(f"Aucune librairie : {DEFAULT_LIB_PATH.name} introuvable dans le projet. Charge un fichier librairie.")
        st.stop()

    try:
        st.session_state.lib = pd.read_excel(lib_path)
        st.session_state.lib_path = lib_path
    except Exception as e:
        st.error(f"Erreur de lecture de la librairie : {e}")
        st.stop()

    # ---- BOM ----
    bom_settings = []
    if bom_files:
        with st.expander("Nom et nombre de PCB de chaque BOM", expanded=True):
            for i, f in enumerate(bom_files):
                c_name, c_n = st.columns([3, 1])
                name = c_name.text_input(
                    f"Nom attribué à « {f.name} »", value=Path(f.name).stem,
                    key=f"bom_name_{i}_{f.name}_{f.size}",
                )
                n = c_n.number_input(
                    "Nb de PCB", min_value=1, value=10, step=1, key=f"bom_npcb_{i}_{f.name}_{f.size}",
                    help="Sert aux quantités totales, au stock à vérifier et au prix.",
                )
                bom_settings.append({"name": name.strip() or Path(f.name).stem, "n_pcb": int(n)})

            noms = [b["name"] for b in bom_settings]
            if len(set(noms)) < len(noms):
                st.warning("Deux BOM ont le même nom : donne-leur des noms différents.")

        signature = tuple((f.name, f.size) for f in bom_files)
        if st.session_state.bom_signature != signature:
            try:
                frames = [_prepare_bom_frame(pd.read_excel(f, header=7).dropna(how="all"), i)
                          for i, f in enumerate(bom_files)]
            except Exception as e:
                st.error(f"Erreur de lecture du BOM : {e}")
                st.stop()
            raw = pd.concat(frames, ignore_index=True)
            st.session_state.bom_raw = raw
            st.session_state.bom = raw.copy()
            st.session_state.bom_signature = signature
            st.session_state.matches = None
            st.session_state.pending = {}
            st.session_state.validated = {}

        # Nom / nb de PCB modifiables à tout moment sans perdre la recherche
        names = {i: b["name"] for i, b in enumerate(bom_settings)}
        npcbs = {i: b["n_pcb"] for i, b in enumerate(bom_settings)}
        for key in ("bom_raw", "bom"):
            st.session_state[key]["BOM"] = st.session_state[key]["BOM_id"].map(names)
            st.session_state[key]["N_PCB"] = st.session_state[key]["BOM_id"].map(npcbs)

    elif st.session_state.bom_signature is not None:
        # BOM retirés : on remet tout à zéro (onglets BOM re-verrouillés)
        for key in ("bom_raw", "bom", "bom_signature", "matches"):
            st.session_state[key] = None
        st.session_state.pending = {}
        st.session_state.validated = {}

    has_bom = st.session_state.bom_raw is not None
    searched = has_bom and st.session_state.matches is not None
    n_pending = len(st.session_state.pending)

    # ---- Statut + bouton de recherche ----
    c_status, c_btn = st.columns([3, 1.2], vertical_alignment="center")

    with c_btn:
        if st.button("🔍 Lancer la recherche", type="primary", **STRETCH, disabled=not has_bom,
                     help="Obligatoire pour débloquer les onglets liés au BOM."):
            run_search()
            searched, n_pending = True, len(st.session_state.pending)
        if n_pending:
            if st.button(f"🧩 {n_pending} à trancher", **STRETCH):
                st.session_state.auto_open_dialog = True

    with c_status:
        chips = chip(f"📚 Librairie : {lib_label} · {len(st.session_state.lib)} composants", "ok")
        if has_bom:
            chips += chip(f"📄 {len(bom_files)} BOM · {len(st.session_state.bom_raw)} lignes", "ok")
        else:
            chips += chip("📄 Aucun BOM chargé", "off")
        if searched:
            chips += chip("🔍 Recherche faite", "ok")
        elif has_bom:
            chips += chip("🔍 Recherche à lancer", "warn")
        else:
            chips += chip("🔍 Recherche en attente d'un BOM", "off")
        if n_pending:
            chips += chip(f"🧩 {n_pending} ligne(s) à trancher", "warn")
        if searched and st.session_state.search_lib_version != st.session_state.lib_version:
            chips += chip("⚠️ Librairie modifiée depuis la recherche : relance-la", "warn")
        st.markdown(chips, unsafe_allow_html=True)

    if st.session_state.pending and st.session_state.pop("auto_open_dialog", False):
        choose_component_dialog()


# ============================================================
# NAVIGATION (onglets verrouillables)
# ============================================================

TABS = {
    "results":    {"icon": "🔍", "label": "Résultats BOM",       "needs_bom": True},
    "components": {"icon": "📦", "label": "Composants & stock",  "needs_bom": True},
    "order":      {"icon": "🛒", "label": "Commande",            "needs_bom": True},
    "production": {"icon": "🏭", "label": "Production",          "needs_bom": True},
    "report":     {"icon": "💰", "label": "Prix & rapport",      "needs_bom": True},
    "library":    {"icon": "📚", "label": "Librairie",           "needs_bom": False},
    "stock":      {"icon": "📥", "label": "Stock & prix",        "needs_bom": False},
}


def _is_locked(key: str) -> bool:
    return TABS[key]["needs_bom"] and not searched


def _tab_label(key: str) -> str:
    return f"{'🔒' if _is_locked(key) else TABS[key]['icon']} {TABS[key]['label']}"


keys = list(TABS)
if st.session_state.nav not in keys:   # ex. clic sur l'onglet déjà actif (désélection)
    st.session_state.nav = st.session_state.get("_last_nav", "stock")

if hasattr(st, "segmented_control"):
    selected = st.segmented_control("Navigation", keys, format_func=_tab_label, key="nav",
                                    label_visibility="collapsed")
else:  # Streamlit plus ancien
    selected = st.radio("Navigation", keys, format_func=_tab_label, key="nav", horizontal=True,
                        label_visibility="collapsed")
selected = selected or st.session_state.get("_last_nav", "stock")   # clic sur l'onglet déjà actif
st.session_state["_last_nav"] = selected


def lock_card(missing: list):
    items = "<br>".join(f"• {m}" for m in missing)
    st.markdown(
        f'<div class="lock-card"><h3>🔒 Onglet verrouillé</h3>'
        f"<div>Pour le débloquer :<br>{items}</div></div>",
        unsafe_allow_html=True,
    )


def missing_for_bom_tabs() -> list:
    out = []
    if not has_bom:
        out.append("charge au moins un BOM (en haut)")
    out.append("lance la recherche (bouton « 🔍 Lancer la recherche » en haut)")
    return out


# ============================================================
# ONGLET : Résultats BOM
# ============================================================

def _validate(index):
    """Le BOM adopte le composant proposé par la lib pour cette ligne."""
    bom, matches, lib = st.session_state.bom, st.session_state.matches, st.session_state.lib
    st.session_state.validated[index] = validate_row(bom, lib, index, matches.loc[index, "Lib_idx"])


def _unvalidate(index):
    revert_row(st.session_state.bom, index, st.session_state.validated.pop(index))


def tab_results():
    bom, matches = st.session_state.bom, st.session_state.matches
    validated = st.session_state.validated

    apercu = bom.join(matches)
    res = [row_status(r, idx in validated) for idx, r in apercu.iterrows()]
    apercu["_kind"] = [k for k, _ in res]
    apercu["Statut"] = [t for _, t in res]
    apercu["_can"] = [can_validate(k, r) for k, (_, r) in zip(apercu["_kind"], apercu.iterrows())]
    apercu["Valider"] = apercu["_kind"] == "validated"

    kinds = apercu["_kind"]
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("🟢 OK", int(kinds.isin(["ok", "validated"]).sum()))
    m2.metric("🟣 Réf différente", int((kinds == "ref").sum()))
    m3.metric("🔵 Réf Mouser", int((kinds == "mouser").sum()))
    m4.metric("🟠 À valider", int((kinds == "mismatch").sum()))
    m5.metric("🔴 Problèmes", int((kinds == "bad").sum()))

    st.caption(
        "🟢 OK · 🟣 trouvé mais réf fabricant différente · 🔵 réf Mouser seulement · "
        "🟠 réf LCSC du BOM ≠ lib · 🔴 problème.  "
        "Pour 🟣 et 🟠, coche **Valider** : le BOM adopte le composant de la lib et la ligne passe en 🟢 "
        "(décoche pour annuler)."
    )

    if st.session_state.get("results_note"):
        st.info(st.session_state.pop("results_note"))

    c_bulk, c_filter = st.columns([2, 1], vertical_alignment="center")
    n_ref = int(((kinds == "ref") & apercu["_can"]).sum())
    if n_ref and c_bulk.button(f"✔ Valider les {n_ref} ligne(s) 🟣 d'un coup"):
        for idx in apercu.index[(kinds == "ref") & apercu["_can"]]:
            _validate(idx)
        st.session_state.results_ver += 1
        st.session_state.results_note = f"{n_ref} ligne(s) validée(s) : le BOM utilise maintenant les réfs de la lib."
        st.rerun()
    only_check = c_filter.checkbox("Uniquement les lignes à vérifier")
    if only_check:
        apercu = apercu[~apercu["_kind"].isin(["ok", "validated"])]

    cols = ["Valider", "Statut", "BOM", "References", "Manufacturer Ref", "Lib_Manufacturer_Ref",
            "Value", "Lib_Value", "Footprint", "Lib_Footprint", "LCSC_part_number", "Lib_LCSC",
            "Nb_matches", "Comments"]
    cols = [c for c in cols if c in apercu.columns]
    view = apercu[cols]

    edited = st.data_editor(
        view, hide_index=True, height=480, **STRETCH,
        key=f"results_editor_{st.session_state.results_ver}_{int(only_check)}",   # reset après chaque validation
        disabled=[c for c in cols if c != "Valider"],
        column_config={
            "Valider": st.column_config.CheckboxColumn("Valider", width="small"),
            "Statut": st.column_config.TextColumn("Statut", width="medium"),
        },
    )

    changed = edited.index[edited["Valider"] != view["Valider"]]
    if len(changed):
        refused = 0
        for idx in changed:
            if bool(edited.loc[idx, "Valider"]):
                if apercu.loc[idx, "_can"]:
                    _validate(idx)
                else:
                    refused += 1
            elif idx in validated:
                _unvalidate(idx)
        if refused:
            st.session_state.results_note = (f"{refused} ligne(s) sans proposition unique à valider : "
                                             "seules les lignes 🟣 et 🟠 se valident.")
        st.session_state.results_ver += 1
        st.rerun()

    st.download_button(
        "⬇️ Télécharger le BOM vérifié",
        data=_df_to_excel_bytes(st.session_state.bom.drop(columns=["BOM_id"], errors="ignore")),
        file_name="bom_verifie.xlsx", mime=XLSX_MIME,
    )


# ============================================================
# ONGLET : Composants & stock (liste consolidée + stock)
# ============================================================

def tab_components():
    bom, lib = st.session_state.bom, st.session_state.lib

    comps, per_bom = build_component_list(bom, lib)
    check = build_stock_check(comps, lib)

    if STOCK_COL not in lib.columns:
        st.warning(f"La librairie n'a pas de colonne « {STOCK_COL} » : tout le stock est considéré à 0. "
                   "Ajoute du stock dans l'onglet 📚 Librairie & stock.")

    manque = check[check["Manquant"].fillna(0) > 0]
    cout = manque["Coût du manquant (€)"].sum(skipna=True)
    nb_unknown = int((check["Statut"] == "NON IDENTIFIÉ").sum())

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Composants distincts", len(check))
    k2.metric("OK", int((check["Statut"] == "OK").sum()))
    k3.metric("En manque", len(manque))
    k4.metric("Pièces manquantes", int(manque["Manquant"].sum()))
    k5.metric("Coût du réassort", f"{cout:.2f} €")

    if nb_unknown:
        st.warning(f"{nb_unknown} ligne(s) du BOM ne sont pas rattachées à la librairie (NON IDENTIFIÉ) : "
                   "leur stock n'est pas vérifié. Résous-les via « 🧩 à trancher » en haut, ou ajoute-les à la lib.")

    f1, f2 = st.columns([1.2, 2])
    vue_mode = f1.radio("Afficher", ["À traiter", "Tout", "OK"], horizontal=True, key="comp_view")
    q = f2.text_input("🔎 Filtrer", placeholder="composant, valeur, type, BOM, réf LCSC / Mouser…", key="comp_q")

    vue = check
    if vue_mode == "À traiter":
        vue = vue[vue["Statut"] != "OK"]
    elif vue_mode == "OK":
        vue = vue[vue["Statut"] == "OK"]
    if q:
        mask = vue.astype(str).apply(lambda c: c.str.contains(q, case=False, na=False, regex=False)).any(axis=1)
        vue = vue[mask]

    cols = ["Statut", "Composant", "Valeur", "Description", "Type", "Footprint", "Ref LCSC", "Ref Mouser",
            "BOM concernées", "Qté par BOM", "Besoin", "Stock perso", "Manquant",
            "Prix unitaire (€)", "Coût du manquant (€)", "Ligne lib (Excel)"]
    st.dataframe(vue[cols].style.map(_color_stock, subset=["Statut"]), **STRETCH,
                 hide_index=True, height=460)

    buf = io.BytesIO()
    export_stock_check_excel(check, buf, per_bom=per_bom)
    st.download_button("⬇️ Télécharger la liste (composants, stock, à commander)", data=buf.getvalue(),
                       file_name="composants_stock.xlsx", mime=XLSX_MIME)


# ============================================================
# ONGLET : Prix & rapport
# ============================================================

def tab_report():
    bom, lib = st.session_state.bom, st.session_state.lib

    n_pcb = int(st.session_state.bom_raw.drop_duplicates("BOM_id")["N_PCB"].sum())
    c_info, c_top = st.columns([3, 1])
    c_info.caption(f"Le rapport couvre {n_pcb} PCB au total sur {st.session_state.bom_raw['BOM_id'].nunique()} BOM "
                   "(nombre de PCB réglé par BOM dans l'entête).")
    top_n = c_top.number_input("Lignes dans les classements", min_value=3, max_value=30, value=10, step=1)

    if "Price" not in lib.columns or pd.to_numeric(lib["Price"], errors="coerce").notna().sum() == 0:
        st.warning("Aucun prix dans la librairie. Lance « Mettre à jour les prix » dans l'onglet 📚 Librairie & stock.")

    report = build_bom_report(bom, lib, n_pcb=n_pcb, top_n=top_n)
    detail, missing, total_price = report["detail"], report["missing"], report["total_price"]

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Prix total ({n_pcb} PCB)", f"{total_price:.2f} €")
    c2.metric("Lignes OK", int((detail["Status"] == "OK").sum()))
    c3.metric("Lignes à vérifier", len(missing))

    if not missing.empty:
        with st.expander(f"Composants à vérifier ({len(missing)})"):
            cols = [c for c in ["BOM", "Manufacturer Ref", "Value", "Footprint", "LCSC_part_number", "Status", "Comments"]
                    if c in missing.columns]
            st.dataframe(missing[cols], **STRETCH, hide_index=True)
    else:
        st.success("Tous les composants sont correctement identifiés.")

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("💸 Composants les plus chers")
        te = report.get("top_expensive")
        if te is not None and not te.empty:
            st.dataframe(te, **STRETCH, hide_index=True)
        else:
            st.info("Pas assez de données de prix pour ce classement.")
    with col_b:
        st.subheader("🔁 Composants les plus récurrents")
        mr = report.get("most_recurring")
        if mr is not None and not mr.empty:
            st.dataframe(mr, **STRETCH, hide_index=True)
        else:
            st.info("Pas de colonne Manufacturer Ref pour ce classement.")

    st.subheader("🏷️ Coût par famille de composant")
    cbt = report.get("cost_by_type")
    if cbt is not None and not cbt.empty:
        c_chart, c_table = st.columns([2, 1])
        c_chart.bar_chart(cbt)
        c_table.dataframe(cbt.rename("Coût total (€)").to_frame(), **STRETCH)
    else:
        st.info("Colonne 'type' absente de la librairie : impossible de grouper par famille.")

    with st.expander("Voir le détail complet"):
        st.dataframe(detail, **STRETCH)

    report_path = st.session_state.workdir / "rapport_bom.xlsx"
    export_report_excel(report, str(report_path), n_pcb=n_pcb)
    st.download_button("⬇️ Télécharger le rapport (.xlsx)", data=report_path.read_bytes(),
                       file_name="rapport_bom.xlsx", mime=XLSX_MIME)


# ============================================================
# ONGLET : Librairie & stock
# ============================================================

SOURCES = ["LCSC / JLCPCB (.csv)", "Mouser (.xls / .xlsx)", "Saisie manuelle"]


def _stock_import_block(source: str):
    """Import d'un fichier de commande (LCSC ou Mouser) avec aperçu avant écriture."""
    is_lcsc = source.startswith("LCSC")
    files = st.file_uploader(
        "Commandes / paniers LCSC-JLCPCB (.csv)" if is_lcsc else "Commandes / paniers Mouser (.xls / .xlsx)",
        type=["csv"] if is_lcsc else ["xls", "xlsx"],
        accept_multiple_files=True, key=f"stock_up_{'lcsc' if is_lcsc else 'mouser'}",
    )
    if not files:
        st.info("Charge un ou plusieurs fichiers : les quantités seront **ajoutées** au stock personnel.")
        return

    try:
        reader = read_lcsc_file if is_lcsc else read_mouser_file
        entries = pd.concat([reader(f, f.name) for f in files], ignore_index=True)
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
        return
    if entries.empty:
        st.warning("Aucune ligne exploitable dans ce(s) fichier(s).")
        return

    plan = plan_stock_import(st.session_state.lib, entries)
    ok, ko = plan[plan["Statut"] == "OK"], plan[plan["Statut"] != "OK"]

    m1, m2, m3 = st.columns(3)
    m1.metric("Lignes rapprochées", len(ok))
    m2.metric("Quantité à ajouter", int(ok["Qty"].sum()))
    m3.metric("Introuvables dans la lib", len(ko))

    cols_plan = ["Source", "Ref", "MPN", "Qty", "Ligne lib (Excel)", "Composant lib", "Trouvé par",
                 "Stock avant", "Stock après", "Suggestion lib", "Statut"]
    st.dataframe(plan[cols_plan], **STRETCH, hide_index=True)

    if not ko.empty:
        st.warning(f"{len(ko)} ligne(s) introuvable(s) dans la librairie. Regarde « Suggestion lib » : si c'est le même "
                   "composant, ajoute la réf fournisseur dans la lib puis relance l'import. "
                   "Sinon, coche la case ci-dessous pour les ajouter.")
    add_missing = st.checkbox("Ajouter à la librairie les composants introuvables (nouvelles lignes en bas)",
                              value=False, disabled=ko.empty, key=f"add_missing_{source}")

    deja = already_imported(st.session_state.lib_path, [f.name for f in files])
    confirm = True
    if deja:
        st.warning("Déjà importé : " + ", ".join(f"{k} ({v})" for k, v in deja.items()) +
                   ". Un 2e import compterait ces quantités deux fois.")
        confirm = st.checkbox("Importer quand même", key=f"force_{source}")

    if st.button("📥 Ajouter au stock personnel", type="primary", disabled=not confirm, key=f"apply_{source}"):
        n = apply_stock_import(st.session_state.lib_path, plan, add_missing=add_missing)
        st.session_state.lib = pd.read_excel(st.session_state.lib_path)
        st.session_state.stock_msg = f"Stock mis à jour : {n} ligne(s) modifiée(s)/ajoutée(s)."
        st.rerun()


def _stock_manual_block():
    """Recherche d'un composant dans la lib + édition directe du stock."""
    lib = st.session_state.lib

    q = st.text_input("🔎 Rechercher un composant dans la librairie",
                      placeholder="réf fabricant, valeur, description, réf LCSC / Mouser…", key="manual_q")

    search_cols = [c for c in ["Manufacturer Ref", "Value", "description", "Footprint", "type",
                               "reference_LCSC", "reference_Mouser", "Manufacturer Part"] if c in lib.columns]
    mask = pd.Series(True, index=lib.index)
    if q:
        mask = lib[search_cols].astype(str).apply(
            lambda c: c.str.contains(q, case=False, na=False, regex=False)).any(axis=1)

    show_cols = [c for c in ["Manufacturer Ref", "Value", "Footprint", "type", "reference_LCSC", "reference_Mouser"]
                 if c in lib.columns]
    view = lib.loc[mask, show_cols].copy()
    current = (pd.to_numeric(lib[STOCK_COL], errors="coerce").fillna(0)
               if STOCK_COL in lib.columns else pd.Series(0.0, index=lib.index))
    view[STOCK_COL] = current[mask]

    st.caption(f"{len(view)} composant(s) affiché(s). Modifie la colonne « {STOCK_COL} » puis enregistre "
               "(enregistre avant de changer la recherche).")

    edited = st.data_editor(
        view, hide_index=True, **STRETCH, height=420,
        disabled=show_cols, key=f"manual_editor_{st.session_state.get('manual_ver', 0)}_{q}",
        column_config={STOCK_COL: st.column_config.NumberColumn(STOCK_COL, min_value=0, step=1, format="%d")},
    )

    changes = {idx: edited.loc[idx, STOCK_COL] for idx in edited.index
               if pd.notna(edited.loc[idx, STOCK_COL]) and edited.loc[idx, STOCK_COL] != view.loc[idx, STOCK_COL]}

    if st.button(f"💾 Enregistrer {len(changes)} modification(s)", type="primary", disabled=not changes):
        n = set_stock_manual(st.session_state.lib_path, lib, changes)
        st.session_state.lib = pd.read_excel(st.session_state.lib_path)
        st.session_state.manual_ver = st.session_state.get("manual_ver", 0) + 1
        st.session_state.stock_msg = f"Stock mis à jour manuellement : {n} composant(s)."
        st.rerun()


def tab_stock():
    lib = st.session_state.lib

    nb_stock = int(pd.to_numeric(lib[STOCK_COL], errors="coerce").notna().sum()) if STOCK_COL in lib.columns else 0
    k1, k2, k3 = st.columns(3)
    k1.metric("Composants dans la lib", len(lib))
    k2.metric("Avec un stock renseigné", nb_stock)
    k3.metric("Pièces en stock (total)",
              int(pd.to_numeric(lib[STOCK_COL], errors="coerce").sum()) if STOCK_COL in lib.columns else 0)

    st.subheader("Stock personnel")
    source = st.selectbox("Comment ajouter du stock ?", SOURCES, key="stock_source")
    if source == "Saisie manuelle":
        _stock_manual_block()
    else:
        _stock_import_block(source)

    st.divider()

    with st.expander("💰 Mettre à jour les prix et stocks LCSC (site)"):
        st.caption("Interroge LCSC.com pour chaque composant de la librairie (1 requête/seconde). "
                   "Peut prendre plusieurs minutes selon la taille de la librairie.")
        if st.button("Lancer la mise à jour des prix"):
            with st.spinner("Récupération des prix sur LCSC..."):
                try:
                    Update_Price_Stock(str(st.session_state.lib_path))
                    st.session_state.lib = pd.read_excel(st.session_state.lib_path)
                    st.session_state.stock_msg = "Prix et stocks LCSC mis à jour."
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur pendant la mise à jour : {e}")

    st.download_button(
        "⬇️ Télécharger la librairie mise à jour (.xlsx)",
        data=Path(st.session_state.lib_path).read_bytes(),
        file_name=Path(st.session_state.lib_path).name, mime=XLSX_MIME,
        help="Contient le stock personnel, les prix à jour et la feuille « Historique stock ». "
             "Remplace ton fichier de librairie par celui-ci.",
    )



# ============================================================
# ONGLET : Commande (paniers LCSC / Mouser)
# ============================================================

def tab_order():
    bom, lib = st.session_state.bom, st.session_state.lib
    comps, per_bom = build_component_list(bom, lib)
    check = build_stock_check(comps, lib)

    st.caption("Ce qui manque en stock (toutes BOM), prêt à importer dans l'outil BOM de LCSC / JLCPCB et de Mouser.")

    p1, p2, p3, p4 = st.columns([1.1, 1, 1.6, 1])
    prefer = p1.radio("Fournisseur préféré", ["LCSC", "Mouser"], horizontal=True, key="order_prefer",
                      help="L'autre sert de repli si le composant n'a pas de réf chez le fournisseur préféré.")
    margin = p2.number_input("Marge générale (%)", 0, 100, 0, 5, key="order_margin")
    types = sorted(str(t) for t in check["Type"].dropna().unique() if str(t).strip())
    pass_types = p3.multiselect("Types « passifs » (marge dédiée)", types,
                                default=[t for t in PASSIVE_TYPES if t in types], key="order_passives")
    margin_p = p4.number_input("Marge passifs (%)", 0, 100, 10, 5, key="order_margin_p",
                               help="Les résistances / condensateurs se perdent facilement : prévois un peu de marge.")

    nb_unknown = int((check["Statut"] == "NON IDENTIFIÉ").sum())
    if nb_unknown:
        st.warning(f"{nb_unknown} ligne(s) du BOM ne sont pas rattachées à la librairie : elles ne sont pas "
                   "dans cette commande (résous-les avec « 🧩 à trancher » ou ajoute-les à la lib).")

    lines = build_order_lines(check, lib, prefer, margin / 100, margin_p / 100, pass_types)
    if lines.empty:
        st.success("🎉 Rien à commander : le stock couvre tous les besoins.")
        return

    shown = ["Inclure", "Fournisseur", "Composant", "Valeur", "Type", "Ref fournisseur", "Besoin", "Stock perso",
             "Manquant", "Qté à commander", "Prix unitaire (€)", "Coût estimé (€)", "Remarque"]
    editor_key = (f"order_editor_{prefer}_{margin}_{margin_p}_{'-'.join(pass_types)}_"
                  f"{st.session_state.lib_version}_{len(lines)}")
    edited = st.data_editor(
        lines, column_order=shown, hide_index=True, key=editor_key, height=420, **STRETCH,
        disabled=[c for c in lines.columns if c not in ("Inclure", "Qté à commander")],
        column_config={
            "Inclure": st.column_config.CheckboxColumn("Inclure"),
            "Qté à commander": st.column_config.NumberColumn("Qté à commander", min_value=0, step=1),
        },
    )
    st.caption("Tu peux décocher une ligne ou ajuster la quantité avant de télécharger.")

    lc, mo = to_lcsc_df(edited), to_mouser_df(edited)
    mask_lc = (edited["Inclure"] == True) & (edited["Fournisseur"] == "LCSC")  # noqa: E712
    cost_lc = (edited.loc[mask_lc, "Prix unitaire (€)"].fillna(0) * edited.loc[mask_lc, "Qté à commander"].fillna(0)).sum()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Lignes LCSC", len(lc))
    m2.metric("Lignes Mouser", len(mo))
    m3.metric("Pièces à commander", int(lc["Quantity"].sum() + mo["Quantity"].sum()))
    m4.metric("Coût estimé LCSC", f"{cost_lc:.2f} €")

    d1, d2, d3 = st.columns(3)
    d1.download_button("⬇️ Panier LCSC / JLCPCB (.csv)", data=df_to_csv_bytes(lc), file_name="commande_lcsc.csv",
                       mime="text/csv", disabled=lc.empty, **STRETCH)
    d2.download_button("⬇️ Panier Mouser (.csv)", data=df_to_csv_bytes(mo), file_name="commande_mouser.csv",
                       mime="text/csv", disabled=mo.empty, **STRETCH)
    d3.download_button("⬇️ Panier Mouser (.xlsx)", data=df_to_xlsx_bytes(mo), file_name="commande_mouser.xlsx",
                       mime=XLSX_MIME, disabled=mo.empty, **STRETCH)

    with st.expander("Comment importer ces fichiers ?"):
        st.markdown(
            "**LCSC / JLCPCB** : sur lcsc.com, outil *BOM* → *Upload a BOM File* (csv, xls ou xlsx, 800 lignes max). "
            "L'étape de correspondance demande de désigner les colonnes : **Quantity** (obligatoire) et "
            "**LCSC Part Number**. Les en-têtes sont ceux de l'export de commande LCSC.\n\n"
            "**Mouser** : sur mouser.fr, *Services & outils* → *BOM Tool* → importer un tableur (csv, xls ou xlsx), "
            "une référence et une quantité par ligne. Désigne **Mouser Part Number** et **Quantity** à l'étape de "
            "correspondance des colonnes. Si le csv s'ouvre mal chez toi, utilise la version .xlsx."
        )


# ============================================================
# ONGLET : Production (retirer du stock les PCB fabriqués)
# ============================================================

def tab_production():
    bom, lib = st.session_state.bom, st.session_state.lib
    raw = st.session_state.bom_raw
    names = list(dict.fromkeys(raw["BOM"]))
    planned = raw.drop_duplicates("BOM_id").set_index("BOM")["N_PCB"].to_dict()
    ver = st.session_state.get("prod_ver", 0)

    st.caption("Quand tu as fabriqué des PCB, valide-les ici : les composants utilisés sont retirés du « Stock Personnel » "
               "(jamais en dessous de 0) et le mouvement est noté dans « Historique stock ».")

    # --- ce que le stock permet de faire
    rows = []
    for name in names:
        b = max_buildable(bom, lib, name)
        rows.append({
            "BOM": name, "PCB prévus": int(planned.get(name, 1)),
            "PCB réalisables avec le stock": "?" if b["n"] is None else b["n"],
            "Composant limitant": b["limiting"],
            "Composants non vérifiables": b["unverified"],
        })
    st.markdown("##### Que permet le stock actuel ?")
    st.dataframe(pd.DataFrame(rows), hide_index=True, **STRETCH)
    st.caption("Calculé BOM par BOM : si deux BOM utilisent la même pièce, le stock n'est pas partagé dans ce tableau.")

    st.divider()
    st.markdown("##### Valider une production")
    c1, c2 = st.columns([2, 1])
    sel = c1.selectbox("BOM fabriquée", names, key=f"prod_sel_{ver}")
    n = c2.number_input("Nombre de PCB fabriqués", min_value=1, value=int(planned.get(sel, 1)), step=1,
                        key=f"prod_n_{sel}_{ver}")

    chk = production_check(bom, lib, sel, n)
    deduct, skipped = production_deductions(chk)
    n_insuf = int((deduct["Statut"] == "INSUFFISANT").sum())

    k1, k2, k3 = st.columns(3)
    k1.metric("Composants déduits", len(deduct))
    k2.metric("Pièces déduites", int(deduct[["Stock avant", "Besoin"]].min(axis=1).sum()))
    k3.metric("Stock insuffisant", n_insuf)

    def _color_prod(val):
        return {"OK": "background-color: #d4edda; color: #1b4332",
                "INSUFFISANT": "background-color: #f8d7da; color: #58151c"}.get(val, "")

    st.dataframe(deduct.drop(columns=["lib_idx"]).style.map(_color_prod, subset=["Statut"]),
                 hide_index=True, height=340, **STRETCH)

    if not skipped.empty:
        with st.expander(f"{len(skipped)} composant(s) non déduit(s) (absents de la lib ou non identifiés)"):
            st.dataframe(skipped[["Composant", "Valeur", "Ref LCSC", "Besoin", "Statut"]], hide_index=True, **STRETCH)

    force = True
    if n_insuf:
        st.warning(f"{n_insuf} composant(s) n'ont pas assez de stock : leur stock sera mis à 0 (pas de stock négatif). "
                   "Vérifie que ton stock est à jour (onglet 📥 Stock & prix).")
        force = st.checkbox("Valider quand même", key=f"prod_force_{ver}")
    confirm = st.checkbox(f"Je confirme avoir fabriqué {n} PCB « {sel} » : retirer ces pièces du stock",
                          key=f"prod_confirm_{ver}")

    if st.button("✅ Valider la production", type="primary", disabled=not (confirm and force and len(deduct))):
        label = f"Production – {sel} ×{n}"
        modified = apply_production(st.session_state.lib_path, lib, deduct, label)
        st.session_state.lib = pd.read_excel(st.session_state.lib_path)
        st.session_state.prod_ver = ver + 1       # remet les cases de confirmation à zéro
        st.session_state.stock_msg = f"Production validée ({label}) : {modified} composant(s) mis à jour."
        st.rerun()


# ============================================================
# ONGLET : Librairie (ajouter / modifier un composant)
# ============================================================

def _component_fields(prefix: str, defaults: dict) -> dict:
    """Formulaire d'un composant (un champ par colonne de la lib). Retourne {colonne: valeur}."""
    lib = st.session_state.lib
    cols = list(lib.columns)
    OTHER, NONE = "➕ Autre…", "(aucun)"
    values = {}

    def _d(col):
        v = defaults.get(col)
        return None if v is None or (not isinstance(v, str) and pd.isna(v)) else v

    # Réf LCSC + pré-remplissage depuis le site
    if "reference_LCSC" in cols:
        t1, t2 = st.columns([3, 1], vertical_alignment="bottom")
        values["reference_LCSC"] = t1.text_input("reference_LCSC", value=str(_d("reference_LCSC") or ""),
                                                 key=f"{prefix}_reference_LCSC")
        if t2.button("🔎 Pré-remplir depuis LCSC", key=f"{prefix}_fetch", **STRETCH,
                     help="Récupère prix, stock et, si la page le permet, réf fabricant / package / description. "
                          "Ne remplace pas ce que tu as déjà saisi."):
            ref = values["reference_LCSC"].strip()
            if not ref:
                st.warning("Saisis d'abord une réf LCSC.")
            else:
                with st.spinner("Lecture de la page LCSC..."):
                    info = get_lcsc_product_info(ref)
                if not info:
                    st.error("Page LCSC inaccessible (réseau ou protection anti-robot).")
                else:
                    filled = []
                    for k, col in {"mpn": "Manufacturer Ref", "manufacturer": "Manufacturer Part",
                                   "package": "Footprint", "description": "description"}.items():
                        key = f"{prefix}_{col}"
                        if k in info and col in cols and not str(st.session_state.get(key, _d(col) or "")).strip():
                            st.session_state[key] = info[k]
                            filled.append(col)
                    for k, col, cast in (("price", "Price", float), ("stock", "Stock website", int)):
                        if info.get(k) is not None and col in cols:
                            st.session_state[f"{prefix}_{col}"] = cast(info[k])
                            filled.append(col)
                    st.toast("Rempli : " + (", ".join(filled) if filled else "rien de nouveau"))

    grid = st.columns(3)
    others = [c for c in cols if c != "reference_LCSC"]
    for i, col in enumerate(others):
        box, key, d = grid[i % 3], f"{prefix}_{col}", _d(col)

        if col == "type":
            types = sorted(t for t in lib["type"].dropna().astype(str).unique() if t.strip())
            opts = [NONE] + types + [OTHER]
            cur = str(d) if d is not None else NONE
            choice = box.selectbox("type", opts, index=opts.index(cur) if cur in opts else 0, key=key)
            if choice == OTHER:
                choice = box.text_input("Nouveau type", key=f"{key}_new").strip() or NONE
            values[col] = None if choice == NONE else choice
        elif col == STOCK_COL:
            values[col] = box.number_input(col, min_value=0, step=1, value=int(d) if d is not None else 0, key=key)
        elif col == "Price":
            values[col] = box.number_input(col, min_value=0.0, step=0.001, format="%.5f",
                                           value=float(d) if d is not None else None, key=key)
        elif col in NUMERIC_COLUMNS:
            values[col] = box.number_input(col, min_value=0, step=1, value=int(d) if d is not None else None, key=key)
        else:
            values[col] = box.text_input(col, value=str(d) if d is not None else "", key=key)
    return values


def _library_edit_block():
    lib = st.session_state.lib
    ver = st.session_state.get("libedit_ver", 0)

    q = st.text_input("🔎 Rechercher le composant à modifier", key="libedit_q",
                      placeholder="réf fabricant, valeur, description, réf LCSC / Mouser…")
    search_cols = [c for c in ["Manufacturer Ref", "Value", "description", "Footprint", "type",
                               "reference_LCSC", "reference_Mouser"] if c in lib.columns]
    mask = pd.Series(True, index=lib.index)
    if q:
        mask = lib[search_cols].astype(str).apply(
            lambda c: c.str.contains(q, case=False, na=False, regex=False)).any(axis=1)
    cand = lib[mask].head(80)
    if cand.empty:
        st.info("Aucun composant ne correspond.")
        return

    def _lbl(i):
        r = lib.loc[i]
        return f"l.{i + 2} — {r.get('Manufacturer Ref', '')} — {r.get('Value', '')} — {r.get('Footprint', '')}"

    idx = st.selectbox(f"Composant ({int(mask.sum())} résultat(s))", list(cand.index), format_func=_lbl,
                       key="libedit_sel")
    st.divider()

    values = _component_fields(f"edit_{idx}_{ver}", lib.loc[idx].to_dict())

    dups = find_duplicates(lib, {k: v for k, v in values.items() if k != "Manufacturer Ref"}, exclude_idx=idx)
    if dups:
        st.warning("Une autre ligne a la même réf : " + ", ".join(f"l.{r} ({n}, {c})" for r, n, c in dups))

    if st.button("💾 Enregistrer les modifications", type="primary", key=f"libedit_save_{ver}"):
        changed = update_component(st.session_state.lib_path, idx, values)
        if changed:
            st.session_state.lib = pd.read_excel(st.session_state.lib_path)
            st.session_state.lib_version += 1
            st.session_state.libedit_ver = ver + 1
            st.session_state.stock_msg = f"Composant l.{idx + 2} modifié ({', '.join(changed)})."
            st.rerun()
        else:
            st.info("Aucune modification à enregistrer.")


def _library_add_block():
    lib = st.session_state.lib
    ver = st.session_state.get("libadd_ver", 0)
    prefix = f"add_{ver}"

    values = _component_fields(prefix, {STOCK_COL: 0})
    ok_name = bool((values.get("Manufacturer Ref") or "").strip())
    if not ok_name:
        st.caption("« Manufacturer Ref » est obligatoire.")

    dups = find_duplicates(lib, values)
    force = True
    if dups:
        st.warning("Ce composant semble déjà exister : " + ", ".join(f"l.{r} ({n}, {c})" for r, n, c in dups))
        force = st.checkbox("Ajouter quand même", key=f"{prefix}_force")

    if st.button("➕ Ajouter à la librairie", type="primary", disabled=not (ok_name and force), key=f"{prefix}_save"):
        row = add_component(st.session_state.lib_path, values)
        st.session_state.lib = pd.read_excel(st.session_state.lib_path)
        st.session_state.lib_version += 1
        st.session_state.libadd_ver = ver + 1        # formulaire vide pour le suivant
        st.session_state.stock_msg = f"Composant ajouté en ligne {row} : {values.get('Manufacturer Ref')}."
        st.rerun()


def tab_library():
    lib = st.session_state.lib
    st.caption("Ajoute ou corrige un composant de la librairie. Les modifications sont écrites dans le fichier "
               "(hyperliens conservés) ; relance la recherche pour qu'elles soient prises en compte dans tes BOM.")
    mode = st.radio("Action", ["✏️ Modifier un composant", "➕ Ajouter un composant"], horizontal=True,
                    label_visibility="collapsed", key="lib_mode")
    if mode.startswith("✏️"):
        _library_edit_block()
    else:
        _library_add_block()


# ============================================================
# Affichage de l'onglet sélectionné
# ============================================================

if st.session_state.get("stock_msg"):
    st.success(st.session_state.pop("stock_msg"))

if _is_locked(selected):
    lock_card(missing_for_bom_tabs())
elif selected == "results":
    tab_results()
elif selected == "components":
    tab_components()
elif selected == "order":
    tab_order()
elif selected == "production":
    tab_production()
elif selected == "report":
    tab_report()
elif selected == "library":
    tab_library()
else:
    tab_stock()