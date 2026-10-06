"""
Stock personnel : import des paniers/commandes Mouser et LCSC (JLCPCB) dans la
colonne "Stock Personnel" de la librairie, puis vérification du stock par
rapport à la liste de composants générée à partir des BOM.

La librairie est modifiée avec openpyxl (jamais réécrite via pandas) pour ne pas
perdre les hyperliens / la mise en forme. On travaille sur la 1re feuille, comme
le reste de l'appli (pd.read_excel).
"""

from datetime import datetime

import openpyxl
import pandas as pd

STOCK_COL = "Stock Personnel"
HISTORY_SHEET = "Historique stock"


# ============================================================
# Lecture des fichiers fournisseurs
# ============================================================

def _norm(x) -> str:
    return "" if pd.isna(x) else str(x).strip().upper()


def _find_col(columns, *needles):
    for c in columns:
        lc = str(c).lower()
        if any(n in lc for n in needles):
            return c
    return None


def _find_qty_col(columns):
    for c in columns:
        lc = str(c).lower().strip()
        if lc.startswith(("qté", "qte", "qty", "quantity", "order qty")):
            return c
    return None


def read_mouser_file(file, source_name: str = "") -> pd.DataFrame:
    """Export Mouser (.xls / .xlsx) : détail de commande ou panier."""
    raw = pd.read_excel(file, header=None)

    header_idx = None
    for i in range(min(len(raw), 20)):
        if raw.iloc[i].astype(str).str.lower().str.contains("mouser").any():
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Colonne 'N° Mouser' introuvable dans le fichier Mouser.")

    df = raw.iloc[header_idx + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[header_idx]]

    c_ref = _find_col(df.columns, "mouser")
    c_mpn = _find_col(df.columns, "fab", "mfr", "manufacturer part")
    c_qty = _find_qty_col(df.columns)
    c_desc = _find_col(df.columns, "desc")
    if c_ref is None or c_qty is None:
        raise ValueError("Colonnes référence Mouser / quantité introuvables.")

    out = pd.DataFrame({
        "Fournisseur": "Mouser",
        "Ref": df[c_ref].astype(str).str.strip(),
        "MPN": df[c_mpn] if c_mpn else "",
        "Description": df[c_desc] if c_desc else "",
        "Package": "",
        "Fabricant": "",
        "Qty": pd.to_numeric(df[c_qty], errors="coerce"),
    })
    out = out[df[c_ref].notna() & out["Qty"].notna() & (out["Qty"] > 0)].copy()
    out["Source"] = source_name
    return out.reset_index(drop=True)


def read_lcsc_file(file, source_name: str = "") -> pd.DataFrame:
    """Export LCSC / JLCPCB (.csv)."""
    df = pd.read_csv(file, sep=None, engine="python", encoding="utf-8-sig")

    c_ref = _find_col(df.columns, "lcsc part")
    c_mpn = _find_col(df.columns, "manufacture part", "manufacturer part")
    c_qty = _find_qty_col(df.columns)
    c_desc = _find_col(df.columns, "description")
    c_pkg = _find_col(df.columns, "package")
    c_maker = _find_col(df.columns, "manufacturer")
    if c_ref is None or c_qty is None:
        raise ValueError("Colonnes 'LCSC Part Number' / 'Quantity' introuvables.")

    out = pd.DataFrame({
        "Fournisseur": "LCSC",
        "Ref": df[c_ref].astype(str).str.strip(),
        "MPN": df[c_mpn] if c_mpn else "",
        "Description": df[c_desc] if c_desc else "",
        "Package": df[c_pkg] if c_pkg else "",
        "Fabricant": df[c_maker] if c_maker else "",
        "Qty": pd.to_numeric(df[c_qty], errors="coerce"),
    })
    out = out[df[c_ref].notna() & out["Qty"].notna() & (out["Qty"] > 0)].copy()
    out["Source"] = source_name
    return out.reset_index(drop=True)


# ============================================================
# Rapprochement avec la librairie
# ============================================================

def _lib_stock(lib: pd.DataFrame, idx) -> float:
    if STOCK_COL not in lib.columns:
        return 0.0
    v = pd.to_numeric(lib.loc[idx, STOCK_COL], errors="coerce")
    return 0.0 if pd.isna(v) else float(v)


def _index(lib: pd.DataFrame, col: str) -> dict:
    """clé normalisée -> liste des index de lib (la lib peut contenir des doublons)."""
    d = {}
    if col not in lib.columns:
        return d
    for idx, v in lib[col].items():
        k = _norm(v)
        if k:
            d.setdefault(k, []).append(idx)
    return d


def _best(lib: pd.DataFrame, idxs):
    """Parmi des doublons : la 1re ligne qui a déjà un stock renseigné, sinon la 1re."""
    if not idxs:
        return None
    if STOCK_COL in lib.columns:
        for i in idxs:
            v = pd.to_numeric(lib.loc[i, STOCK_COL], errors="coerce")
            if pd.notna(v) and v > 0:
                return i
    return idxs[0]


def plan_stock_import(lib: pd.DataFrame, entries: pd.DataFrame) -> pd.DataFrame:
    """
    Prévisualise l'import : pour chaque ligne de commande, quelle ligne de la
    librairie est trouvée et quel sera le stock avant / après.
    Rien n'est écrit. Le stock est mis sur la PREMIÈRE ligne de lib trouvée
    (évite de compter deux fois si la lib contient un doublon).
    """
    by_lcsc, by_mouser = _index(lib, "reference_LCSC"), _index(lib, "reference_Mouser")
    by_mpn = _index(lib, "Manufacturer Ref")

    current = {}
    rows = []
    for _, e in entries.iterrows():
        ref_idx = _best(lib, (by_lcsc if e["Fournisseur"] == "LCSC" else by_mouser).get(_norm(e["Ref"])))
        method = "Réf fournisseur"
        lib_idx = ref_idx
        if lib_idx is None:
            lib_idx = _best(lib, by_mpn.get(_norm(e["MPN"])))
            method = "Réf fabricant"

        if lib_idx is None:
            # Piste : une ligne de lib dont la réf fabricant ressemble (l'une contient l'autre)
            k, suggestion = _norm(e["MPN"]), ""
            if len(k) >= 5:
                for i2, r2 in lib.iterrows():
                    v = _norm(r2.get("Manufacturer Ref"))
                    if len(v) >= 5 and (v in k or k in v):
                        suggestion = f"l.{int(i2) + 2} : {r2.get('Manufacturer Ref')}"
                        break
            rows.append({**e.to_dict(), "lib_idx": pd.NA, "Ligne lib (Excel)": pd.NA,
                         "Composant lib": "", "Trouvé par": "", "Stock avant": pd.NA,
                         "Stock après": pd.NA, "Suggestion lib": suggestion,
                         "Statut": "NON TROUVÉ DANS LA LIB"})
            continue

        before = current.get(lib_idx, _lib_stock(lib, lib_idx))
        after = before + float(e["Qty"])
        current[lib_idx] = after
        rows.append({**e.to_dict(), "lib_idx": lib_idx, "Ligne lib (Excel)": int(lib_idx) + 2,
                     "Composant lib": lib.loc[lib_idx, "Manufacturer Ref"], "Trouvé par": method,
                     "Stock avant": before, "Stock après": after, "Suggestion lib": "", "Statut": "OK"})
    return pd.DataFrame(rows)


# ============================================================
# Écriture dans la librairie (openpyxl)
# ============================================================

def _ws(wb):
    return wb.worksheets[0]


def _headers(ws) -> dict:
    return {str(c.value).strip().lower(): c.column for c in ws[1] if c.value}


def _num(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _as_int_if_whole(x: float):
    return int(x) if float(x).is_integer() else x


def already_imported(lib_path, source_names) -> dict:
    """Sources déjà présentes dans l'historique : {nom_fichier: date}."""
    wb = openpyxl.load_workbook(lib_path, read_only=True)
    if HISTORY_SHEET not in wb.sheetnames:
        return {}
    found = {}
    for row in wb[HISTORY_SHEET].iter_rows(min_row=2, values_only=True):
        if row and row[1] in source_names:
            found[row[1]] = row[0]
    return found


def _open_for_stock(lib_path):
    """Ouvre la lib, crée la colonne Stock Personnel et la feuille d'historique si besoin."""
    wb = openpyxl.load_workbook(lib_path)
    ws = _ws(wb)
    heads = _headers(ws)

    if STOCK_COL.lower() not in heads:
        col = ws.max_column + 1
        ws.cell(row=1, column=col, value=STOCK_COL)
        heads[STOCK_COL.lower()] = col

    if HISTORY_SHEET not in wb.sheetnames:
        h = wb.create_sheet(HISTORY_SHEET)
        h.append(["Date", "Source", "Fournisseur", "Ref", "Composant", "Ligne lib", "Qté ajoutée", "Stock avant", "Stock après"])
    return wb, ws, heads, heads[STOCK_COL.lower()], wb[HISTORY_SHEET]


def apply_stock_import(lib_path, plan: pd.DataFrame, add_missing: bool = False) -> int:
    """
    Ajoute les quantités du plan dans "Stock Personnel" (crée la colonne si
    besoin), journalise dans la feuille "Historique stock".
    Si add_missing, les composants introuvables sont ajoutés en bas de la lib.
    Retourne le nombre de lignes de stock modifiées / ajoutées.
    """
    wb, ws, heads, col_stock, hs = _open_for_stock(lib_path)

    # dernière ligne réellement remplie (ws.max_row compte les lignes vides formatées)
    c_name = heads.get("manufacturer ref", 1)
    last = 1
    for r in range(ws.max_row, 1, -1):
        if ws.cell(row=r, column=c_name).value not in (None, ""):
            last = r
            break

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    count = 0

    for _, p in plan.iterrows():
        qty = float(p["Qty"])

        if pd.notna(p["lib_idx"]):
            row = int(p["lib_idx"]) + 2
            cell = ws.cell(row=row, column=col_stock)
            before = _num(cell.value)
            cell.value = _as_int_if_whole(before + qty)
            name = p["Composant lib"]
        elif add_missing:
            last += 1
            row = last
            before = 0.0
            values = {
                "manufacturer ref": p["MPN"] or p["Ref"],
                "description": p["Description"],
                "footprint": p["Package"],
                "manufacturer part": p.get("Fabricant", ""),
                "reference_lcsc": p["Ref"] if p["Fournisseur"] == "LCSC" else None,
                "reference_mouser": p["Ref"] if p["Fournisseur"] == "Mouser" else None,
            }
            for h, v in values.items():
                if h in heads and v not in (None, "") and not pd.isna(v):
                    ws.cell(row=row, column=heads[h], value=v)
            ws.cell(row=row, column=col_stock, value=_as_int_if_whole(qty))
            name = p["MPN"] or p["Ref"]
        else:
            continue

        hs.append([now, p["Source"], p["Fournisseur"], p["Ref"], name, row,
                   _as_int_if_whole(qty), _as_int_if_whole(before), _as_int_if_whole(before + qty)])
        count += 1

    wb.save(lib_path)
    return count


def set_stock_manual(lib_path, lib: pd.DataFrame, new_values: dict, source: str = "Saisie manuelle") -> int:
    """
    Saisie manuelle : écrit un stock ABSOLU pour des lignes de lib.
    new_values : {index de ligne dans lib (DataFrame) : nouveau stock}
    Journalise l'écart dans l'historique. Retourne le nombre de lignes modifiées.
    """
    wb, ws, heads, col_stock, hs = _open_for_stock(lib_path)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    count = 0

    for idx, new in new_values.items():
        row = int(idx) + 2
        cell = ws.cell(row=row, column=col_stock)
        before = _num(cell.value)
        new = _num(new)
        if new == before:
            continue
        cell.value = _as_int_if_whole(new)

        ref = lib.loc[idx, "reference_LCSC"] if "reference_LCSC" in lib.columns else None
        if pd.isna(ref) or not str(ref).strip():
            ref = lib.loc[idx, "reference_Mouser"] if "reference_Mouser" in lib.columns else ""
        hs.append([now, source, "", "" if pd.isna(ref) else ref, lib.loc[idx, "Manufacturer Ref"], row,
                   _as_int_if_whole(new - before), _as_int_if_whole(before), _as_int_if_whole(new)])
        count += 1

    wb.save(lib_path)
    return count


# ============================================================
# Vérification du stock vs liste de composants des BOM
# ============================================================

def build_stock_check(components: pd.DataFrame, lib: pd.DataFrame) -> pd.DataFrame:
    """
    components : sortie de build_component_list (colonnes Ref LCSC, Ref Mouser, Qté totale...)
    lib        : librairie relue avec la colonne "Stock Personnel"

    Retourne une ligne par composant avec : besoin, stock, manquant, statut,
    prix unitaire et coût du manquant.
    """
    by_lcsc, by_mouser = _index(lib, "reference_LCSC"), _index(lib, "reference_Mouser")

    rows = []
    for _, c in components.iterrows():
        idx = _best(lib, by_lcsc.get(_norm(c["Ref LCSC"]))) if _norm(c["Ref LCSC"]) else None
        if idx is None and _norm(c["Ref Mouser"]):
            idx = _best(lib, by_mouser.get(_norm(c["Ref Mouser"])))

        need = int(c["Qté totale"])
        unknown = c["Statut"] != "OK"   # ligne de BOM non rattachée à la lib : on ne peut pas connaître son stock
        stock = _lib_stock(lib, idx) if idx is not None else 0.0
        missing = max(need - stock, 0)

        price = pd.to_numeric(lib.loc[idx, "Price"], errors="coerce") if (idx is not None and "Price" in lib.columns) else float("nan")

        if unknown:
            statut = "NON IDENTIFIÉ"
        elif idx is None:
            statut = "ABSENT DE LA LIB"
        elif missing > 0:
            statut = "MANQUE" if stock > 0 else "RIEN EN STOCK"
        else:
            statut = "OK"

        rows.append({
            "Composant": c["Composant"], "Valeur": c["Valeur"], "Description": c.get("Description", ""),
            "Type": c["Type"], "Footprint": c["Footprint"],
            "Ref LCSC": c["Ref LCSC"], "Ref Mouser": c["Ref Mouser"],
            "BOM concernées": c["BOM concernées"], "Qté par BOM": c.get("Qté par BOM", ""),
            "Besoin": need,
            "Stock perso": None if unknown else _as_int_if_whole(stock),
            "Manquant": None if unknown else _as_int_if_whole(missing),
            "Ligne lib (Excel)": None if idx is None else int(idx) + 2,
            "Prix unitaire (€)": None if pd.isna(price) else float(price),
            "Coût du manquant (€)": None if pd.isna(price) else round(float(price) * missing, 2),
            "Statut": statut,
        })

    out = pd.DataFrame(rows)
    order = {"RIEN EN STOCK": 0, "MANQUE": 1, "ABSENT DE LA LIB": 2, "NON IDENTIFIÉ": 3, "OK": 4}
    return out.sort_values("Statut", key=lambda s: s.map(order), kind="stable").reset_index(drop=True)


def export_stock_check_excel(check: pd.DataFrame, output_path, per_bom: pd.DataFrame = None):
    """Feuilles : 'Composants & stock' (tout), 'A commander' (ce qui manque), 'Detail par BOM'."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        check.to_excel(writer, sheet_name="Composants & stock", index=False)
        check[check["Manquant"].fillna(0) > 0].to_excel(writer, sheet_name="A commander", index=False)
        if per_bom is not None:
            per_bom.to_excel(writer, sheet_name="Detail par BOM", index=False)


# ============================================================
# Production : "j'ai fabriqué n PCB" -> on retire les pièces du stock
# ============================================================

def production_check(bom: pd.DataFrame, lib: pd.DataFrame, bom_name: str, n_built: int) -> pd.DataFrame:
    """Besoin / stock par composant pour fabriquer n_built PCB de la BOM `bom_name`."""
    from backend.BOM_components import build_component_list

    sub = bom[bom["BOM"] == bom_name].copy()
    sub["N_PCB"] = int(n_built)
    comps, _ = build_component_list(sub, lib)
    return build_stock_check(comps, lib)


def max_buildable(bom: pd.DataFrame, lib: pd.DataFrame, bom_name: str) -> dict:
    """
    Combien de PCB de cette BOM le stock actuel permet de faire (indépendamment des autres BOM)
    et quel composant limite. Seuls les composants retrouvés dans la lib sont comptés.
    """
    chk = production_check(bom, lib, bom_name, 1)
    tracked = chk[chk["Ligne lib (Excel)"].notna() & (chk["Besoin"] > 0)]
    unverified = len(chk) - len(tracked)
    if tracked.empty:
        return {"n": None, "limiting": "", "unverified": unverified}

    ratio = (tracked["Stock perso"].astype(float) // tracked["Besoin"].astype(float)).astype(int)
    i = ratio.idxmin()
    limiting = tracked.loc[i, "Composant"] or tracked.loc[i, "Valeur"]
    return {"n": int(ratio.min()), "limiting": limiting, "unverified": unverified}


def production_deductions(check: pd.DataFrame):
    """
    Sépare les lignes qu'on peut déduire du stock (retrouvées dans la lib) des autres.
    Retourne (deduct, skipped). deduct a : Composant, Valeur, Ref LCSC, Besoin, Stock avant,
    Stock après, Statut (OK / INSUFFISANT), lib_idx.
    """
    ok_mask = check["Ligne lib (Excel)"].notna() & (check["Statut"] != "NON IDENTIFIÉ")
    deduct = check[ok_mask].copy()
    skipped = check[~ok_mask].copy()

    deduct["lib_idx"] = deduct["Ligne lib (Excel)"].astype(int) - 2
    deduct["Stock avant"] = deduct["Stock perso"].astype(float)
    deduct["Stock après"] = (deduct["Stock avant"] - deduct["Besoin"]).clip(lower=0)
    deduct["Statut"] = (deduct["Stock avant"] >= deduct["Besoin"]).map({True: "OK", False: "INSUFFISANT"})
    cols = ["Composant", "Valeur", "Ref LCSC", "Besoin", "Stock avant", "Stock après", "Statut", "lib_idx"]
    return deduct[cols].reset_index(drop=True), skipped.reset_index(drop=True)


def apply_production(lib_path, lib: pd.DataFrame, deduct: pd.DataFrame, label: str) -> int:
    """Retire les quantités du stock (jamais en dessous de 0) et journalise avec `label`."""
    g = deduct.groupby("lib_idx").agg(need=("Besoin", "sum"), stock=("Stock avant", "first"))
    new_values = {int(idx): max(r["stock"] - r["need"], 0) for idx, r in g.iterrows()}
    return set_stock_manual(lib_path, lib, new_values, source=label)