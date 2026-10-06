"""
Ajout / modification d'un composant dans la librairie (Excel), via openpyxl pour conserver
les hyperliens et la mise en forme. On travaille sur la 1re feuille, comme le reste de l'appli.
La ligne Excel d'un composant = index du DataFrame + 2 (en-tête en ligne 1).
"""

from datetime import datetime

import pandas as pd

from backend.Stock_personnel import STOCK_COL, _open_for_stock, _as_int_if_whole

LCSC_URL = "https://www.lcsc.com/product-detail/{ref}.html"
NUMERIC_COLUMNS = {"Price", "Stock website", STOCK_COL}


def _clean_value(col: str, val):
    """Valeur prête à écrire : None si vide, nombre pour les colonnes numériques."""
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return None
    if col in NUMERIC_COLUMNS:
        if isinstance(val, str) and not val.strip():
            return None
        return _as_int_if_whole(float(val))
    val = str(val).strip()
    return val or None


def _same(old, new) -> bool:
    if old in (None, "") and new is None:
        return True
    if isinstance(new, (int, float)) and not isinstance(new, bool):
        try:
            return float(old) == float(new)
        except (TypeError, ValueError):
            return False
    return str(old).strip() == str(new).strip() if new is not None and old is not None else False


def _write(ws, row, col_idx, col_name, value):
    cell = ws.cell(row=row, column=col_idx)
    cell.value = value
    if col_name.lower() == "reference_lcsc":
        cell.hyperlink = LCSC_URL.format(ref=value) if value else None


def find_duplicates(lib: pd.DataFrame, values: dict, exclude_idx=None) -> list:
    """Composants déjà présents avec la même réf LCSC / Mouser / fabricant (hors exclude_idx)."""
    found = []
    for col in ("reference_LCSC", "reference_Mouser", "Manufacturer Ref"):
        v = values.get(col)
        if col not in lib.columns or v is None or not str(v).strip():
            continue
        same = lib[lib[col].astype(str).str.strip().str.upper() == str(v).strip().upper()]
        for idx, r in same.iterrows():
            if idx != exclude_idx:
                found.append((int(idx) + 2, r.get("Manufacturer Ref"), col))
    return found


def update_component(lib_path, idx, values: dict) -> list:
    """Modifie la ligne `idx` : n'écrit que les cellules qui changent. Retourne les colonnes modifiées."""
    wb, ws, heads, col_stock, hs = _open_for_stock(lib_path)
    row = int(idx) + 2
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    changed = []

    name_col = heads.get("manufacturer ref")
    name = ws.cell(row=row, column=name_col).value if name_col else ""

    for col, raw in values.items():
        col_idx = heads.get(col.strip().lower())
        if col_idx is None:
            continue
        old = ws.cell(row=row, column=col_idx).value
        new = _clean_value(col, raw)
        if _same(old, new):
            continue

        _write(ws, row, col_idx, col, new)
        changed.append(col)

        if col == STOCK_COL:
            before = float(old) if isinstance(old, (int, float)) else 0.0
            after = float(new) if new is not None else 0.0
            ref = ws.cell(row=row, column=heads["reference_lcsc"]).value if "reference_lcsc" in heads else ""
            hs.append([now, "Modification composant", "", ref or "", values.get("Manufacturer Ref") or name, row,
                       _as_int_if_whole(after - before), _as_int_if_whole(before), _as_int_if_whole(after)])

    wb.save(lib_path)
    return changed


def add_component(lib_path, values: dict) -> int:
    """Ajoute un composant en bas de la lib. Retourne son numéro de ligne Excel."""
    wb, ws, heads, col_stock, hs = _open_for_stock(lib_path)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    c_name = heads.get("manufacturer ref", 1)
    last = 1
    for r in range(ws.max_row, 1, -1):          # ws.max_row compte les lignes vides formatées
        if ws.cell(row=r, column=c_name).value not in (None, ""):
            last = r
            break
    row = last + 1

    for col, raw in values.items():
        col_idx = heads.get(col.strip().lower())
        new = _clean_value(col, raw)
        if col_idx is None or new is None:
            continue
        _write(ws, row, col_idx, col, new)

    stock = _clean_value(STOCK_COL, values.get(STOCK_COL))
    if stock:
        ref = values.get("reference_LCSC") or values.get("reference_Mouser") or ""
        hs.append([now, "Création composant", "", ref, values.get("Manufacturer Ref"), row,
                   stock, 0, stock])

    wb.save(lib_path)
    return row
