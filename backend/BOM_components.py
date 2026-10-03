"""
Liste consolidée des composants sur plusieurs BOM.

Regroupe les lignes identiques de toutes les BOM (même réf LCSC, sinon même
réf Mouser, sinon même Manufacturer Ref / Value / Footprint) et calcule la
quantité totale = quantité par PCB x nombre de PCB de chaque BOM.

Le BOM passé doit contenir (en plus des colonnes habituelles) :
  - "BOM"   : nom donné à la BOM
  - "N_PCB" : nombre de PCB à fabriquer pour cette BOM
"""

import re

import pandas as pd

from backend.BOM_report import _get_quantity_column

_MOUSER_RE = re.compile(r"MOUSER:\s*([^\s|]+)", re.IGNORECASE)

COLS_COMPONENTS = [
    "Composant", "Valeur", "Description", "Ref LCSC", "Ref Mouser",
    "BOM concernées", "Qté totale", "Footprint", "Type", "Qté par BOM", "Statut",
]


def _clean(x) -> str:
    return "" if pd.isna(x) else str(x).strip()


def _first(series):
    for v in series:
        if v != "":
            return v
    return ""


def build_component_list(bom: pd.DataFrame, lib: pd.DataFrame):
    """
    Returns
    -------
    components : DataFrame, une ligne par composant distinct (voir COLS_COMPONENTS)
    per_bom    : DataFrame, une ligne par (composant, BOM) avec la quantité
    """
    bom = bom.copy()

    qty_col = _get_quantity_column(bom)
    if qty_col:
        qty = pd.to_numeric(bom[qty_col], errors="coerce").fillna(1)
    else:
        qty = pd.Series(1, index=bom.index)

    if "N_PCB" in bom.columns:
        n_pcb = pd.to_numeric(bom["N_PCB"], errors="coerce").fillna(1)
    else:
        n_pcb = 1

    bom["_qty_total"] = qty * n_pcb
    if "BOM" not in bom.columns:
        bom["BOM"] = "BOM"

    # Index de la librairie par réf LCSC / Mouser (première occurrence)
    by_lcsc, by_mouser = {}, {}
    for _, r in lib.iterrows():
        l, m = _clean(r.get("reference_LCSC")), _clean(r.get("reference_Mouser"))
        if l:
            by_lcsc.setdefault(l, r)
        if m:
            by_mouser.setdefault(m, r)

    rows = []
    for _, line in bom.iterrows():
        lcsc = _clean(line.get("LCSC_part_number"))
        found = _MOUSER_RE.search(_clean(line.get("Comments")))
        mouser_comment = found.group(1) if found else ""

        lib_row = by_lcsc.get(lcsc) if lcsc else None
        if lib_row is None and mouser_comment:
            lib_row = by_mouser.get(mouser_comment)

        def pick(lib_col, bom_col):
            v = _clean(lib_row.get(lib_col)) if lib_row is not None else ""
            return v or _clean(line.get(bom_col))

        if lib_row is not None:
            lcsc = lcsc or _clean(lib_row.get("reference_LCSC"))
        mouser = (_clean(lib_row.get("reference_Mouser")) if lib_row is not None else "") or mouser_comment

        composant = pick("Manufacturer Ref", "Manufacturer Ref")
        valeur = pick("Value", "Value")
        footprint = pick("Footprint", "Footprint")
        description = pick("description", "Description")
        type_ = pick("type", "Description")

        if lcsc:
            key = f"LCSC:{lcsc}"
        elif mouser:
            key = f"MOUSER:{mouser}"
        else:
            key = "RAW:" + "|".join(s.lower() for s in (composant, valeur, footprint))

        rows.append({
            "key": key,
            "Composant": composant, "Valeur": valeur, "Description": description,
            "Ref LCSC": lcsc, "Ref Mouser": mouser,
            "Footprint": footprint, "Type": type_,
            "BOM": _clean(line.get("BOM")),
            "Qty": float(line["_qty_total"]),
            "Identifié": bool(lcsc or mouser),
        })

    df = pd.DataFrame(rows)

    # Quantité par (composant, BOM)
    per_bom = df.groupby(["key", "BOM"], as_index=False, sort=False).agg(Qty=("Qty", "sum"))

    # Une ligne par composant
    comp = df.groupby("key", sort=False).agg(
        Composant=("Composant", _first),
        Valeur=("Valeur", _first),
        Description=("Description", _first),
        **{"Ref LCSC": ("Ref LCSC", _first), "Ref Mouser": ("Ref Mouser", _first)},
        Footprint=("Footprint", _first),
        Type=("Type", _first),
        Qty=("Qty", "sum"),
        Identifié=("Identifié", "all"),
    )

    per_bom["_txt"] = per_bom["BOM"] + " (" + per_bom["Qty"].round().astype(int).astype(str) + ")"
    comp["BOM concernées"] = per_bom.groupby("key", sort=False)["BOM"].agg(", ".join)
    comp["Qté par BOM"] = per_bom.groupby("key", sort=False)["_txt"].agg(" | ".join)
    comp["Qté totale"] = comp["Qty"].round().astype(int)
    comp["Statut"] = comp["Identifié"].map({True: "OK", False: "NON IDENTIFIÉ"})

    comp = comp.reset_index()
    components = (
        comp[["key"] + COLS_COMPONENTS]
        .sort_values(["Type", "Composant"], kind="stable")
    )

    per_bom = per_bom.merge(
        components[["key", "Composant", "Ref LCSC", "Ref Mouser"]], on="key", how="left"
    )
    per_bom["Qté"] = per_bom["Qty"].round().astype(int)
    per_bom = per_bom[["BOM", "Composant", "Ref LCSC", "Ref Mouser", "Qté"]].sort_values(
        ["BOM", "Composant"], kind="stable"
    )

    return components.drop(columns="key").reset_index(drop=True), per_bom.reset_index(drop=True)


def export_component_list_excel(components, per_bom, output_path):
    """Exporte la liste consolidée (feuille 'Composants') + le détail par BOM."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        components.to_excel(writer, sheet_name="Composants", index=False)
        per_bom.to_excel(writer, sheet_name="Detail par BOM", index=False)