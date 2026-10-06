"""
Génération des listes de commande (ce qui manque en stock) pour LCSC / JLCPCB et Mouser.

- LCSC : l'outil BOM de LCSC accepte .csv / .xls / .xlsx (jusqu'à 800 lignes) ; la colonne
  Quantité est obligatoire, avec la réf LCSC et/ou la réf fabricant. On reprend les en-têtes
  de l'export de commande LCSC pour qu'ils soient reconnus tels quels.
- Mouser : l'outil BOM de Mouser accepte .xls / .xlsx / .csv, avec UNE réf (Mouser ou
  fabricant) et UNE quantité par ligne.
Dans les deux cas, le site demande ensuite d'indiquer quelle colonne est quoi.
"""

import io
import math

import pandas as pd

PASSIVE_TYPES = ["Resistor", "Capacitor", "Inductor"]

LCSC_COLUMNS = ["LCSC Part Number", "Manufacture Part Number", "Manufacturer",
                "Customer NO.", "Package", "Description", "Quantity"]
MOUSER_COLUMNS = ["Mouser Part Number", "Manufacturer Part Number", "Quantity", "Description"]

LINE_COLUMNS = ["Inclure", "Fournisseur", "Composant", "Valeur", "Type", "Footprint", "Ref fournisseur",
                "Besoin", "Stock perso", "Manquant", "Qté à commander",
                "Prix unitaire (€)", "Coût estimé (€)", "Remarque",
                "MPN", "Fabricant", "Package", "Description"]


def _clean(x) -> str:
    return "" if pd.isna(x) else str(x).strip()


def build_order_lines(check: pd.DataFrame, lib: pd.DataFrame, prefer: str = "LCSC",
                      margin: float = 0.0, margin_passive: float = 0.0, passive_types=()) -> pd.DataFrame:
    """
    Une ligne par composant dont le stock ne couvre pas le besoin.

    prefer         : "LCSC" ou "Mouser" (l'autre sert de repli si la réf préférée manque)
    margin         : marge générale (0.1 = +10 %) ajoutée au manquant
    margin_passive : marge pour les types listés dans passive_types (résistances, condensateurs...)
    """
    todo = check[check["Manquant"].fillna(0) > 0]
    rows = []

    for _, c in todo.iterrows():
        lcsc, mouser = _clean(c["Ref LCSC"]), _clean(c["Ref Mouser"])
        order = [("Mouser", mouser), ("LCSC", lcsc)] if prefer == "Mouser" else [("LCSC", lcsc), ("Mouser", mouser)]
        supplier, ref = next(((s, r) for s, r in order if r), (None, ""))

        m = margin_passive if c["Type"] in passive_types else margin
        qty = int(math.ceil(round(float(c["Manquant"]) * (1 + m), 6)))

        maker = ""
        if pd.notna(c["Ligne lib (Excel)"]) and "Manufacturer Part" in lib.columns:
            maker = _clean(lib.loc[int(c["Ligne lib (Excel)"]) - 2, "Manufacturer Part"])

        price = c["Prix unitaire (€)"] if supplier == "LCSC" else None   # le prix de la lib est un prix LCSC
        remarque = []
        if supplier is None:
            remarque.append("Aucune réf fournisseur")
        if c["Statut"] == "ABSENT DE LA LIB":
            remarque.append("Absent de la lib (stock inconnu)")

        rows.append({
            "Inclure": supplier is not None,
            "Fournisseur": supplier or "—",
            "Composant": _clean(c["Composant"]), "Valeur": _clean(c["Valeur"]),
            "Type": _clean(c["Type"]), "Footprint": _clean(c["Footprint"]),
            "Ref fournisseur": ref,
            "Besoin": int(c["Besoin"]),
            "Stock perso": c["Stock perso"],
            "Manquant": c["Manquant"],
            "Qté à commander": qty,
            "Prix unitaire (€)": None if price is None or pd.isna(price) else float(price),
            "Coût estimé (€)": None if price is None or pd.isna(price) else round(float(price) * qty, 2),
            "Remarque": " | ".join(remarque),
            "MPN": _clean(c["Composant"]), "Fabricant": maker,
            "Package": _clean(c["Footprint"]), "Description": _clean(c["Description"]),
        })

    return pd.DataFrame(rows, columns=LINE_COLUMNS)


def _selected(lines: pd.DataFrame, supplier: str) -> pd.DataFrame:
    qty = pd.to_numeric(lines["Qté à commander"], errors="coerce").fillna(0)
    return lines[(lines["Inclure"] == True) & (lines["Fournisseur"] == supplier) & (qty > 0)]  # noqa: E712


def to_lcsc_df(lines: pd.DataFrame) -> pd.DataFrame:
    sel = _selected(lines, "LCSC")
    return pd.DataFrame({
        "LCSC Part Number": sel["Ref fournisseur"],
        "Manufacture Part Number": sel["MPN"],
        "Manufacturer": sel["Fabricant"],
        "Customer NO.": "",
        "Package": sel["Package"],
        "Description": sel["Description"],
        "Quantity": pd.to_numeric(sel["Qté à commander"]).astype(int),
    }, columns=LCSC_COLUMNS).reset_index(drop=True)


def to_mouser_df(lines: pd.DataFrame) -> pd.DataFrame:
    sel = _selected(lines, "Mouser")
    return pd.DataFrame({
        "Mouser Part Number": sel["Ref fournisseur"],
        "Manufacturer Part Number": sel["MPN"],
        "Quantity": pd.to_numeric(sel["Qté à commander"]).astype(int),
        "Description": sel["Description"],
    }, columns=MOUSER_COLUMNS).reset_index(drop=True)


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """CSV UTF-8, virgule, fins de ligne CRLF (comme l'export de commande LCSC)."""
    buf = io.StringIO()
    df.to_csv(buf, index=False, lineterminator="\r\n")
    return buf.getvalue().encode("utf-8")


def df_to_xlsx_bytes(df: pd.DataFrame, sheet: str = "BOM") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)
    return buf.getvalue()
