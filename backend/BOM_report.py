"""
Fonction qui donne un rapport final sur le BOM : prix pour n PCB des
composants, composants incomplets/non détectés, et analyses (composant le
plus cher, coût par famille, composants les plus récurrents).
"""

import pandas as pd


def _get_quantity_column(bom):
    """Essaie de détecter automatiquement la colonne de quantité par PCB."""
    for col in ["Quantity", "Qty", "Qte", "Quantité", "Quantity Per PCB"]:
        if col in bom.columns:
            return col
    return None


def _get_status(row):
    lcsc = row.get("LCSC_part_number")
    comments = str(row.get("Comments", "")) if pd.notna(row.get("Comments")) else ""

    if pd.isna(lcsc) or str(lcsc).strip() == "":
        return "NOT DETECTED"
    if "REFERENCE MISMATCH" in comments:
        return "REFERENCE MISMATCH"
    if "ERROR IN PART NUMBER" in comments:
        return "ERROR"
    if pd.isna(row.get("Unit_price")):
        return "PRICE UNAVAILABLE"
    return "OK"


def build_bom_report(bom, lib, n_pcb=1, quantity_column=None, top_n=10):
    """
    Construit un rapport BOM.

    Parameters
    ----------
    bom : DataFrame du BOM (avec LCSC_part_number, Comments, etc.)
    lib : DataFrame de la librairie RELUE APRES Update_Price_Stock
          (doit contenir reference_LCSC, Price, Stock website, idéalement 'type')
    n_pcb : nombre de PCB pour lesquels calculer le prix total
    quantity_column : nom de la colonne "quantité par PCB" dans le BOM.
                       Si None, détection automatique (voir _get_quantity_column).
    top_n : nombre de lignes à garder dans les classements (composants les
            plus chers / les plus récurrents).

    Returns
    -------
    dict avec :
        "detail"          : DataFrame ligne par ligne (prix unitaire, prix total, statut)
        "missing"         : DataFrame des composants incomplets / non détectés / à vérifier
        "total_price"     : prix total du BOM pour n_pcb (EUR)
        "top_expensive"   : DataFrame des composants les plus chers (Line_total_price)
        "cost_by_type"    : Series du coût total par famille ("type"), triée décroissant
        "most_recurring"  : DataFrame des composants les plus utilisés (par quantité totale)
    """

    bom = bom.copy()

    if quantity_column is None:
        quantity_column = _get_quantity_column(bom)

    # ------------------------------------------------------------
    # Jointure avec la librairie pour récupérer Price / Stock / type
    # ------------------------------------------------------------
    colonnes_lib = [c for c in ["reference_LCSC", "Price", "Stock website", "type"] if c in lib.columns]
    lib_infos = lib[colonnes_lib].drop_duplicates(subset="reference_LCSC")

    merged = bom.merge(
        lib_infos,
        left_on="LCSC_part_number",
        right_on="reference_LCSC",
        how="left",
    )

    # ------------------------------------------------------------
    # Quantité par PCB / quantité totale
    # ------------------------------------------------------------
    if quantity_column and quantity_column in merged.columns:
        qty_per_board = pd.to_numeric(merged[quantity_column], errors="coerce").fillna(1)
    else:
        qty_per_board = pd.Series(1, index=merged.index)

    merged["Qty_per_PCB"] = qty_per_board
    merged["Qty_total"] = qty_per_board * n_pcb

    merged["Unit_price"] = pd.to_numeric(merged.get("Price"), errors="coerce")
    merged["Line_total_price"] = merged["Unit_price"] * merged["Qty_total"]

    # ------------------------------------------------------------
    # Statut de chaque ligne
    # ------------------------------------------------------------
    merged["Status"] = merged.apply(_get_status, axis=1)

    missing = merged[merged["Status"] != "OK"].copy()
    total_price = merged["Line_total_price"].sum(skipna=True)

    # ------------------------------------------------------------
    # Composants les plus chers (pour n_pcb)
    # ------------------------------------------------------------
    cols_expensive = [c for c in [
        "Manufacturer Ref", "Value", "Footprint", "type",
        "Qty_total", "Unit_price", "Line_total_price",
    ] if c in merged.columns]

    top_expensive = (
        merged.dropna(subset=["Line_total_price"])
        .sort_values("Line_total_price", ascending=False)
        [cols_expensive]
        .head(top_n)
    )

    # ------------------------------------------------------------
    # Coût total par famille de composant ("type" : Resistor, Capacitor...)
    # ------------------------------------------------------------
    if "type" in merged.columns:
        cost_by_type = (
            merged.dropna(subset=["Line_total_price"])
            .groupby("type")["Line_total_price"]
            .sum()
            .sort_values(ascending=False)
        )
    else:
        cost_by_type = None

    # ------------------------------------------------------------
    # Composants les plus récurrents (par quantité totale utilisée)
    # ------------------------------------------------------------
    ref_col = "Manufacturer Ref" if "Manufacturer Ref" in merged.columns else None

    if ref_col:
        agg = {"Qty_total": "sum"}
        cols_group = [ref_col]
        extra_cols = [c for c in ["Value", "Footprint", "type"] if c in merged.columns]

        most_recurring = (
            merged.groupby(cols_group + extra_cols, dropna=False)
            .agg(Qty_total=("Qty_total", "sum"), Nb_lignes_BOM=("Qty_total", "size"))
            .reset_index()
            .sort_values("Qty_total", ascending=False)
            .head(top_n)
        )
    else:
        most_recurring = None

    return {
        "detail": merged,
        "missing": missing,
        "total_price": total_price,
        "top_expensive": top_expensive,
        "cost_by_type": cost_by_type,
        "most_recurring": most_recurring,
    }


def print_report(report, n_pcb=1):
    """Affiche un résumé lisible du rapport dans la console."""

    detail = report["detail"]
    missing = report["missing"]
    total_price = report["total_price"]

    print("=" * 60)
    print(f"RAPPORT BOM - pour {n_pcb} PCB")
    print("=" * 60)
    print(f"Nombre de lignes BOM       : {len(detail)}")
    print(f"Lignes OK                  : {(detail['Status'] == 'OK').sum()}")
    print(f"Lignes incomplètes/erreurs : {len(missing)}")
    print(f"Prix total estimé          : {total_price:.2f} EUR")

    print("-" * 60)
    print("Top composants les plus chers :")
    if report.get("top_expensive") is not None and not report["top_expensive"].empty:
        print(report["top_expensive"].to_string(index=False))
    else:
        print("  (pas assez de données de prix)")

    print("-" * 60)
    print("Coût par famille de composant :")
    if report.get("cost_by_type") is not None and not report["cost_by_type"].empty:
        for type_, cout in report["cost_by_type"].items():
            print(f"  {type_!s:25} {cout:8.2f} EUR")
    else:
        print("  (colonne 'type' absente de la librairie)")

    print("-" * 60)
    print("Composants les plus récurrents :")
    if report.get("most_recurring") is not None and not report["most_recurring"].empty:
        print(report["most_recurring"].to_string(index=False))
    else:
        print("  (pas de colonne Manufacturer Ref)")

    print("-" * 60)

    if not missing.empty:
        print("Composants à vérifier :")
        cols = [c for c in [
            "Reference", "Value", "Footprint", "Manufacturer Ref",
            "LCSC_part_number", "Status", "Comments",
        ] if c in missing.columns]
        print(missing[cols].to_string(index=False))
    else:
        print("Tous les composants sont correctement identifiés.")

    print("=" * 60)


def export_report_excel(report, output_path, n_pcb=1):
    """Exporte le rapport (détail, à vérifier, résumé, et analyses) en Excel."""

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        report["detail"].to_excel(writer, sheet_name="Detail", index=False)
        report["missing"].to_excel(writer, sheet_name="A verifier", index=False)

        if report.get("top_expensive") is not None:
            report["top_expensive"].to_excel(writer, sheet_name="Plus chers", index=False)

        if report.get("cost_by_type") is not None:
            report["cost_by_type"].rename("Line_total_price").to_frame().to_excel(
                writer, sheet_name="Cout par famille"
            )

        if report.get("most_recurring") is not None:
            report["most_recurring"].to_excel(writer, sheet_name="Plus recurrents", index=False)

        summary = pd.DataFrame({
            "Metrique": ["Nombre de PCB", "Prix total (EUR)", "Lignes OK", "Lignes a verifier"],
            "Valeur": [
                n_pcb,
                round(report["total_price"], 2),
                (report["detail"]["Status"] == "OK").sum(),
                len(report["missing"]),
            ],
        })
        summary.to_excel(writer, sheet_name="Resume", index=False)


# ------------------------------------------------------------
# Exemple d'utilisation dans ton main :
#
#   Update_Price_Stock(CHEMIN_LIB)
#   lib_avec_prix = pd.read_excel(CHEMIN_LIB)  # recharger AVEC Price/Stock
#
#   report = build_bom_report(bom, lib_avec_prix, n_pcb=10)
#   print_report(report, n_pcb=10)
#   export_report_excel(report, "rapport_bom.xlsx", n_pcb=10)
# ------------------------------------------------------------