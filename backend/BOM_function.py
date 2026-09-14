import pandas as pd


def _add_comment(bom_line, message):
    """Ajoute un message dans Comments, en le concaténant s'il y en a déjà un."""
    existing_comment = bom_line.get("Comments", "")

    if pd.notna(existing_comment) and str(existing_comment).strip():
        bom_line["Comments"] = f"{existing_comment} | {message}"
    else:
        bom_line["Comments"] = message

    return bom_line


def fill_bom_result(bom_line, result):
    """
    Remplit LCSC_part_number uniquement s'il est vide.
    Ne touche jamais à une valeur déjà présente (voir check_bom pour la
    vérification de cohérence dans ce cas).

    Si le composant trouvé n'a pas de reference_LCSC mais a une
    reference_Mouser, on ne peut rien mettre dans LCSC_part_number :
    on signale la référence Mouser disponible dans Comments à la place.
    """

    if result is None or result.empty or len(result) != 1:
        return bom_line

    current_lcsc = bom_line.get("LCSC_part_number")

    if pd.notna(current_lcsc) and str(current_lcsc).strip() != "":
        return bom_line

    reference_lcsc = result.iloc[0].get("reference_LCSC")

    if pd.notna(reference_lcsc) and str(reference_lcsc).strip() != "":
        bom_line["LCSC_part_number"] = reference_lcsc
        return bom_line

    # Pas de reference_LCSC : on regarde s'il existe une reference_Mouser
    reference_mouser = result.iloc[0].get("reference_Mouser")

    if pd.notna(reference_mouser) and str(reference_mouser).strip() != "":
        _add_comment(bom_line, f"NO LCSC REF - MOUSER: {reference_mouser}")

    return bom_line


def check_bom(bom_line, result):
    """
    Vérifie que le LCSC_part_number déjà renseigné dans le BOM correspond
    bien au composant retrouvé dans la librairie (par Manufacturer Ref /
    Footprint / Value).

    Si le BOM a déjà une référence ET qu'elle diffère de ce que la
    recherche retrouve dans la lib, ajoute "REFERENCE MISMATCH" dans
    Comments (ex: le BOM référence "WS2812B-B/W" mais la lib ne contient
    que la variante "WS2812B-V6" -> même s'il y a "un" résultat, ce n'est
    pas la bonne pièce).

    Ne dit rien si :
    - le BOM n'a pas encore de LCSC_part_number (rien à comparer)
    - la recherche n'a rien trouvé, ou trouve plusieurs candidats ambigus
      (dans ce cas c'est fill_bom_result / le statut "NOT DETECTED" ou
      "AMBIGUOUS" du rapport qui s'applique, pas un mismatch)
    """

    current_lcsc = bom_line.get("LCSC_part_number")

    if pd.isna(current_lcsc) or str(current_lcsc).strip() == "":
        return bom_line

    if result is None or result.empty:
        return bom_line

    if len(result) != 1:
        # Résultat ambigu : on ne peut pas affirmer qu'il y a un mismatch
        return bom_line

    reference_lcsc = result.iloc[0].get("reference_LCSC")

    if pd.isna(reference_lcsc) or str(reference_lcsc).strip() == "":
        # La lib n'a pas de LCSC pour ce composant (Mouser only) : le BOM a
        # une référence que la lib ne peut pas confirmer par ce biais.
        reference_mouser = result.iloc[0].get("reference_Mouser")
        if pd.notna(reference_mouser) and str(reference_mouser).strip() != "":
            _add_comment(bom_line, f"LCSC IN BOM NOT IN LIBRARY - MOUSER: {reference_mouser}")
        return bom_line

    if str(current_lcsc).strip() != str(reference_lcsc).strip():
        _add_comment(bom_line, "REFERENCE MISMATCH")

    return bom_line