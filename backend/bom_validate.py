"""
Statut unifié d'une ligne de BOM (pastilles de couleur) + validation du composant proposé par la lib.

  🟢 ok        : composant trouvé, réf cohérente (ou pas de réf dans le BOM : rien à contredire)
  🟣 ref       : trouvé (même valeur / footprint) mais la réf fabricant du BOM diffère de celle de la lib
                 -> peut être validé : le BOM adopte la réf de la lib
  🔵 mouser    : le composant de la lib n'a qu'une réf Mouser (pas de LCSC)
  🟠 mismatch  : la réf LCSC du BOM diffère de celle de la lib (une seule proposition)
                 -> peut être validé : le BOM adopte la réf LCSC de la lib
  🔴 bad       : ndddon détecté / plusieurs candiddddats / absent de la lib / erreur
  🟢 validated : ligne validée par l'utilisateur
"""

import re

import pandas as pd

VALIDATABLE = ("ref", "mismatch")


def _filled(x) -> bool:
    return pd.notna(x) and str(x).strip() != ""


def _compact(x) -> str:
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return ""
    return re.sub(r"[^a-z0-9]", "", str(x).lower())


def refs_match(a, b) -> bool:
    """Deux réfs fabricant désignent-elles la même pièce ? (ponctuation ignorée, l'une peut contenir l'autre)"""
    na, nb = _compact(a), _compact(b)
    if not na or not nb:
        return True
    if na == nb:
        return True
    return min(len(na), len(nb)) >= 5 and (na in nb or nb in na)


def row_status(row, validated: bool = False):
    """Retourne (kind, texte avec pastille). `row` = ligne de BOM jointe à `matches`."""
    if validated:
        return "validated", "🟢 Validé (réf de la lib)"

    n = row.get("Nb_matches", 0)
    n = 0 if pd.isna(n) else int(n)
    has_lcsc = _filled(row.get("LCSC_part_number"))
    comments = str(row.get("Comments")) if pd.notna(row.get("Comments")) else ""

    if "ERROR IN PART NUMBER" in comments:
        return "bad", "🔴 Erreur dans la référence"
    if "REFERENCE MISMATCH" in comments:
        return "mismatch", "🟠 Réf LCSC du BOM ≠ lib (1 proposition)"
    if n == 0:
        return ("bad", "🔴 Absent de la lib") if has_lcsc else ("bad", "🔴 Non détecté")
    if n > 1 and not has_lcsc:
        return "bad", f"🔴 {n} candidats : à trancher"
    if not has_lcsc:
        if "MOUSER:" in comments:
            return "mouser", "🔵 OK : réf Mouser uniquement"
        return "bad", "🔴 Composant de la lib sans réf LCSC ni Mouser"
    if "LCSC IN BOM NOT IN LIBRARY" in comments:
        return "mouser", "🔵 OK : réf Mouser (LCSC du BOM absent de la lib)"

    if n == 1 and _filled(row.get("Manufacturer Ref")) and _filled(row.get("Lib_Manufacturer_Ref")) \
            and not refs_match(row["Manufacturer Ref"], row["Lib_Manufacturer_Ref"]):
        return "ref", "🟣 OK mais réf différente"
    return "ok", "🟢 OK"


def can_validate(kind: str, row) -> bool:
    return kind in VALIDATABLE and pd.notna(row.get("Lib_idx"))


def clean_comment(comments, remove=("REFERENCE MISMATCH",)):
    if not _filled(comments):
        return comments
    parts = [p.strip() for p in str(comments).split("|")]
    parts = [p for p in parts if p and p not in remove]
    return " | ".join(parts) if parts else pd.NA


_SAVED_COLUMNS = ("Manufacturer Ref", "Manufacturer", "LCSC_part_number", "Comments")


def validate_row(bom: pd.DataFrame, lib: pd.DataFrame, index, lib_idx) -> dict:
    """
    Le BOM adopte le composant de la lib (réf fabricant, fabricant, réf LCSC).
    Retourne les anciennes valeurs pour pouvoir annuler.
    """
    saved = {c: bom.loc[index, c] for c in _SAVED_COLUMNS if c in bom.columns}
    lib_row = lib.loc[int(lib_idx)]

    if _filled(lib_row.get("Manufacturer Ref")):
        bom.loc[index, "Manufacturer Ref"] = str(lib_row["Manufacturer Ref"]).strip()
    if "Manufacturer" in bom.columns and _filled(lib_row.get("Manufacturer Part")):
        bom.loc[index, "Manufacturer"] = str(lib_row["Manufacturer Part"]).strip()
    if _filled(lib_row.get("reference_LCSC")):
        bom.loc[index, "LCSC_part_number"] = str(lib_row["reference_LCSC"]).strip()
    if "Comments" in bom.columns:
        bom.loc[index, "Comments"] = clean_comment(bom.loc[index, "Comments"])
    return saved


def revert_row(bom: pd.DataFrame, index, saved: dict):
    for col, val in saved.items():
        bom.loc[index, col] = val
