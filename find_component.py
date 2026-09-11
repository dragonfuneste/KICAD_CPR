import os
import logging
from difflib import SequenceMatcher
import pandas as pd

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# poids par champ : un nom ou une empreinte qui matchent comptent plus
# qu'une simple ressemblance de type/description
POIDS_CHAMPS = {
    "nom du composant": 1.1,
    "empreinte": 1.1,
    "type": 1,
    "specification": 0.8,
}


def _score_texte(val_bom, val_lib):
    """Score de base texte à texte (égalité / sous-chaîne / similarité floue)."""
    if val_bom == val_lib:
        return 1.0
    if val_bom in val_lib or val_lib in val_bom:
        return 0.85
    return SequenceMatcher(None, val_bom, val_lib).ratio()


def _score_champ(val_bom, val_lib, champ):
    """Score pour un champ donné, avec bonus spécifique au nom (préfixe commun)."""
    base = _score_texte(val_bom, val_lib)

    if champ == "nom du composant":
        prefixe = os.path.commonprefix([val_bom, val_lib])
        if len(prefixe) >= 4:
            ratio_prefixe = len(prefixe) / max(len(val_bom), len(val_lib))
            base = max(base, ratio_prefixe)

    return base


def find_by_row(library, bom_row, bom_fields=None, lib_fields=None, threshold=0.35):
    """
    Recherche le(s) composant(s) de la bibliothèque qui correspondent le mieux
    à une ligne entière du BOM, en pondérant certains champs (nom, empreinte)
    plus fort que d'autres (type, description).
    """
    bom_fields = bom_fields or ["Manufacturer Ref", "Value", "Description", "LCSC_part_number"]
    lib_fields = lib_fields or ["nom du composant", "type", "specification", "empreinte"]

    valeurs_bom = []
    for f in bom_fields:
        if f in bom_row.index and pd.notna(bom_row[f]):
            v = str(bom_row[f]).strip().lower()
            if v:
                valeurs_bom.append(v)

    if not valeurs_bom:
        return None

    scores = []
    for _, lib_row in library.iterrows():
        best_score = 0.0
        for lf in lib_fields:
            val_lib = lib_row.get(lf)
            if pd.isna(val_lib):
                continue
            val_lib_norm = str(val_lib).strip().lower()
            poids = POIDS_CHAMPS.get(lf, 1.0)

            for val_bom in valeurs_bom:
                score_brut = _score_champ(val_bom, val_lib_norm, lf)
                score_pondere = min(score_brut * poids, 1.0)
                best_score = max(best_score, score_pondere)
        scores.append(best_score)

    result = library.copy()
    result["__score"] = scores
    result = result[result["__score"] >= threshold].sort_values("__score", ascending=False)

    if result.empty:
        logger.info("Aucun résultat trouvé")
        return None
    if len(result) == 1:
        logger.info("Un seul résultat trouvé")
        return result.iloc[0]

    top_score = result.iloc[0]["__score"]
    second_score = result.iloc[1]["__score"]
    if top_score - second_score >= 0.05 or top_score == 1.0:
        logger.info("Un résultat trouvé avec un score nettement supérieur aux autres")
        return result.iloc[0]

    logger.info("plusieurs résultats trouvés")
    return result




def find_by_keyword(library, keyword, fields=None, threshold=0.35):
    """
    Recherche un mot-clef unique dans plusieurs champs de la bibliothèque
    (contrairement à find_by_row qui utilise toute une ligne de BOM).
    Retourne None / Series / DataFrame, comme find_by_row.
    """
    fields = fields or ["nom du composant", "type", "specification", "empreinte"]
    fields = [f for f in fields if f in library.columns]
    keyword_norm = str(keyword).strip().lower()
    if not keyword_norm:
        return None

    scores = []
    for _, row in library.iterrows():
        best_field_score = 0.0
        for f in fields:
            val = row.get(f)
            if pd.isna(val):
                continue
            val_norm = str(val).strip().lower()
            best_field_score = max(best_field_score, _score_champ(keyword_norm, val_norm, f))
        scores.append(best_field_score)

    result = library.copy()
    result["__score"] = scores
    result = result[result["__score"] >= threshold].sort_values("__score", ascending=False)

    if result.empty:
        return None
    if len(result) == 1:
        return result.iloc[0]

    top_score = result.iloc[0]["__score"]
    second_score = result.iloc[1]["__score"]
    if top_score - second_score >= 0.05 or top_score == 1.0:
        return result.iloc[0]

    return result