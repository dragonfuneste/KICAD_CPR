from rapidfuzz import process, fuzz
import pandas as pd

def find_by_value(column, library, value, top_n=5, score_gap=10):
    """
    Recherche une valeur dans une colonne.

    1. Recherche exacte.
    2. Recherche par contenu.
    3. Recherche floue (sans colonne de probabilité).

    Retourne toujours un DataFrame.
    """

    # ---------------------------------------------------------
    # 1. Recherche exacte
    # ---------------------------------------------------------
    exact = library[library[column] == value]

    if not exact.empty:
        return exact

    # ---------------------------------------------------------
    # 2. Recherche de la valeur contenue dans une chaîne
    # ---------------------------------------------------------
    value_str = str(value)

    contains = library[
        library[column]
        .astype(str)
        .str.contains(
            value_str,
            case=False,
            na=False,
            regex=False
        )
    ]

    if not contains.empty:
        return contains

    # ---------------------------------------------------------
    # 3. Recherche floue
    # ---------------------------------------------------------
    choices = library[column].dropna().unique()

    if len(choices) == 0:
        return None

    matches = process.extract(
        value_str,
        choices,
        scorer=fuzz.ratio,
        limit=top_n
    )

    # On garde uniquement les valeurs dans la limite de l'écart de score
    selected_values = []

    previous_score = None

    for match, score, _ in matches:

        # Premier résultat
        if previous_score is None:
            selected_values.append(match)
            previous_score = score
            continue

        # Écart avec le résultat précédent
        gap = previous_score - score

        if gap > score_gap:
            break

        selected_values.append(match)
        previous_score = score

    # ---------------------------------------------------------
    # Créer le DataFrame avec les lignes correspondantes
    # ---------------------------------------------------------
    result = library[library[column].isin(selected_values)].copy()

    return result


def search_in_lib(library, bom_line):
    """
    Recherche un composant en combinant Manufacturer Ref, Footprint, Value.

    Chaque critère est cherché dans la librairie complète. On accumule les
    critères par intersection SEULEMENT si cela ne vide pas le résultat déjà
    obtenu : un critère peu fiable (ex: Value générique comme "STM32G070CBTx"
    ou "WS2812B") ne doit jamais effacer un match exact déjà trouvé par un
    critère plus spécifique comme Manufacturer Ref.

    On s'arrête dès qu'un seul composant reste.
    """

    criteres = [
        ("Manufacturer Ref", "Manufacturer Ref"),
        ("Footprint", "Footprint"),
        ("Value", "Value"),
    ]

    result = None

    for bom_column, lib_column in criteres:

        if bom_column not in bom_line or lib_column not in library.columns:
            continue

        value = bom_line[bom_column]

        if pd.isna(value) or str(value).strip() == "":
            continue

        candidate = find_by_value(lib_column, library, value)

        if candidate is None or candidate.empty:
            continue

        if result is None:
            result = candidate
        else:
            intersection = result[result.index.isin(candidate.index)]
            if not intersection.empty:
                result = intersection
            # sinon : critère contradictoire avec ce qu'on a déjà trouvé,
            # on l'ignore plutôt que d'effacer un résultat déjà fiable.

        if len(result) == 1:
            break

    return result