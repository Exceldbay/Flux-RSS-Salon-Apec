from pathlib import Path
from urllib.parse import quote_plus
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import threading
import time
import re

import feedparser
import pandas as pd
import yaml
import requests
from requests.adapters import HTTPAdapter


# ============================================================
# CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parent.parent

CONFIG = ROOT / "config"
DATA = ROOT / "data"
OUTPUT = ROOT / "output"

MAX_WORKERS = 16
TIMEOUT = 10
MAX_RETRIES = 2

thread_local = threading.local()


# ============================================================
# CHARGEMENT CONFIGURATION
# ============================================================

def load_yaml(filename):
    path = CONFIG / filename

    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


# ============================================================
# SESSION HTTP
# ============================================================

def get_session():

    if not hasattr(thread_local, "session"):

        session = requests.Session()

        adapter = HTTPAdapter(
            pool_connections=MAX_WORKERS,
            pool_maxsize=MAX_WORKERS,
            max_retries=0
        )

        session.mount("http://", adapter)
        session.mount("https://", adapter)

        session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "Chrome/131.0 Safari/537.36"
            )
        })

        thread_local.session = session

    return thread_local.session


# ============================================================
# MOTS-CLÉS
# ============================================================

REQUETES_EMPLOI = [
    "salon emploi",
    "forum emploi",
    "job dating",
    "forum recrutement",
    "salon recrutement",
    "salon formation",
    "forum formation",
    "emploi cadre"
]

REQUETES_SUPERIEUR = [
    "salon étudiant",
    "forum étudiant",
    "enseignement supérieur",
    "études supérieures",
    "portes ouvertes supérieur",
    "journée portes ouvertes université",
    "alumni",
    "rencontre entreprises étudiants"
]


# ============================================================
# SCORING — VERSION APEC RENFORCÉE
# ============================================================

TERMES_EMPLOI_FORTS = {
    "salon emploi": 6,
    "forum emploi": 6,
    "forum de l'emploi": 6,
    "job dating": 6,
    "forum recrutement": 6,
    "salon recrutement": 6,
    "recrutement": 4,
    "emploi cadre": 8,
    "emploi cadres": 8,
    "recrutement cadre": 8,
    "recrutement cadres": 8,
    "jeunes diplômés": 5,
    "jeune diplômé": 5,
    "cadres": 5,
}

TERMES_FORMATION_FORTS = {
    "salon formation": 5,
    "forum formation": 5,
    "formation professionnelle": 5,
    "forum de la formation": 5,
    "formation continue": 4,
    "reconversion professionnelle": 4,
    "évolution professionnelle": 4,
}

TERMES_SUPERIEUR_FORTS = {
    "salon étudiant": 6,
    "salon étudiants": 6,
    "forum étudiant": 6,
    "forum étudiants": 6,
    "enseignement supérieur": 6,
    "études supérieures": 6,
    "école d'ingénieurs": 6,
    "école ingénieurs": 6,
    "école de commerce": 5,
    "université": 4,
    "faculté": 4,
    "alumni": 4,
    "rencontre alumni": 5,
    "rencontre entreprises étudiants": 6,
    "portes ouvertes supérieur": 6,
    "journée portes ouvertes université": 6,
    "portes ouvertes école": 5,
    "journée portes ouvertes école": 5,
}

TERMES_APEC = {
    "cadre": 3,
    "cadres": 5,
    "jeune diplômé": 4,
    "jeunes diplômés": 5,
    "bac+5": 4,
    "bac +5": 4,
    "ingénieur": 3,
    "ingénieurs": 4,
    "management": 2,
    "ressources humaines": 2,
    "rh": 1,
    "carrière": 3,
    "carrières": 3,
}

TERMES_COMPLEMENTAIRES = {
    "emploi": 2,
    "recrutement": 2,
    "formation": 2,
    "étudiant": 2,
    "étudiants": 2,
    "supérieur": 2,
    "université": 2,
    "école": 1,
    "entreprise": 1,
    "entreprises": 1,
    "recruteur": 2,
    "recruteurs": 2,
    "alternance": 2,
    "apprentissage": 1,
    "forum": 1,
    "salon": 1,
    "rencontre": 1,
    "orientation": 2,
}

TERMES_BRUIT = {
    "concert": -6,
    "spectacle": -6,
    "théâtre": -6,
    "football": -5,
    "rugby": -5,
    "basket": -5,
    "marché": -5,
    "vide-grenier": -6,
    "brocante": -6,
    "braderie": -5,
    "patrimoine": -4,
    "agriculture": -4,
    "météo": -5,
    "faits divers": -6,
    "incendie": -6,
    "accident": -5,
    "politique": -4,
    "mariage": -5,
    "gastronomie": -4,
    "tourisme": -4,
}

TERMES_EXCLUSION_FORTS = [
    "cabinet de recrutement privé",
    "cabinet recrutement privé",
    "agence de recrutement",
    "cabinet de recrutement",
]

TERMES_CCI = [
    "chambre de commerce et d'industrie organise",
    "chambre de commerce et d'industrie organise le",
    "chambre de commerce organise",
    "cci organise",
    "cci organise le",
    "organisé par la cci",
    "organisee par la cci",
    "organisé par la chambre de commerce",
    "organisee par la chambre de commerce",
]

VILLES_DEPARTEMENTS = {
    "86": [
        "poitiers", "châtellerault", "chatellerault",
        "buxerolles", "jaunay-marigny", "mignaloux-beauvoir"
    ],
    "79": [
        "niort", "bressuire", "parthenay", "thouars"
    ],
    "87": [
        "limoges", "panazol", "saint-junien", "bellac"
    ],
    "17": [
        "la rochelle", "rochefort", "saintes", "royan",
        "surgères", "surgere", "jonzac"
    ],
    "16": [
        "angoulême", "angouleme", "cognac", "jarnac", "confolens"
    ],
    "23": [
        "guéret", "gueret", "aubusson", "la souterraine"
    ],
    "19": [
        "brive", "tulle", "ussel", "égletons", "egletons"
    ],
}


# ============================================================
# NORMALISATION
# ============================================================

def normaliser(texte):

    texte = str(texte).lower()

    texte = re.sub(
        r"\s+",
        " ",
        texte
    )

    return texte.strip()


# ============================================================
# CONSTRUCTION DES REQUÊTES
# ============================================================

def construire_requetes(departements):

    requetes = []

    for code, nom in departements["departements"].items():

        for annee in departements["annees"]:

            for mot in REQUETES_EMPLOI:

                requetes.append({
                    "departement": code,
                    "nom_departement": nom,
                    "annee": annee,
                    "categorie_numero": 2,
                    "categorie": (
                        "Emploi / recrutement / formation"
                    ),
                    "mot_cle": mot
                })

            for mot in REQUETES_SUPERIEUR:

                requetes.append({
                    "departement": code,
                    "nom_departement": nom,
                    "annee": annee,
                    "categorie_numero": 1,
                    "categorie": (
                        "Salon étudiant / enseignement supérieur"
                    ),
                    "mot_cle": mot
                })

    return requetes


# ============================================================
# RECHERCHE GOOGLE NEWS RSS
# ============================================================

def rechercher_google_news(requete):

    # Requête volontairement simple.
    # On évite les grosses expressions avec OR et parenthèses
    # qui provoquaient la collecte à zéro.

    texte = (
        f'"{requete["mot_cle"]}" '
        f'"{requete["nom_departement"]}" '
        f'{requete["annee"]}'
    )

    url = (
        "https://news.google.com/rss/search?"
        f"q={quote_plus(texte)}"
        "&hl=fr"
        "&gl=FR"
        "&ceid=FR:fr"
    )

    session = get_session()

    derniere_erreur = ""

    for tentative in range(MAX_RETRIES):

        try:

            response = session.get(
                url,
                timeout=TIMEOUT
            )

            response.raise_for_status()

            flux = feedparser.parse(
                response.content
            )

            resultats = []

            for article in flux.entries:

                titre = article.get(
                    "title",
                    ""
                ).strip()

                lien = article.get(
                    "link",
                    ""
                ).strip()

                date_publication = article.get(
                    "published",
                    ""
                ).strip()

                # Description / résumé RSS.
                # Elle permet d'améliorer fortement le scoring.
                description = article.get(
                    "summary",
                    ""
                )

                description = re.sub(
                    r"<[^>]+>",
                    " ",
                    description
                )

                description = normaliser(
                    description
                )

                source = ""

                if hasattr(article, "source"):

                    source = article.source.get(
                        "title",
                        ""
                    )

                if not titre or not lien:
                    continue

                resultats.append({
                    "titre": titre,
                    "url": lien,
                    "source": source,
                    "description": description,
                    "date_publication": date_publication,
                    "departement": (
                        requete["departement"]
                    ),
                    "nom_departement": (
                        requete["nom_departement"]
                    ),
                    "annee_recherchee": (
                        requete["annee"]
                    ),
                    "categorie_numero": (
                        requete["categorie_numero"]
                    ),
                    "categorie": (
                        requete["categorie"]
                    ),
                    "mot_cle": (
                        requete["mot_cle"]
                    )
                })

            return {
                "ok": True,
                "resultats": resultats,
                "erreur": ""
            }

        except Exception as erreur:

            derniere_erreur = str(erreur)

            if tentative < MAX_RETRIES - 1:
                time.sleep(0.5)

    return {
        "ok": False,
        "resultats": [],
        "erreur": derniere_erreur
    }


# ============================================================
# COLLECTE PARALLÈLE
# ============================================================

def collecter(requetes):

    resultats = []

    erreurs = 0
    requetes_avec_resultats = 0

    total = len(requetes)
    termine = 0

    debut = time.time()

    print()
    print("========================================")
    print("COLLECTE RSS")
    print("========================================")
    print(
        f"Requêtes : {total}"
    )
    print(
        f"Workers : {MAX_WORKERS}"
    )
    print()

    with ThreadPoolExecutor(
        max_workers=MAX_WORKERS
    ) as executor:

        futures = {
            executor.submit(
                rechercher_google_news,
                requete
            ): requete
            for requete in requetes
        }

        for future in as_completed(futures):

            termine += 1

            try:

                resultat = future.result()

                if resultat["ok"]:

                    if resultat["resultats"]:
                        requetes_avec_resultats += 1

                    resultats.extend(
                        resultat["resultats"]
                    )

                else:

                    erreurs += 1

            except Exception:

                erreurs += 1

            if (
                termine % 20 == 0
                or termine == total
            ):

                print(
                    f"Progression : "
                    f"{termine}/{total} "
                    f"| résultats : "
                    f"{len(resultats)} "
                    f"| erreurs : "
                    f"{erreurs}"
                )

    duree = time.time() - debut

    print()
    print(
        f"Collecte terminée en "
        f"{duree:.1f} secondes"
    )

    print(
        f"Requêtes ayant retourné "
        f"au moins un résultat : "
        f"{requetes_avec_resultats}/{total}"
    )

    return resultats, erreurs


# ============================================================
# DÉDOUBLONNAGE
# ============================================================

def dedoublonner(resultats):

    urls = set()
    titres = set()

    uniques = []

    for resultat in resultats:

        url = normaliser(
            resultat.get("url", "")
        )

        titre = normaliser(
            resultat.get("titre", "")
        )

        titre_cle = re.sub(
            r"[^a-z0-9àâäéèêëîïôöùûüç ]",
            "",
            titre
        )

        if not url:
            continue

        if url in urls:
            continue

        if titre_cle in titres:
            continue

        urls.add(url)
        titres.add(titre_cle)

        uniques.append(resultat)

    return uniques


# ============================================================
# SCORE
# ============================================================

def contient_terme(texte, terme):
    """
    Recherche après normalisation.
    """

    texte = normaliser(texte)
    terme = normaliser(terme)

    return terme in texte


def score_termes(texte, dictionnaire):
    """
    Additionne les points des termes détectés.
    Un terme ne rapporte des points qu'une seule fois.
    """

    score = 0
    details = []

    for terme, points in dictionnaire.items():

        if contient_terme(
            texte,
            terme
        ):

            score += points

            details.append(
                f"{terme} ({points:+d})"
            )

    return score, details


def calculer_score(resultat):

    titre = normaliser(
        resultat.get("titre", "")
    )

    source = normaliser(
        resultat.get("source", "")
    )

    description = normaliser(
        resultat.get("description", "")
    )

    texte_global = " ".join([
        titre,
        source,
        description
    ])

    score_emploi = 0
    score_formation = 0
    score_superieur = 0
    score_apec = 0
    score_geographie = 0
    score_date = 0
    score_bruit = 0

    details = []

    # --------------------------------------------------------
    # 1. EMPLOI / RECRUTEMENT
    # --------------------------------------------------------

    score, elements = score_termes(
        texte_global,
        TERMES_EMPLOI_FORTS
    )

    score_emploi += score
    details.extend(elements)

    # --------------------------------------------------------
    # 2. FORMATION
    # --------------------------------------------------------

    score, elements = score_termes(
        texte_global,
        TERMES_FORMATION_FORTS
    )

    score_formation += score
    details.extend(elements)

    # --------------------------------------------------------
    # 3. ENSEIGNEMENT SUPÉRIEUR
    # --------------------------------------------------------

    score, elements = score_termes(
        texte_global,
        TERMES_SUPERIEUR_FORTS
    )

    score_superieur += score
    details.extend(elements)

    # --------------------------------------------------------
    # 4. DIMENSION APEC / CADRES
    # --------------------------------------------------------

    score, elements = score_termes(
        texte_global,
        TERMES_APEC
    )

    score_apec += score
    details.extend(elements)

    # --------------------------------------------------------
    # 5. TERMES COMPLÉMENTAIRES
    # --------------------------------------------------------

    score, elements = score_termes(
        texte_global,
        TERMES_COMPLEMENTAIRES
    )

    score_complementaire = score
    details.extend(elements)

    # --------------------------------------------------------
    # 6. BONUS TITRE
    # --------------------------------------------------------

    bonus_titre = 0

    expressions_titre = (
        list(TERMES_EMPLOI_FORTS.keys())
        + list(TERMES_FORMATION_FORTS.keys())
        + list(TERMES_SUPERIEUR_FORTS.keys())
        + list(TERMES_APEC.keys())
    )

    for terme in expressions_titre:

        if contient_terme(
            titre,
            terme
        ):

            bonus_titre += 2

    bonus_titre = min(
        bonus_titre,
        8
    )

    if bonus_titre:
        details.append(
            f"Bonus titre (+{bonus_titre})"
        )

    # --------------------------------------------------------
    # 7. GÉOGRAPHIE
    # --------------------------------------------------------

    departement = str(
        resultat.get("departement", "")
    )

    nom_departement = normaliser(
        resultat.get("nom_departement", "")
    )

    if (
        departement
        and departement in VILLES_DEPARTEMENTS
    ):

        if (
            departement in texte_global
            or nom_departement in texte_global
        ):

            score_geographie += 3

            details.append(
                "Département ciblé (+3)"
            )

        for ville in VILLES_DEPARTEMENTS[
            departement
        ]:

            if ville in texte_global:

                score_geographie += 2

                details.append(
                    f"Ville ciblée : {ville} (+2)"
                )

                break

    # --------------------------------------------------------
    # 8. ANNÉE
    # --------------------------------------------------------

    annee = str(
        resultat.get(
            "annee_recherchee",
            ""
        )
    )

    if annee == "2027":

        score_date += 3

        details.append(
            "Édition 2027 (+3)"
        )

    elif annee == "2026":

        score_date += 2

        details.append(
            "Édition 2026 (+2)"
        )

    elif annee == "2025":

        score_date += 1

        details.append(
            "Édition 2025 (+1)"
        )

    # --------------------------------------------------------
    # 9. DATE D'ÉVÉNEMENT DÉTECTÉE
    # --------------------------------------------------------

    motif_date = (
        r"\b\d{1,2}\s+"
        r"(janvier|février|fevrier|mars|avril|mai|juin|"
        r"juillet|août|aout|septembre|octobre|novembre|"
        r"décembre|decembre)\b"
    )

    if re.search(
        motif_date,
        texte_global
    ):

        score_date += 2

        details.append(
            "Date d'événement détectée (+2)"
        )

    # --------------------------------------------------------
    # 10. BRUIT
    # --------------------------------------------------------

    for terme, points in TERMES_BRUIT.items():

        if contient_terme(
            texte_global,
            terme
        ):

            score_bruit += points

            details.append(
                f"{terme} ({points:+d})"
            )

    # --------------------------------------------------------
    # 11. SCORE FINAL
    # --------------------------------------------------------

    score_brut = (
        score_emploi
        + score_formation
        + score_superieur
        + score_apec
        + score_complementaire
        + bonus_titre
        + score_geographie
        + score_date
        + score_bruit
    )

    score = max(
        -20,
        min(score_brut, 40)
    )

    return {
        "score": score,
        "score_emploi": score_emploi,
        "score_formation": score_formation,
        "score_superieur": score_superieur,
        "score_apec": score_apec,
        "score_geographie": score_geographie,
        "score_date": score_date,
        "score_bruit": score_bruit,
        "score_complementaire": score_complementaire,
        "bonus_titre": bonus_titre,
        "details_score": " | ".join(details),
    }


# ============================================================
# TRAITEMENT
# ============================================================

def traiter(resultats):

    resultats = dedoublonner(
        resultats
    )

    pertinents = []
    a_verifier = []
    non_pertinents = []

    for resultat in resultats:

        titre = normaliser(
            resultat.get("titre", "")
        )

        source = normaliser(
            resultat.get("source", "")
        )

        description = normaliser(
            resultat.get("description", "")
        )

        texte = " ".join([
            titre,
            source,
            description
        ])

        # ----------------------------------------------------
        # EXCLUSIONS FORTES
        # ----------------------------------------------------

        exclusion = False
        motif_exclusion = ""

        for terme in TERMES_EXCLUSION_FORTS:

            if terme in texte:

                exclusion = True
                motif_exclusion = terme
                break

        # ----------------------------------------------------
        # CCI
        # ----------------------------------------------------

        exclusion_cci = False

        for terme in TERMES_CCI:

            if terme in texte:

                exclusion_cci = True
                motif_exclusion = terme
                break

        if exclusion:

            resultat["score"] = -20
            resultat["statut"] = (
                "Non pertinent"
            )
            resultat["motif"] = (
                motif_exclusion
            )
            resultat["raison"] = (
                "Exclusion : recrutement privé"
            )

            non_pertinents.append(
                resultat
            )

            continue

        if exclusion_cci:

            resultat["score"] = -20
            resultat["statut"] = (
                "Non pertinent"
            )
            resultat["motif"] = (
                motif_exclusion
            )
            resultat["raison"] = (
                "Événement organisé par une CCI"
            )

            non_pertinents.append(
                resultat
            )

            continue

        # ----------------------------------------------------
        # CALCUL DU SCORE
        # ----------------------------------------------------

        scoring = calculer_score(
            resultat
        )

        for cle, valeur in scoring.items():
            resultat[cle] = valeur

        score = scoring["score"]

        resultat["motif"] = scoring[
            "details_score"
        ]

        # ----------------------------------------------------
        # SCORE MÉTIER
        # ----------------------------------------------------

        score_metier = (
            scoring["score_emploi"]
            + scoring["score_formation"]
            + scoring["score_superieur"]
            + scoring["score_apec"]
        )

        # ----------------------------------------------------
        # PERTINENT
        # ----------------------------------------------------

        if (
            score >= 12
            and score_metier >= 5
        ):

            resultat["statut"] = (
                "Pertinent"
            )

            resultat["raison"] = (
                "Signal métier Apec suffisamment fort"
            )

            pertinents.append(
                resultat
            )

        # ----------------------------------------------------
        # À VÉRIFIER
        # ----------------------------------------------------

        elif score >= 6:

            resultat["statut"] = (
                "À vérifier"
            )

            resultat["raison"] = (
                "Potentiel événement Apec à vérifier"
            )

            a_verifier.append(
                resultat
            )

        # ----------------------------------------------------
        # NON PERTINENT
        # ----------------------------------------------------

        else:

            resultat["statut"] = (
                "Non pertinent"
            )

            resultat["raison"] = (
                "Score insuffisant"
            )

            non_pertinents.append(
                resultat
            )

    # --------------------------------------------------------
    # TRI
    # --------------------------------------------------------

    pertinents.sort(
        key=lambda x: x.get(
            "score",
            0
        ),
        reverse=True
    )

    a_verifier.sort(
        key=lambda x: x.get(
            "score",
            0
        ),
        reverse=True
    )

    non_pertinents.sort(
        key=lambda x: x.get(
            "score",
            0
        ),
        reverse=True
    )

    return (
        pertinents,
        a_verifier,
        non_pertinents
    )


# ============================================================
# SAUVEGARDE SÉCURISÉE
# ============================================================

def sauvegarder(
    pertinents,
    a_verifier,
    non_pertinents,
    resultats_bruts
):

    date_collecte = datetime.now().isoformat()

    tous = (
        pertinents
        + a_verifier
        + non_pertinents
    )

    for resultat in tous:

        resultat["date_collecte"] = (
            date_collecte
        )

    colonnes = [
        "titre",
        "url",
        "source",
        "description",
        "date_publication",
        "departement",
        "nom_departement",
        "annee_recherchee",
        "categorie_numero",
        "categorie",
        "mot_cle",

        "score",

        "score_emploi",
        "score_formation",
        "score_superieur",
        "score_apec",
        "score_geographie",
        "score_date",
        "score_bruit",
        "score_complementaire",
        "bonus_titre",

        "statut",
        "raison",
        "motif",
        "details_score",

        "date_collecte"
    ]

    def dataframe(resultats):

        df = pd.DataFrame(
            resultats
        )

        for colonne in colonnes:

            if colonne not in df.columns:
                df[colonne] = ""

        return df[colonnes]

    fichier_principal = (
        DATA / "resultats_rss.csv"
    )

    fichier_verification = (
        DATA / "a_verifier.csv"
    )

    fichier_non_pertinent = (
        DATA / "non_pertinent.csv"
    )

    # --------------------------------------------------------
    # PROTECTION CONTRE UNE COLLECTE VIDE
    # --------------------------------------------------------

    if len(resultats_bruts) == 0:

        print()
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print("ATTENTION")
        print("COLLECTE RSS = 0 RESULTAT")
        print("LES ANCIENS RESULTATS SONT CONSERVÉS")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print()

        return False

    # --------------------------------------------------------
    # SAUVEGARDE
    # --------------------------------------------------------

    dataframe(
        pertinents
    ).to_csv(
        fichier_principal,
        index=False,
        encoding="utf-8-sig"
    )

    dataframe(
        a_verifier
    ).to_csv(
        fichier_verification,
        index=False,
        encoding="utf-8-sig"
    )

    dataframe(
        non_pertinents
    ).to_csv(
        fichier_non_pertinent,
        index=False,
        encoding="utf-8-sig"
    )

    return True


# ============================================================
# MAIN
# ============================================================

def main():

    debut_total = time.time()

    DATA.mkdir(
        exist_ok=True
    )

    OUTPUT.mkdir(
        exist_ok=True
    )

    print()
    print("========================================")
    print("VEILLE APEC")
    print("COLLECTEUR RSS + SCORING APEC RENFORCÉ")
    print("========================================")

    departements = load_yaml(
        "departements.yml"
    )

    requetes = construire_requetes(
        departements
    )

    print()
    print(
        f"Requêtes générées : "
        f"{len(requetes)}"
    )

    # --------------------------------------------------------
    # COLLECTE
    # --------------------------------------------------------

    resultats_bruts, erreurs = collecter(
        requetes
    )

    print()
    print("========================================")
    print("TRAITEMENT")
    print("========================================")

    print(
        f"Résultats bruts : "
        f"{len(resultats_bruts)}"
    )

    # --------------------------------------------------------
    # SI ZÉRO RÉSULTAT
    # --------------------------------------------------------

    if len(resultats_bruts) == 0:

        print()
        print("⚠️ AUCUN RÉSULTAT RSS.")

        print(
            "Les fichiers existants "
            "ne seront pas écrasés."
        )

        print()
        print(
            f"Erreurs HTTP/RSS : "
            f"{erreurs}"
        )

        print()
        print("========================================")
        print("COLLECTE NON VALIDÉE")
        print("========================================")

        return

    # --------------------------------------------------------
    # TRAITEMENT
    # --------------------------------------------------------

    (
        pertinents,
        a_verifier,
        non_pertinents
    ) = traiter(
        resultats_bruts
    )

    total_apres_dedoublonnage = (
        len(pertinents)
        + len(a_verifier)
        + len(non_pertinents)
    )

    print()
    print(
        f"Après dédoublonnage : "
        f"{total_apres_dedoublonnage}"
    )

    print(
        f"Pertinents : "
        f"{len(pertinents)}"
    )

    print(
        f"À vérifier : "
        f"{len(a_verifier)}"
    )

    print(
        f"Non pertinents : "
        f"{len(non_pertinents)}"
    )

    # --------------------------------------------------------
    # SAUVEGARDE
    # --------------------------------------------------------

    sauvegarde_ok = sauvegarder(
        pertinents,
        a_verifier,
        non_pertinents,
        resultats_bruts
    )

    duree_total = time.time() - debut_total

    print()
    print("========================================")
    print("STATISTIQUES")
    print("========================================")

    print(
        f"Erreurs : "
        f"{erreurs}"
    )

    print(
        f"Temps total : "
        f"{duree_total:.1f} secondes"
    )

    print(
        f"Sauvegarde validée : "
        f"{'OUI' if sauvegarde_ok else 'NON'}"
    )

    print()
    print("========================================")

    if sauvegarde_ok:

        print(
            "VEILLE TERMINÉE AVEC SUCCÈS"
        )

    else:

        print(
            "VEILLE NON VALIDÉE"
        )

    print("========================================")


if __name__ == "__main__":
    main()
