from pathlib import Path
from urllib.parse import quote_plus
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import threading
import re
import time

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
TIMEOUT = 8
MAX_RETRIES = 2

thread_local = threading.local()


# ============================================================
# CHARGEMENT YAML
# ============================================================

def load_yaml(filename):

    with open(
        CONFIG / filename,
        "r",
        encoding="utf-8"
    ) as f:
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
            "User-Agent":
                "Mozilla/5.0 "
                "(compatible; Veille-APEC/1.0)"
        })

        thread_local.session = session

    return thread_local.session


# ============================================================
# GROUPES DE RECHERCHE
# ============================================================

GROUPES_EMPLOI = [
    [
        "salon emploi",
        "forum emploi",
        "forum de l'emploi",
        "job dating"
    ],
    [
        "forum recrutement",
        "salon recrutement",
        "recrutement",
        "rencontre recruteurs"
    ],
    [
        "formation",
        "formation professionnelle",
        "salon formation",
        "forum formation"
    ],
    [
        "emploi cadre",
        "emploi cadres",
        "rencontre entreprises",
        "conférence carrière"
    ]
]


GROUPES_SUPERIEUR = [
    [
        "salon étudiant",
        "salon étudiants",
        "forum étudiant",
        "forum étudiants"
    ],
    [
        "enseignement supérieur",
        "salon enseignement supérieur",
        "salon supérieur",
        "études supérieures"
    ],
    [
        "portes ouvertes",
        "journée portes ouvertes",
        "open day",
        "journée d'information"
    ],
    [
        "alumni",
        "rencontre alumni",
        "anciens élèves",
        "rencontre entreprises"
    ]
]


# ============================================================
# TERMES DE PERTINENCE
# ============================================================

TERMES_FORTS = [

    "salon emploi",
    "forum emploi",
    "job dating",
    "forum recrutement",
    "salon recrutement",
    "salon formation",
    "forum formation",
    "emploi cadre",
    "emploi cadres",
    "rencontre recruteurs",
    "rencontre entreprises",
    "salon étudiant",
    "forum étudiant",
    "enseignement supérieur",
    "études supérieures",
    "portes ouvertes",
    "journée portes ouvertes",
    "alumni",
    "école d'ingénieurs",
    "école de commerce",
    "université",
    "faculté"
]


TERMES_POSITIFS = [

    "emploi",
    "recrutement",
    "formation",
    "carrière",
    "cadre",
    "cadres",
    "étudiant",
    "étudiants",
    "supérieur",
    "université",
    "école",
    "ingénieur",
    "ingénieurs",
    "alumni",
    "entreprise",
    "entreprises",
    "recruteur",
    "recruteurs",
    "alternance",
    "apprentissage"
]


TERMES_BRUIT = [

    "concert",
    "spectacle",
    "théâtre",
    "festival",
    "football",
    "rugby",
    "basket",
    "marché",
    "vide-grenier",
    "brocante",
    "braderie",
    "patrimoine",
    "exposition",
    "agriculture",
    "météo",
    "faits divers",
    "incendie",
    "accident",
    "politique",
    "municipal",
    "loisirs"
]


TERMES_EXCLUSION_FORTE = [

    "chambre de commerce et d'industrie",
    "chambre de commerce",
    "cci",
    "cabinet de recrutement privé",
    "agence de recrutement"
]


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

    groupes = [
        (
            1,
            "Salon étudiant / enseignement supérieur",
            GROUPES_SUPERIEUR
        ),
        (
            2,
            "Emploi / recrutement / formation",
            GROUPES_EMPLOI
        )
    ]

    for code, nom in departements["departements"].items():

        for annee in departements["annees"]:

            for numero, categorie, groupes_mots in groupes:

                for groupe in groupes_mots:

                    mots = " OR ".join(
                        f'"{mot}"'
                        for mot in groupe
                    )

                    texte = (
                        f"({mots}) "
                        f'"{nom}" '
                        f"{annee}"
                    )

                    requetes.append({
                        "departement": code,
                        "nom_departement": nom,
                        "annee": annee,
                        "categorie_numero": numero,
                        "categorie": categorie,
                        "requete": texte
                    })

    return requetes


# ============================================================
# RECHERCHE RSS
# ============================================================

def rechercher(requete):

    url = (
        "https://news.google.com/rss/search?"
        f"q={quote_plus(requete['requete'])}"
        "&hl=fr&gl=FR&ceid=FR:fr"
    )

    session = get_session()

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
                    "date_publication":
                        date_publication,
                    "departement":
                        requete["departement"],
                    "nom_departement":
                        requete["nom_departement"],
                    "annee_recherchee":
                        requete["annee"],
                    "categorie_numero":
                        requete["categorie_numero"],
                    "categorie":
                        requete["categorie"],
                    "requete":
                        requete["requete"]
                })

            return resultats

        except Exception:

            if tentative < MAX_RETRIES - 1:
                time.sleep(0.5)

    return []


# ============================================================
# SCORE DE PERTINENCE
# ============================================================

def calculer_score(resultat):

    texte = normaliser(
        resultat["titre"]
        + " "
        + resultat["source"]
    )

    score = 0

    # Termes très pertinents
    for terme in TERMES_FORTS:

        if terme in texte:
            score += 15

    # Termes positifs
    for terme in TERMES_POSITIFS:

        if terme in texte:
            score += 3

    # Bruit
    for terme in TERMES_BRUIT:

        if terme in texte:
            score -= 8

    # Exclusions fortes
    for terme in TERMES_EXCLUSION_FORTE:

        if terme in texte:
            score -= 40

    # Bonus catégorie
    if resultat["categorie_numero"] == 1:

        if any(
            terme in texte
            for terme in [
                "université",
                "école",
                "étudiant",
                "supérieur",
                "alumni",
                "portes ouvertes"
            ]
        ):
            score += 10

    if resultat["categorie_numero"] == 2:

        if any(
            terme in texte
            for terme in [
                "emploi",
                "recrutement",
                "formation",
                "job dating",
                "carrière"
            ]
        ):
            score += 10

    return score


# ============================================================
# CLASSIFICATION
# ============================================================

def classer(score):

    if score >= 25:
        return "Pertinent"

    if score >= 10:
        return "À vérifier"

    return "Non pertinent"


# ============================================================
# EXCLUSIONS STRICTES
# ============================================================

def exclusion_forte(resultat):

    texte = normaliser(
        resultat["titre"]
        + " "
        + resultat["source"]
    )

    for terme in TERMES_EXCLUSION_FORTE:

        if terme in texte:
            return True, terme

    return False, ""


# ============================================================
# DÉDOUBLONNAGE
# ============================================================

def dedoublonner(resultats):

    vus_url = set()
    vus_titre = set()

    uniques = []

    for resultat in resultats:

        url = normaliser(
            resultat["url"]
        )

        titre = normaliser(
            resultat["titre"]
        )

        titre_cle = re.sub(
            r"[^a-z0-9àâäéèêëîïôöùûüç ]",
            "",
            titre
        )

        if url in vus_url:
            continue

        if titre_cle in vus_titre:
            continue

        vus_url.add(url)
        vus_titre.add(titre_cle)

        uniques.append(resultat)

    return uniques


# ============================================================
# COLLECTE PARALLÈLE
# ============================================================

def collecter(requetes):

    resultats = []

    total = len(requetes)
    termine = 0

    debut = time.time()

    print()
    print("========================================")
    print("COLLECTE RSS OPTIMISÉE")
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

        futures = [
            executor.submit(
                rechercher,
                requete
            )
            for requete in requetes
        ]

        for future in as_completed(futures):

            termine += 1

            try:

                nouveaux = future.result()

                resultats.extend(
                    nouveaux
                )

            except Exception:
                pass

            if (
                termine % 20 == 0
                or termine == total
            ):

                print(
                    f"Progression : "
                    f"{termine}/{total} "
                    f"| résultats : "
                    f"{len(resultats)}"
                )

    duree = time.time() - debut

    print()
    print(
        f"Collecte terminée en "
        f"{duree:.1f} secondes"
    )

    return resultats


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

        exclusion, motif = (
            exclusion_forte(
                resultat
            )
        )

        if exclusion:

            resultat["score"] = -40
            resultat["statut"] = (
                "Non pertinent"
            )
            resultat["motif"] = motif

            non_pertinents.append(
                resultat
            )

            continue

        score = calculer_score(
            resultat
        )

        statut = classer(score)

        resultat["score"] = score
        resultat["statut"] = statut
        resultat["motif"] = ""

        if statut == "Pertinent":

            pertinents.append(
                resultat
            )

        elif statut == "À vérifier":

            a_verifier.append(
                resultat
            )

        else:

            non_pertinents.append(
                resultat
            )

    pertinents.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    a_verifier.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    return (
        pertinents,
        a_verifier,
        non_pertinents
    )


# ============================================================
# SAUVEGARDE
# ============================================================

def sauvegarder(
    pertinents,
    a_verifier,
    non_pertinents
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
        "date_publication",
        "departement",
        "nom_departement",
        "annee_recherchee",
        "categorie_numero",
        "categorie",
        "score",
        "statut",
        "motif",
        "requete",
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

    return (
        fichier_principal,
        fichier_verification,
        fichier_non_pertinent
    )


# ============================================================
# MAIN
# ============================================================

def main():

    debut = time.time()

    DATA.mkdir(
        exist_ok=True
    )

    OUTPUT.mkdir(
        exist_ok=True
    )

    print()
    print("========================================")
    print("VEILLE APEC")
    print("COLLECTEUR OPTIMISÉ")
    print("========================================")

    departements = load_yaml(
        "departements.yml"
    )

    requetes = construire_requetes(
        departements
    )

    print(
        f"\nRequêtes générées : "
        f"{len(requetes)}"
    )

    # Collecte
    resultats_bruts = collecter(
        requetes
    )

    print(
        f"\nRésultats bruts : "
        f"{len(resultats_bruts)}"
    )

    # Traitement
    (
        pertinents,
        a_verifier,
        non_pertinents
    ) = traiter(
        resultats_bruts
    )

    print()
    print("========================================")
    print("RÉSULTATS")
    print("========================================")

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

    # Sauvegarde
    fichiers = sauvegarder(
        pertinents,
        a_verifier,
        non_pertinents
    )

    duree = time.time() - debut

    print()
    print("========================================")
    print("FICHIERS")
    print("========================================")

    for fichier in fichiers:
        print(f"- {fichier}")

    print()
    print(
        f"Temps total : "
        f"{duree:.1f} secondes"
    )

    print()
    print("========================================")
    print("VEILLE TERMINÉE")
    print("========================================")


if __name__ == "__main__":
    main()
