from pathlib import Path
from urllib.parse import quote_plus
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
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

# Nombre de recherches simultanées.
# 12 = bon compromis vitesse / stabilité.
MAX_WORKERS = 12

# Timeout par requête.
TIMEOUT = 10

# Petite pause entre certaines opérations.
RETRY_DELAY = 0.5

# Nombre maximum de tentatives par requête.
MAX_RETRIES = 2


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

thread_local = threading.local()


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
                "(compatible; Veille-APEC-RSS/1.0)"
            )
        })

        thread_local.session = session

    return thread_local.session


# ============================================================
# CONSTRUCTION DES REQUÊTES
# ============================================================

def construire_requetes(departements, categories):

    requetes = []

    for code, nom in departements["departements"].items():

        for annee in departements["annees"]:

            for numero, categorie in categories["categories"].items():

                for mot_cle in categorie["mots_cles"]:

                    requetes.append({
                        "departement": code,
                        "nom_departement": nom,
                        "annee": annee,
                        "categorie_numero": numero,
                        "categorie": categorie["nom"],
                        "mot_cle": mot_cle
                    })

    return requetes


# ============================================================
# RECHERCHE GOOGLE NEWS RSS
# ============================================================

def rechercher_google_news(requete):

    texte = (
        f'"{requete["mot_cle"]}" '
        f'"{requete["nom_departement"]}" '
        f'{requete["annee"]}'
    )

    url = (
        "https://news.google.com/rss/search?"
        f"q={quote_plus(texte)}"
        "&hl=fr&gl=FR&ceid=FR:fr"
    )

    session = get_session()

    derniere_erreur = None

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
                    "date_publication": date_publication,
                    "departement": requete["departement"],
                    "nom_departement": requete["nom_departement"],
                    "annee_recherchee": requete["annee"],
                    "categorie_numero": requete["categorie_numero"],
                    "categorie": requete["categorie"],
                    "mot_cle": requete["mot_cle"]
                })

            return {
                "statut": "OK",
                "resultats": resultats
            }

        except Exception as erreur:

            derniere_erreur = str(erreur)

            if tentative < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)

    return {
        "statut": "ERREUR",
        "resultats": [],
        "erreur": derniere_erreur
    }


# ============================================================
# DÉDOUBLONNAGE
# ============================================================

def supprimer_doublons(resultats):

    uniques = {}
    resultats_final = []

    for resultat in resultats:

        url = resultat.get("url", "").strip()

        if not url:
            continue

        if url not in uniques:

            uniques[url] = True
            resultats_final.append(resultat)

    return resultats_final


# ============================================================
# EXCLUSIONS
# ============================================================

def appliquer_exclusions(resultats, exclusions):

    tous_termes = []

    for groupe in exclusions["exclusions"].values():

        if isinstance(groupe, list):
            tous_termes.extend(groupe)

    resultats_valides = []
    resultats_exclus = []

    for resultat in resultats:

        texte = (
            resultat.get("titre", "")
            + " "
            + resultat.get("source", "")
        ).lower()

        exclusion = False
        terme_trouve = ""

        for terme in tous_termes:

            if terme.lower() in texte:

                exclusion = True
                terme_trouve = terme
                break

        if exclusion:

            resultat["motif_exclusion"] = terme_trouve
            resultat["statut"] = "Non pertinent"

            resultats_exclus.append(resultat)

        else:

            resultat["statut"] = "À vérifier"

            resultats_valides.append(resultat)

    return resultats_valides, resultats_exclus


# ============================================================
# COLLECTE PARALLÈLE
# ============================================================

def collecter(requetes):

    resultats = []
    erreurs = 0

    total = len(requetes)

    debut = time.time()

    print()
    print("========================================")
    print("COLLECTE PARALLÈLE")
    print("========================================")
    print(f"Requêtes : {total}")
    print(f"Connexions simultanées : {MAX_WORKERS}")
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

        termine = 0

        for future in as_completed(futures):

            requete = futures[future]

            termine += 1

            try:

                resultat = future.result()

                if resultat["statut"] == "OK":

                    resultats.extend(
                        resultat["resultats"]
                    )

                else:

                    erreurs += 1

            except Exception as erreur:

                erreurs += 1

                print(
                    f"Erreur : "
                    f"{requete['departement']} "
                    f"{requete['annee']} "
                    f"{erreur}"
                )

            if termine % 25 == 0 or termine == total:

                duree = time.time() - debut

                print(
                    f"Progression : "
                    f"{termine}/{total} "
                    f"| résultats : {len(resultats)} "
                    f"| erreurs : {erreurs} "
                    f"| temps : {duree:.1f}s"
                )

    duree = time.time() - debut

    print()
    print(
        f"Collecte terminée en {duree:.1f} secondes."
    )

    return resultats, erreurs


# ============================================================
# SAUVEGARDE
# ============================================================

def sauvegarder(resultats_valides, resultats_exclus):

    maintenant = datetime.now().isoformat()

    for resultat in resultats_valides:
        resultat["date_collecte"] = maintenant

    for resultat in resultats_exclus:
        resultat["date_collecte"] = maintenant

    fichier_resultats = (
        DATA / "resultats_rss.csv"
    )

    fichier_exclus = (
        DATA / "resultats_exclus.csv"
    )

    pd.DataFrame(
        resultats_valides
    ).to_csv(
        fichier_resultats,
        index=False,
        encoding="utf-8-sig"
    )

    pd.DataFrame(
        resultats_exclus
    ).to_csv(
        fichier_exclus,
        index=False,
        encoding="utf-8-sig"
    )

    return fichier_resultats, fichier_exclus


# ============================================================
# PROGRAMME PRINCIPAL
# ============================================================

def main():

    debut_total = time.time()

    DATA.mkdir(exist_ok=True)
    OUTPUT.mkdir(exist_ok=True)

    print()
    print("========================================")
    print("VEILLE APEC - COLLECTEUR RSS OPTIMISÉ")
    print("========================================")

    print("\nChargement de la configuration...")

    departements = load_yaml(
        "departements.yml"
    )

    categories = load_yaml(
        "categories.yml"
    )

    exclusions = load_yaml(
        "exclusions.yml"
    )

    requetes = construire_requetes(
        departements,
        categories
    )

    print(
        f"Nombre total de requêtes : "
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
    # DÉDOUBLONNAGE
    # --------------------------------------------------------

    resultats_dedoublonnes = (
        supprimer_doublons(
            resultats_bruts
        )
    )

    print(
        f"Après dédoublonnage : "
        f"{len(resultats_dedoublonnes)}"
    )

    # --------------------------------------------------------
    # EXCLUSIONS
    # --------------------------------------------------------

    resultats_valides, resultats_exclus = (
        appliquer_exclusions(
            resultats_dedoublonnes,
            exclusions
        )
    )

    print(
        f"Résultats conservés : "
        f"{len(resultats_valides)}"
    )

    print(
        f"Résultats exclus : "
        f"{len(resultats_exclus)}"
    )

    # --------------------------------------------------------
    # SAUVEGARDE
    # --------------------------------------------------------

    fichier_resultats, fichier_exclus = (
        sauvegarder(
            resultats_valides,
            resultats_exclus
        )
    )

    duree_total = time.time() - debut_total

    print()
    print("========================================")
    print("TERMINÉ")
    print("========================================")

    print(
        f"Temps total : "
        f"{duree_total:.1f} secondes"
    )

    print(
        f"Erreurs de requêtes : "
        f"{erreurs}"
    )

    print()
    print("Fichiers créés :")
    print(f"- {fichier_resultats}")
    print(f"- {fichier_exclus}")

    print()
    print("========================================")


if __name__ == "__main__":
    main()
