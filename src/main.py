from pathlib import Path
from urllib.parse import quote_plus
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import threading
import time
import re
import shutil

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
# TERMES POSITIFS
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
    "salon étudiant",
    "forum étudiant",
    "enseignement supérieur",
    "études supérieures",
    "portes ouvertes",
    "journée portes ouvertes",
    "alumni",
    "rencontre entreprises"
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
    "apprentissage",
    "forum",
    "salon"
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
    "agriculture",
    "météo",
    "faits divers",
    "incendie",
    "accident",
    "politique",
    "exposition"
]

TERMES_EXCLUSION = [
    "chambre de commerce et d'industrie",
    "chambre de commerce",
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
    #
    # IMPORTANT :
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
                    "date_publication": (
                        date_publication
                    ),
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

def calculer_score(resultat):

    texte = normaliser(
        resultat["titre"]
        + " "
        + resultat["source"]
    )

    score = 0

    for terme in TERMES_FORTS:

        if terme in texte:
            score += 15

    for terme in TERMES_POSITIFS:

        if terme in texte:
            score += 3

    for terme in TERMES_BRUIT:

        if terme in texte:
            score -= 7

    for terme in TERMES_EXCLUSION:

        if terme in texte:
            score -= 40

    return score


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

        texte = normaliser(
            resultat["titre"]
            + " "
            + resultat["source"]
        )

        exclusion = False
        motif = ""

        for terme in TERMES_EXCLUSION:

            if terme in texte:

                exclusion = True
                motif = terme
                break

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

        resultat["score"] = score
        resultat["motif"] = ""

        if score >= 25:

            resultat["statut"] = (
                "Pertinent"
            )

            pertinents.append(
                resultat
            )

        elif score >= 8:

            resultat["statut"] = (
                "À vérifier"
            )

            a_verifier.append(
                resultat
            )

        else:

            resultat["statut"] = (
                "Non pertinent"
            )

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
        "date_publication",
        "departement",
        "nom_departement",
        "annee_recherchee",
        "categorie_numero",
        "categorie",
        "mot_cle",
        "score",
        "statut",
        "motif",
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
    print("COLLECTEUR RSS FIABILISÉ")
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

    print()
    print(
        f"Après dédoublonnage : "
        f"{len(pertinents) + len(a_verifier) + len(non_pertinents)}"
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
        f"Erreurs : {erreurs}"
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

        print("VEILLE TERMINÉE AVEC SUCCÈS")

    else:

        print("VEILLE NON VALIDÉE")

    print("========================================")


if __name__ == "__main__":
    main()
