from pathlib import Path
from urllib.parse import quote_plus
from datetime import datetime

import feedparser
import pandas as pd
import yaml


ROOT = Path(__file__).resolve().parent.parent

CONFIG = ROOT / "config"
DATA = ROOT / "data"
OUTPUT = ROOT / "output"


def load_yaml(filename):
    path = CONFIG / filename

    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def construire_requetes(departements, categories):
    requetes = []

    for code, nom in departements["departements"].items():

        for annee in departements["annees"]:

            for numero, categorie in categories["categories"].items():

                for mot_cle in categorie["mots_cles"]:

                    requetes.append(
                        {
                            "departement": code,
                            "nom_departement": nom,
                            "annee": annee,
                            "categorie_numero": numero,
                            "categorie": categorie["nom"],
                            "mot_cle": mot_cle,
                        }
                    )

    return requetes


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

    flux = feedparser.parse(url)

    resultats = []

    for article in flux.entries:

        titre = article.get("title", "").strip()
        lien = article.get("link", "").strip()
        date_publication = article.get("published", "").strip()
        source = ""

        if "source" in article:
            source = article.source.get("title", "")

        if not titre or not lien:
            continue

        resultats.append(
            {
                "titre": titre,
                "url": lien,
                "source": source,
                "date_publication": date_publication,
                "departement": requete["departement"],
                "nom_departement": requete["nom_departement"],
                "annee_recherchee": requete["annee"],
                "categorie_numero": requete["categorie_numero"],
                "categorie": requete["categorie"],
                "mot_cle": requete["mot_cle"],
            }
        )

    return resultats


def supprimer_doublons(resultats):

    uniques = {}
    resultats_final = []

    for resultat in resultats:

        cle = resultat["url"]

        if cle not in uniques:
            uniques[cle] = True
            resultats_final.append(resultat)

    return resultats_final


def appliquer_exclusions(resultats, exclusions):

    tous_termes = []

    for groupe in exclusions["exclusions"].values():
        tous_termes.extend(groupe)

    resultats_valides = []
    resultats_exclus = []

    for resultat in resultats:

        texte = (
            resultat["titre"]
            + " "
            + resultat["source"]
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
            resultats_exclus.append(resultat)

        else:

            resultats_valides.append(resultat)

    return resultats_valides, resultats_exclus


def main():

    DATA.mkdir(exist_ok=True)
    OUTPUT.mkdir(exist_ok=True)

    departements = load_yaml("departements.yml")
    categories = load_yaml("categories.yml")
    exclusions = load_yaml("exclusions.yml")

    print("========================================")
    print("VEILLE APEC - COLLECTE RSS")
    print("========================================")

    requetes = construire_requetes(
        departements,
        categories
    )

    print(f"\nNombre de requêtes prévues : {len(requetes)}")

    resultats = []

    for index, requete in enumerate(requetes, start=1):

        print(
            f"[{index}/{len(requetes)}] "
            f"{requete['departement']} - "
            f"{requete['annee']} - "
            f"{requete['mot_cle']}"
        )

        try:

            nouveaux = rechercher_google_news(requete)

            resultats.extend(nouveaux)

        except Exception as erreur:

            print(
                "Erreur pendant la recherche : "
                f"{erreur}"
            )

    print(
        f"\nRésultats bruts récupérés : "
        f"{len(resultats)}"
    )

    resultats = supprimer_doublons(resultats)

    print(
        f"Résultats après dédoublonnage : "
        f"{len(resultats)}"
    )

    resultats_valides, resultats_exclus = appliquer_exclusions(
        resultats,
        exclusions
    )

    print(
        f"Résultats conservés : "
        f"{len(resultats_valides)}"
    )

    print(
        f"Résultats exclus : "
        f"{len(resultats_exclus)}"
    )

    maintenant = datetime.now().isoformat()

    for resultat in resultats_valides:
        resultat["date_collecte"] = maintenant
        resultat["statut"] = "À vérifier"

    for resultat in resultats_exclus:
        resultat["date_collecte"] = maintenant
        resultat["statut"] = "Non pertinent"

    fichier_resultats = (
        DATA / "resultats_rss.csv"
    )

    fichier_exclus = (
        DATA / "resultats_exclus.csv"
    )

    pd.DataFrame(resultats_valides).to_csv(
        fichier_resultats,
        index=False,
        encoding="utf-8-sig"
    )

    pd.DataFrame(resultats_exclus).to_csv(
        fichier_exclus,
        index=False,
        encoding="utf-8-sig"
    )

    print("\nFichiers créés :")

    print(f"- {fichier_resultats}")
    print(f"- {fichier_exclus}")

    print("\n========================================")
    print("VEILLE APEC - FIN")
    print("========================================")


if __name__ == "__main__":
    main()
