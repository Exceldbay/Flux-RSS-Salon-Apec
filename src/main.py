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
    with open(CONFIG / filename, "r", encoding="utf-8") as f:
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
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 Chrome/131.0 Safari/537.36"
            )
        })
        thread_local.session = session
    return thread_local.session


# ============================================================
# REQUÊTES RSS
# ============================================================

REQUETES_EMPLOI = [
    "salon emploi",
    "forum emploi",
    "job dating",
    "forum recrutement",
    "salon recrutement",
    "salon formation",
    "forum formation",
    "emploi cadre",
]

REQUETES_SUPERIEUR = [
    "salon étudiant",
    "forum étudiant",
    "enseignement supérieur",
    "études supérieures",
    "portes ouvertes supérieur",
    "journée portes ouvertes université",
    "alumni",
    "rencontre entreprises étudiants",
]


# ============================================================
# VOCABULAIRE DE SCORING
# ============================================================

EVENTS_EMPLOI = [
    "salon emploi",
    "forum emploi",
    "job dating",
    "job-dating",
    "forum recrutement",
    "salon recrutement",
    "forum formation",
    "salon formation",
    "forum emploi et formation",
    "forum emploi-formation",
    "journée recrutement",
]

EVENTS_SUPERIEUR = [
    "salon étudiant",
    "salon des études supérieures",
    "salon de l'enseignement supérieur",
    "salon enseignement supérieur",
    "forum étudiant",
    "forum de l'orientation",
    "forum orientation",
    "salon orientation",
    "portes ouvertes",
    "journée portes ouvertes",
    "rencontre alumni",
    "rassemblement alumni",
    "rencontre entreprises étudiants",
]

TERMES_EMPLOI = [
    "emploi",
    "recrutement",
    "recruteur",
    "recruteurs",
    "demandeurs d'emploi",
    "demandeurs d’emploi",
    "candidats",
    "métiers",
    "carrière",
    "alternance",
    "apprentissage",
]

TERMES_FORMATION = [
    "formation",
    "formation professionnelle",
    "orientation",
    "études",
    "études supérieures",
]

TERMES_SUPERIEUR = [
    "enseignement supérieur",
    "études supérieures",
    "étudiant",
    "étudiants",
    "université",
    "école d'ingénieurs",
    "école d’ingénieurs",
    "école de commerce",
    "alumni",
    "anciens élèves",
]

TERMES_APEC = [
    "apec",
    "cadre",
    "cadres",
]

BRUIT_FORT = [
    "budget",
    "président",
    "présidence",
    "élection",
    "déficit",
    "manifestation",
    "grève",
    "polémique",
    "chantier",
    "bâtiment",
    "extension",
    "congé",
    "enquête",
    "ministre",
    "loi",
    "salaire",
    "classement",
    "logement",
    "coût de la vie",
    "intentions de recrutement",
]

EVENEMENTS_HORS_PERIMETRE = [
    "salon des seniors",
    "salon senior",
    "thalasso",
    "cure thermale",
    "salon des vins",
    "salon du vin",
    "salon des minéraux",
    "minéraux fossiles",
    "pierres fines",
    "brocante",
    "vide-grenier",
    "marché",
    "festival",
    "concert",
    "spectacle",
    "football",
    "rugby",
    "basket",
    "fête du cuir",
]

# Villes / territoires explicitement identifiés comme hors périmètre
# lors de l'audit des 264 résultats potentiellement utiles.
VILLES_HORS_CIBLE = [
    "tours",
    "nice",
    "agen",
    "mulhouse",
    "bordeaux",
    "saint-nazaire",
    "la réunion",
    "montpellier",
    "lille",
    "chambéry",
    "saint-malo",
    "langon",
    "colmar",
    "condrieu",
    "alès",
    "ales",
    "paris",
    "pyrénées-atlantiques",
    "gironde",
    "alsace",
]

VILLES_DEPARTEMENTS = {
    "16": [
        "angoulême", "angouleme", "cognac", "barbezieux",
        "la rochefoucauld", "jarnac", "ruffec"
    ],
    "17": [
        "la rochelle", "rochefort", "saintes", "royan",
        "jonzac", "châtelaillon", "chatelaillon", "pons",
        "surgères", "saujon", "saint-sulpice-de-royan",
        "arvert", "la palmyre"
    ],
    "19": [
        "brive", "tulle", "ussel", "égletons", "egletons", "objat"
    ],
    "23": [
        "guéret", "gueret", "aubusson", "la souterraine",
        "bourganeuf", "felletin", "gouzon", "argenton-sur-creuse"
    ],
    "79": [
        "niort", "bressuire", "parthenay", "thouars",
        "melle", "la crèche", "la creche", "prahecq"
    ],
    "86": [
        "poitiers", "châtellerault", "chatellerault", "loudun",
        "civray", "montmorillon", "chasseneuil", "jaunay-marigny",
        "saint-benoît", "saint-benoit", "gençay", "gencay",
        "saint-pardoux"
    ],
    "87": [
        "limoges", "panazol", "saint-junien", "bellac",
        "isle", "couzeix", "ambazac"
    ],
}

NOMS_DEPARTEMENTS = {
    "16": "charente",
    "17": "charente-maritime",
    "19": "corrèze",
    "23": "creuse",
    "79": "deux-sèvres",
    "86": "vienne",
    "87": "haute-vienne",
}


# ============================================================
# OUTILS TEXTE
# ============================================================

def normaliser(texte):
    texte = str(texte or "").lower()
    texte = re.sub(r"\s+", " ", texte)
    return texte.strip()


def contient_terme(texte, termes):
    texte = normaliser(texte)
    return [terme for terme in termes if terme in texte]


def extraire_annees(texte):
    return sorted(set(re.findall(r"\b20(?:25|26|27|28)\b", normaliser(texte))))


def detecter_geographie(resultat):
    code = str(resultat["departement"])
    texte = normaliser(
        resultat.get("titre", "")
        + " "
        + resultat.get("description", "")
    )

    nom_departement = NOMS_DEPARTEMENTS.get(code, "")
    villes = VILLES_DEPARTEMENTS.get(code, [])

    departement_present = nom_departement in texte
    villes_presentes = [v for v in villes if v in texte]

    hors_cible = [v for v in VILLES_HORS_CIBLE if v in normaliser(resultat.get("titre", ""))]

    positif = departement_present or bool(villes_presentes)

    return positif, departement_present, villes_presentes, hors_cible


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
                    "categorie": "Emploi / recrutement / formation",
                    "mot_cle": mot,
                })

            for mot in REQUETES_SUPERIEUR:
                requetes.append({
                    "departement": code,
                    "nom_departement": nom,
                    "annee": annee,
                    "categorie_numero": 1,
                    "categorie": "Salon étudiant / enseignement supérieur",
                    "mot_cle": mot,
                })

    return requetes


# ============================================================
# GOOGLE NEWS RSS
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
        "&hl=fr"
        "&gl=FR"
        "&ceid=FR:fr"
    )

    session = get_session()
    derniere_erreur = ""

    for tentative in range(MAX_RETRIES):
        try:
            response = session.get(url, timeout=TIMEOUT)
            response.raise_for_status()

            flux = feedparser.parse(response.content)
            resultats = []

            for article in flux.entries:
                titre = article.get("title", "").strip()
                lien = article.get("link", "").strip()
                date_publication = article.get("published", "").strip()

                source = ""
                if hasattr(article, "source"):
                    source = article.source.get("title", "")

                description = article.get("summary", "")
                if not description:
                    description = article.get("description", "")

                if not titre or not lien:
                    continue

                resultats.append({
                    "titre": titre,
                    "url": lien,
                    "source": source,
                    "description": normaliser(description),
                    "date_publication": date_publication,
                    "departement": requete["departement"],
                    "nom_departement": requete["nom_departement"],
                    "annee_recherchee": requete["annee"],
                    "categorie_numero": requete["categorie_numero"],
                    "categorie": requete["categorie"],
                    "mot_cle": requete["mot_cle"],
                })

            return {
                "ok": True,
                "resultats": resultats,
                "erreur": "",
            }

        except Exception as erreur:
            derniere_erreur = str(erreur)
            if tentative < MAX_RETRIES - 1:
                time.sleep(0.5)

    return {
        "ok": False,
        "resultats": [],
        "erreur": derniere_erreur,
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
    print(f"Requêtes : {total}")
    print(f"Workers : {MAX_WORKERS}")
    print()

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(rechercher_google_news, requete): requete
            for requete in requetes
        }

        for future in as_completed(futures):
            termine += 1

            try:
                resultat = future.result()

                if resultat["ok"]:
                    if resultat["resultats"]:
                        requetes_avec_resultats += 1
                    resultats.extend(resultat["resultats"])
                else:
                    erreurs += 1

            except Exception:
                erreurs += 1

            if termine % 20 == 0 or termine == total:
                print(
                    f"Progression : {termine}/{total} "
                    f"| résultats : {len(resultats)} "
                    f"| erreurs : {erreurs}"
                )

    duree = time.time() - debut

    print()
    print(f"Collecte terminée en {duree:.1f} secondes")
    print(
        "Requêtes ayant retourné au moins un résultat : "
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
        url = normaliser(resultat.get("url", ""))
        titre = normaliser(resultat.get("titre", ""))

        titre_cle = re.sub(
            r"[^a-z0-9àâäéèêëîïôöùûüç ]",
            "",
            titre
        )

        if not url:
            continue

        if url in urls or titre_cle in titres:
            continue

        urls.add(url)
        titres.add(titre_cle)
        uniques.append(resultat)

    return uniques


# ============================================================
# SCORING RENFORCÉ APRÈS AUDIT
# ============================================================

def calculer_score(resultat):

    titre = normaliser(resultat.get("titre", ""))
    description = normaliser(resultat.get("description", ""))
    source = normaliser(resultat.get("source", ""))

    texte = f"{titre} {description}"
    code_dept = str(resultat.get("departement", ""))
    categorie = int(resultat.get("categorie_numero", 0))
    annee_recherchee = str(resultat.get("annee_recherchee", ""))

    score = 0
    raisons = []

    events_emploi = contient_terme(texte, EVENTS_EMPLOI)
    events_superieur = contient_terme(texte, EVENTS_SUPERIEUR)

    termes_emploi = contient_terme(texte, TERMES_EMPLOI)
    termes_formation = contient_terme(texte, TERMES_FORMATION)
    termes_superieur = contient_terme(texte, TERMES_SUPERIEUR)
    termes_apec = contient_terme(texte, TERMES_APEC)

    bruit = contient_terme(texte, BRUIT_FORT)
    hors_perimetre = contient_terme(texte, EVENEMENTS_HORS_PERIMETRE)

    geo_ok, dept_present, villes_presentes, hors_geo = detecter_geographie(resultat)

    annees = extraire_annees(texte)

    # --------------------------------------------------------
    # 1. NATURE D'ÉVÉNEMENT
    # --------------------------------------------------------

    event_emploi = bool(events_emploi)
    event_superieur = bool(events_superieur)

    if event_emploi or event_superieur:
        score += 10
        raisons.append(
            "Événement identifié : "
            + ", ".join((events_emploi + events_superieur)[:4])
        )
    else:
        score -= 8
        raisons.append("Pas de format événementiel explicite")

    # --------------------------------------------------------
    # 2. ALIGNEMENT CATÉGORIE
    # --------------------------------------------------------

    if categorie == 2:
        if event_emploi:
            score += 4
            raisons.append("Événement emploi/recrutement/formation")
        if termes_emploi:
            score += 6
            raisons.append(
                "Emploi : " + ", ".join(termes_emploi[:4])
            )
        if termes_formation:
            score += 3
            raisons.append(
                "Formation : " + ", ".join(termes_formation[:3])
            )

    elif categorie == 1:
        if event_superieur:
            score += 4
            raisons.append("Événement enseignement supérieur")
        if termes_superieur:
            score += 5
            raisons.append(
                "Supérieur : " + ", ".join(termes_superieur[:4])
            )
        if termes_formation:
            score += 3
            raisons.append(
                "Orientation/formation : "
                + ", ".join(termes_formation[:3])
            )

    # --------------------------------------------------------
    # 3. GÉOGRAPHIE — OBLIGATOIRE POUR UN PERTINENT
    # --------------------------------------------------------

    if geo_ok:
        score += 5
        if villes_presentes:
            raisons.append(
                "Ville cible : " + ", ".join(villes_presentes[:3])
            )
        elif dept_present:
            raisons.append("Département cible")
    else:
        score -= 6
        raisons.append("Localisation cible non démontrée")

    if hors_geo:
        score -= 14
        raisons.append(
            "Hors périmètre détecté : "
            + ", ".join(hors_geo[:3])
        )

    # --------------------------------------------------------
    # 4. HORS PÉRIMÈTRE THÉMATIQUE
    # --------------------------------------------------------

    if hors_perimetre:
        score -= 18
        raisons.append(
            "Événement hors périmètre : "
            + ", ".join(hors_perimetre[:3])
        )

    # --------------------------------------------------------
    # 5. BRUIT JOURNALISTIQUE
    # --------------------------------------------------------

    if bruit:
        score -= 10
        raisons.append(
            "Bruit / actualité générale : "
            + ", ".join(bruit[:4])
        )

    # Une mention Apec seule ne doit plus transformer
    # une actualité générale en événement pertinent.
    if termes_apec and not (event_emploi or event_superieur):
        score -= 6
        raisons.append("Mention Apec sans événement identifié")

    # --------------------------------------------------------
    # 6. COHÉRENCE ANNÉE
    # --------------------------------------------------------

    if annee_recherchee in annees:
        score += 3
        raisons.append(
            f"Année {annee_recherchee} explicitement détectée"
        )
    elif annees:
        score -= 6
        raisons.append(
            "Année différente détectée : "
            + ", ".join(annees)
        )

    if annee_recherchee in normaliser(
        resultat.get("date_publication", "")
    ):
        score += 1

    # --------------------------------------------------------
    # 7. DÉCISION
    # --------------------------------------------------------

    # Un résultat ne peut être "Pertinent" que s'il présente
    # simultanément un événement + une géographie cible.
    evenement = event_emploi or event_superieur

    if (
        evenement
        and geo_ok
        and not hors_geo
        and not hors_perimetre
        and score >= 15
    ):
        statut = "Pertinent"
        raison = "Événement local clairement identifié"

    elif (
        evenement
        and not hors_perimetre
        and score >= 8
    ):
        statut = "À vérifier"
        raison = "Événement potentiel, validation nécessaire"

    else:
        statut = "Non pertinent"
        raison = "Signal insuffisant ou hors périmètre"

    details = " | ".join(raisons)

    return {
        "score": score,
        "score_emploi": 6 if termes_emploi else 0,
        "score_formation": 3 if termes_formation else 0,
        "score_superieur": 5 if termes_superieur else 0,
        "score_apec": 1 if termes_apec else 0,
        "score_geographie": 5 if geo_ok else -6,
        "score_date": 3 if annee_recherchee in annees else 0,
        "score_bruit": -10 if bruit else 0,
        "score_complementaire": 4 if (event_emploi or event_superieur) else 0,
        "bonus_titre": 2 if (events_emploi or events_superieur) else 0,
        "statut": statut,
        "raison": raison,
        "motif": "",
        "details_score": details,
        "evenement_detecte": "OUI" if evenement else "NON",
        "geographie_verifiee": "OUI" if geo_ok else "NON",
        "annee_detectee": ", ".join(annees),
        "hors_perimetre": ", ".join(hors_perimetre),
        "hors_geographie": ", ".join(hors_geo),
    }


# ============================================================
# TRAITEMENT
# ============================================================

def traiter(resultats):
    resultats = dedoublonner(resultats)

    pertinents = []
    a_verifier = []
    non_pertinents = []

    for resultat in resultats:
        score_data = calculer_score(resultat)
        resultat.update(score_data)

        if resultat["statut"] == "Pertinent":
            pertinents.append(resultat)
        elif resultat["statut"] == "À vérifier":
            a_verifier.append(resultat)
        else:
            non_pertinents.append(resultat)

    pertinents.sort(key=lambda x: x["score"], reverse=True)
    a_verifier.sort(key=lambda x: x["score"], reverse=True)
    non_pertinents.sort(key=lambda x: x["score"], reverse=True)

    return pertinents, a_verifier, non_pertinents


# ============================================================
# SAUVEGARDE
# ============================================================

def sauvegarder(
    pertinents,
    a_verifier,
    non_pertinents,
    resultats_bruts
):
    if not resultats_bruts:
        return False

    date_collecte = datetime.now().isoformat()

    tous = pertinents + a_verifier + non_pertinents

    for resultat in tous:
        resultat["date_collecte"] = date_collecte

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
        "evenement_detecte",
        "geographie_verifiee",
        "annee_detectee",
        "hors_perimetre",
        "hors_geographie",
        "date_collecte",
    ]

    def dataframe(resultats):
        df = pd.DataFrame(resultats)

        for colonne in colonnes:
            if colonne not in df.columns:
                df[colonne] = ""

        return df[colonnes]

    DATA.mkdir(exist_ok=True)
    OUTPUT.mkdir(exist_ok=True)

    dataframe(pertinents).to_csv(
        DATA / "resultats_rss.csv",
        index=False,
        encoding="utf-8-sig"
    )

    dataframe(a_verifier).to_csv(
        DATA / "a_verifier.csv",
        index=False,
        encoding="utf-8-sig"
    )

    dataframe(non_pertinents).to_csv(
        DATA / "non_pertinent.csv",
        index=False,
        encoding="utf-8-sig"
    )

    # Rapport synthétique destiné au contrôle de la qualité
    synthese = pd.DataFrame([
        {
            "indicateur": "Résultats bruts",
            "valeur": len(resultats_bruts),
        },
        {
            "indicateur": "Après dédoublonnage",
            "valeur": len(tous),
        },
        {
            "indicateur": "Pertinents",
            "valeur": len(pertinents),
        },
        {
            "indicateur": "À vérifier",
            "valeur": len(a_verifier),
        },
        {
            "indicateur": "Non pertinents",
            "valeur": len(non_pertinents),
        },
    ])

    synthese.to_csv(
        DATA / "synthese_rss.csv",
        index=False,
        encoding="utf-8-sig"
    )

    return True


# ============================================================
# MAIN
# ============================================================

def main():

    debut_total = time.time()

    DATA.mkdir(exist_ok=True)
    OUTPUT.mkdir(exist_ok=True)

    print()
    print("========================================")
    print("VEILLE APEC")
    print("COLLECTEUR RSS + SCORING APEC AUDITÉ")
    print("========================================")

    departements = load_yaml("departements.yml")
    requetes = construire_requetes(departements)

    print()
    print(f"Requêtes générées : {len(requetes)}")

    resultats_bruts, erreurs = collecter(requetes)

    print()
    print("========================================")
    print("TRAITEMENT")
    print("========================================")
    print(f"Résultats bruts : {len(resultats_bruts)}")

    if not resultats_bruts:
        print()
        print("⚠️ AUCUN RÉSULTAT RSS.")
        print("Les anciens résultats ne sont pas écrasés.")
        print(f"Erreurs HTTP/RSS : {erreurs}")
        return

    pertinents, a_verifier, non_pertinents = traiter(
        resultats_bruts
    )

    total_dedoublonne = (
        len(pertinents)
        + len(a_verifier)
        + len(non_pertinents)
    )

    print()
    print(f"Après dédoublonnage : {total_dedoublonne}")
    print(f"Pertinents : {len(pertinents)}")
    print(f"À vérifier : {len(a_verifier)}")
    print(f"Non pertinents : {len(non_pertinents)}")

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
    print(f"Erreurs : {erreurs}")
    print(f"Temps total : {duree_total:.1f} secondes")
    print(
        "Sauvegarde validée : "
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
