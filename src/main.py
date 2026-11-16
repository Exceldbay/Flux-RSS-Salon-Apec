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
from bs4 import BeautifulSoup

from ia import analyser_resultat


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
MAX_IA_ANALYSES = 50

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
    ],
    "87": [
        "limoges", "panazol", "saint-junien", "bellac",
        "isle", "couzeix", "ambazac", "saint-pardoux"
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
    return sorted(set(re.findall(r"\b20(?:26|27|28)\b", normaliser(texte))))


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
# ENRICHISSEMENT DES PAGES SOURCES
# ============================================================

def extraire_meta(soup, noms):
    for nom in noms:
        balise = soup.find("meta", attrs={"name": nom})
        if balise and balise.get("content"):
            return balise.get("content", "").strip()
        balise = soup.find("meta", attrs={"property": nom})
        if balise and balise.get("content"):
            return balise.get("content", "").strip()
    return ""


def enrichir_page(resultat):
    """
    Récupère des informations complémentaires sur la page source.
    L'enrichissement est limité aux candidats IA afin de ne pas
    ralentir la collecte RSS.
    """
    url = resultat.get("url", "")
    if not url:
        return resultat

    try:
        session = get_session()
        response = session.get(url, timeout=TIMEOUT, allow_redirects=True)
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "").lower()
        if "html" not in content_type:
            return resultat

        soup = BeautifulSoup(response.content, "lxml")

        titre_page = soup.title.get_text(" ", strip=True) if soup.title else ""
        description_page = extraire_meta(
            soup, ["description", "og:description", "twitter:description"]
        )
        date_evenement = extraire_meta(
            soup,
            ["event:start_time", "event:start_date", "startDate",
             "date", "datePublished"]
        )
        lieu_page = extraire_meta(
            soup, ["event:location", "location", "place"]
        )
        organisateur_page = extraire_meta(
            soup, ["author", "article:author", "og:site_name", "publisher"]
        )

        # Données structurées JSON-LD
        import json
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string or script.get_text())
                elements = data if isinstance(data, list) else [data]

                for element in elements:
                    if not isinstance(element, dict):
                        continue

                    if not date_evenement:
                        date_evenement = str(
                            element.get("startDate")
                            or element.get("datePublished")
                            or ""
                        ).strip()

                    if not lieu_page:
                        location = element.get("location", "")
                        if isinstance(location, dict):
                            lieu_page = str(
                                location.get("name")
                                or location.get("address")
                                or ""
                            ).strip()
                        elif location:
                            lieu_page = str(location).strip()

                    if not organisateur_page:
                        org = element.get("organizer", "")
                        if isinstance(org, dict):
                            organisateur_page = str(org.get("name") or "").strip()
                        elif org:
                            organisateur_page = str(org).strip()

            except Exception:
                continue

        resultat["url_finale"] = response.url
        resultat["titre_page"] = titre_page[:500]
        resultat["description_page"] = normaliser(description_page)[:3000]
        resultat["date_evenement_page"] = date_evenement[:100]
        resultat["lieu_page"] = normaliser(lieu_page)[:300]
        resultat["organisateur_page"] = normaliser(organisateur_page)[:300]
        resultat["page_enrichie"] = "OUI"

        # L'IA reçoit les informations complémentaires sans modifier
        # le score Python déjà calculé.
        enrichissement = " ".join(
            x for x in [
                titre_page,
                description_page,
                date_evenement,
                lieu_page,
                organisateur_page,
            ] if x
        )

        if enrichissement:
            resultat["description"] = (
                normaliser(resultat.get("description", ""))
                + " "
                + normaliser(enrichissement)
            ).strip()

    except Exception as e:
        resultat["page_enrichie"] = "NON"
        resultat["erreur_enrichissement"] = str(e)[:300]

    return resultat


def enrichir_candidats_ia(candidats):
    """
    Enrichit uniquement les candidats susceptibles d'être analysés
    par Gemini. Cela limite les requêtes HTTP supplémentaires.
    """
    print()
    print("========================================")
    print("ENRICHISSEMENT DES PAGES")
    print("========================================")
    print(f"Candidats à enrichir : {min(len(candidats), MAX_IA_ANALYSES)}")

    debut = time.time()
    cible = candidats[:MAX_IA_ANALYSES]

    with ThreadPoolExecutor(max_workers=min(8, MAX_WORKERS)) as executor:
        futures = [executor.submit(enrichir_page, resultat) for resultat in cible]

        for future in as_completed(futures):
            try:
                future.result()
            except Exception:
                pass

    duree = time.time() - debut
    enrichis = sum(1 for r in cible if r.get("page_enrichie") == "OUI")

    print(f"Pages enrichies : {enrichis}/{len(cible)}")
    print(f"Temps enrichissement : {duree:.1f} secondes")

    return candidats


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
        corriger_departement_par_ville(resultat)
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



def corriger_departement_par_ville(resultat):
    """Corrige le département de recherche lorsqu'une ville cible unique
    et explicitement reconnue apparaît dans le contenu RSS.

    Cela évite qu'une requête lancée pour un département conserve ce code
    lorsque l'article concerne en réalité un autre département cible.
    """
    texte = normaliser(
        resultat.get("titre", "") + " " + resultat.get("description", "")
    )

    correspondances = []
    for code, villes in VILLES_DEPARTEMENTS.items():
        for ville in villes:
            if ville and ville in texte:
                correspondances.append(code)

    codes = sorted(set(correspondances))

    # Une seule correspondance départementale = correction sûre.
    if len(codes) == 1:
        code = codes[0]
        resultat["departement"] = code
        resultat["nom_departement"] = NOMS_DEPARTEMENTS.get(code, resultat.get("nom_departement", ""))
        resultat["geographie_corrigee"] = "OUI"
        return resultat

    resultat["geographie_corrigee"] = "NON"
    return resultat


def normaliser_date_ia(date_ia):
    """Conserve uniquement une date événement explicite au format ISO.

    Une simple année comme 2026 n'est pas une date d'événement et est donc
    volontairement laissée vide.
    """
    valeur = str(date_ia or "").strip()
    if re.fullmatch(r"20(?:25|26|27)-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12]\d|3[01])", valeur):
        return valeur
    return ""


def classer_statut_final(resultat, decision):
    """Classement sécurisé : l'IA ne détruit jamais un statut Python fiable."""
    statut_python = str(resultat.get("statut", "À vérifier"))

    if decision == "PERTINENT":
        resultat["statut_final"] = "Pertinent"
    elif decision == "NON_PERTINENT":
        resultat["statut_final"] = "Non pertinent"
    else:
        # IA indisponible, quota, conflit ou information insuffisante :
        # on conserve la décision Python déjà obtenue.
        resultat["statut_final"] = statut_python

    return resultat

# ============================================================
# COUCHE IA GEMINI
# ============================================================

def appliquer_ia(pertinents, a_verifier):
    """Analyse les candidats avec Gemini sans perdre les décisions Python."""

    candidats = pertinents + a_verifier
    candidats.sort(key=lambda x: x.get("score", 0), reverse=True)
    candidats = enrichir_candidats_ia(candidats)

    analyses = 0
    ia_resultats = []
    nouveaux_pertinents = []
    nouveaux_a_verifier = []
    quota_bloque = False

    print()
    print("========================================")
    print("COUCHE IA GEMINI")
    print("========================================")
    print(f"Candidats à analyser : {len(candidats)}")
    print(f"Maximum d'analyses IA : {MAX_IA_ANALYSES}")

    for resultat in candidats:
        statut_python = str(resultat.get("statut", "À vérifier"))

        # ----------------------------------------------------
        # Plus d'appel après la limite de 50
        # ----------------------------------------------------
        if analyses >= MAX_IA_ANALYSES:
            resultat["decision_ia"] = "IA_NON_ANALYSE"
            resultat["confiance_ia"] = 0
            resultat["raison_ia"] = "Limite maximale d'analyses IA atteinte."
            resultat["categorie_ia"] = ""
            resultat["date_ia"] = ""
            resultat["departement_ia"] = ""
            resultat["ia_analyse"] = "NON"
            resultat["statut_final"] = statut_python

            ia_resultats.append(resultat)
            if statut_python == "Pertinent":
                nouveaux_pertinents.append(resultat)
            else:
                nouveaux_a_verifier.append(resultat)
            continue

        # ----------------------------------------------------
        # Quota Gemini déjà atteint : zéro nouvel appel
        # ----------------------------------------------------
        if quota_bloque:
            resultat["decision_ia"] = "IA_INDISPONIBLE"
            resultat["confiance_ia"] = 0
            resultat["raison_ia"] = (
                "Quota Gemini atteint ou indisponible : "
                "aucun nouvel appel IA effectué."
            )
            resultat["categorie_ia"] = ""
            resultat["date_ia"] = ""
            resultat["departement_ia"] = ""
            resultat["ia_analyse"] = "NON"
            resultat["statut_final"] = statut_python

            ia_resultats.append(resultat)
            if statut_python == "Pertinent":
                nouveaux_pertinents.append(resultat)
            else:
                nouveaux_a_verifier.append(resultat)
            continue

        # ----------------------------------------------------
        # Appel Gemini
        # ----------------------------------------------------
        analyse = analyser_resultat(resultat)
        analyses += 1
        resultat.update(analyse)
        resultat["ia_analyse"] = "OUI"

        decision = str(analyse.get("decision_ia", "IA_INDISPONIBLE"))
        raison = str(analyse.get("raison_ia", "")).lower()

        # Normalisation stricte de la date IA : une année seule n'est
        # jamais enregistrée comme date d'événement.
        resultat["date_ia"] = normaliser_date_ia(resultat.get("date_ia", ""))

        # ----------------------------------------------------
        # Protection département
        # ----------------------------------------------------
        dept_python = str(resultat.get("departement", "")).strip()
        dept_ia = str(resultat.get("departement_ia", "")).strip()

        if dept_python and dept_ia and dept_ia != dept_python:
            resultat["raison_ia"] = (
                str(resultat.get("raison_ia", ""))
                + " | Conflit de département avec la détection Python : vérification requise."
            ).strip(" |")
            decision = "A_VERIFIER"
            resultat["decision_ia"] = decision

        if dept_ia and dept_ia not in {"16", "17", "19", "23", "79", "86", "87"}:
            decision = "A_VERIFIER"
            resultat["decision_ia"] = decision

        # ----------------------------------------------------
        # Détection immédiate des quotas
        # ----------------------------------------------------
        quota_detecte = (
            decision == "IA_INDISPONIBLE"
            and any(
                terme in raison
                for terme in (
                    "429",
                    "resource_exhausted",
                    "resource exhausted",
                    "quota",
                    "rate limit",
                    "too many requests",
                )
            )
        )

        if quota_detecte:
            quota_bloque = True
            # L'appel a bien été tenté, donc OUI, mais le résultat IA
            # ne remplace jamais le statut Python.
            resultat["statut_final"] = statut_python

            ia_resultats.append(resultat)
            if statut_python == "Pertinent":
                nouveaux_pertinents.append(resultat)
            else:
                nouveaux_a_verifier.append(resultat)
            continue

        # ----------------------------------------------------
        # Décision IA normale
        # ----------------------------------------------------
        classer_statut_final(resultat, decision)

        if decision == "PERTINENT":
            nouveaux_pertinents.append(resultat)
        elif decision == "NON_PERTINENT":
            # Traité dans main() pour conserver le groupe non pertinent.
            pass
        else:
            nouveaux_a_verifier.append(resultat)

        ia_resultats.append(resultat)

    nouveaux_pertinents.sort(key=lambda x: x.get("confiance_ia", 0), reverse=True)
    nouveaux_a_verifier.sort(key=lambda x: x.get("score", 0), reverse=True)

    print(f"Analyses IA effectuées : {analyses}")
    print(f"Blocage quota Gemini : {'OUI' if quota_bloque else 'NON'}")
    print(
        "Décisions IA : "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'PERTINENT')} pertinent / "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'A_VERIFIER')} à vérifier / "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'NON_PERTINENT')} non pertinent / "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'IA_INDISPONIBLE')} indisponible / "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'IA_NON_ANALYSE')} non analysé"
    )

    return nouveaux_pertinents, nouveaux_a_verifier, ia_resultats


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
        "url_finale",
        "titre_page",
        "description_page",
        "date_evenement_page",
        "lieu_page",
        "organisateur_page",
        "page_enrichie",
        "erreur_enrichissement",
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
        "decision_ia",
        "confiance_ia",
        "raison_ia",
        "categorie_ia",
        "date_ia",
        "departement_ia",
        "geographie_corrigee",
        "ia_analyse",
        "statut_final",
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

    # Fichier dédié à l'audit IA.
    candidats_ia = [
        r for r in tous
        if r.get("ia_analyse") == "OUI"
        or r.get("decision_ia") == "IA_INDISPONIBLE"
        or r.get("decision_ia") == "IA_NON_ANALYSE"
    ]

    def cle_tri_ia(resultat):
        decision = resultat.get("decision_ia", "")

        if decision == "PERTINENT":
            return (
                0,
                -int(resultat.get("confiance_ia", 0) or 0)
            )

        if decision == "A_VERIFIER":
            return (
                1,
                -int(resultat.get("score", 0) or 0)
            )

        if decision == "IA_INDISPONIBLE":
            return (2, 0)

        return (3, 0)

    candidats_ia.sort(key=cle_tri_ia)

    dataframe(candidats_ia).to_csv(
        DATA / "resultats_ia.csv",
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

    # --------------------------------------------------------
    # COUCHE IA
    # --------------------------------------------------------
    pertinents_ia, a_verifier_ia, ia_resultats = appliquer_ia(
        pertinents,
        a_verifier
    )

    # Les candidats que l'IA déclare NON_PERTINENT restent dans
    # non_pertinents. En cas d'indisponibilité de l'IA, on conserve
    # la décision Python précédente pour ne jamais perdre de résultat.
    non_pertinents_ia = []
    for resultat in ia_resultats:
        decision = resultat.get("decision_ia", "")
        if decision == "NON_PERTINENT":
            non_pertinents_ia.append(resultat)

    # Éviter les doublons : les candidats IA remplacent leurs
    # catégories Python initiales.
    candidats_urls = {x.get("url") for x in (pertinents + a_verifier)}
    non_pertinents_python = [
        x for x in non_pertinents
        if x.get("url") not in candidats_urls
    ]

    pertinents = pertinents_ia
    a_verifier = a_verifier_ia
    non_pertinents = non_pertinents_python + non_pertinents_ia

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
