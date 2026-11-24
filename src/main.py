from pathlib import Path
from urllib.parse import quote_plus
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from email.utils import parsedate_to_datetime
import hashlib
import threading
import time
import re
import json
import difflib
import subprocess
import sys
import os

import feedparser
import pandas as pd
import yaml
import requests
from requests.adapters import HTTPAdapter
from bs4 import BeautifulSoup

from ia import analyser_resultat
from sources_externes import collecter_sources_externes
from teams import construire_resume_teams, envoyer_teams, sauvegarder_payload


def synchroniser_excel():
    """
    Lance le synchroniseur Excel après la génération de resultats_ia.csv.
    Le synchroniseur est exécuté avec le même interpréteur Python que main.py.
    """
    script = ROOT / "src" / "synchroniser_excel.py"

    print()
    print("========================================")
    print("SYNCHRONISATION DU TABLEUR EXCEL")
    print("========================================")

    if not script.exists():
        print(f"⚠️ Synchroniseur introuvable : {script}")
        return False

    try:
        resultat = subprocess.run(
            [sys.executable, str(script)],
            cwd=str(ROOT),
            capture_output=False,
            text=True,
            check=False,
        )

        if resultat.returncode == 0:
            print("Synchronisation Excel terminée avec succès.")
            return True

        print(
            "⚠️ La synchronisation Excel a échoué "
            f"(code retour {resultat.returncode})."
        )
        return False

    except Exception as e:
        print(f"⚠️ Erreur lors du lancement du synchroniseur : {e}")
        return False


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
MAX_IA_ANALYSES = int(os.getenv("MAX_GEMINI_ANALYSES", "50"))
COOLDOWN_IA_SECONDS = int(os.getenv("GEMINI_COOLDOWN_SECONDS", "15"))
GEMINI_SCORE_MIN = int(os.getenv("SEUIL_GEMINI", "24"))
MAX_ERREURS_IA_CONSECUTIVES = 3  # disjoncteur : on stoppe après 3 échecs d'affilée

# Gestion des années : 2027 est la priorité, 2025 et avant sont écartés.
ANNEE_MIN = 2026          # toute information antérieure est retirée
ANNEE_PRIORITAIRE = 2027  # recherchée plus largement et mise en avant
BONUS_2027 = 12
SEUIL_VERIF_2027 = 6      # seuil « À vérifier » assoupli pour 2027 (10 sinon)
MALUS_HORS_PERIODE = 40

# V9 : filtres / priorisation
# Gemini intervient uniquement après le premier filtre Python.
# Les candidats atteignant ce score sont les plus pertinents et sont classés
# par score décroissant avant limitation à MAX_IA_ANALYSES.
SEUIL_IA_MIN = GEMINI_SCORE_MIN
SEUIL_IA_MAX = 999
SEUIL_AUTO_PERTINENT = 32        # score local suffisamment fort
SEUIL_SOURCE_OFFICIELLE = 1
SIMILARITE_TITRE_DEDOUBLONNAGE = 0.94
# Un article publié avant cette date, sans année 2026+ dans le texte, est
# considéré comme portant sur une édition 2025 ou antérieure. Les articles de
# décembre 2025 restent candidats : ils peuvent annoncer un événement 2026.
DATE_PUBLICATION_MIN = datetime(2025, 12, 1)

thread_local = threading.local()


# ============================================================
# CHARGEMENT CONFIGURATION
# ============================================================

def court(texte, n=70):
    """Texte tronqué sur une ligne, pour des logs lisibles."""
    t = " ".join(str(texte or "").split())
    return t if len(t) <= n else t[: n - 1] + "…"


# ============================================================
# VERSION ET MÉMOIRE DES DÉCISIONS IA
# ============================================================

VERSION_PROGRAMME = "2026-11-24-v10-delta-teams"

# À changer quand le prompt de ia.py ou les règles de décision évoluent :
# toutes les décisions mémorisées avec une autre version sont ignorées.
REGLES_VERSION = "2026-11-24-v9"

# Durée (jours) pendant laquelle une décision mémorisée est réutilisée.
MEMOIRE_VALIDITE_JOURS = {
    "PERTINENT": 45,
    "NON_PERTINENT": 45,
    "A_VERIFIER": 10,   # la page peut être mise à jour : on revérifie plus tôt
}
# Décision prise sans avoir pu lire la page : on la retente très vite.
MEMOIRE_VALIDITE_SANS_PAGE_JOURS = 3
MEMOIRE_CONSERVATION_JOURS = 120

MEMOIRE_COLONNES = [
    "cle", "url", "titre", "decision_ia", "confiance_ia", "raison_ia",
    "categorie_ia", "date_ia", "departement_ia", "page_lue",
    "date_analyse", "version_regles",
]


def chemin_memoire():
    return DATA / "memoire_ia.csv"


def cle_memoire(resultat):
    base = str(resultat.get("url") or resultat.get("titre") or "").strip().lower()
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:16]


def charger_memoire():
    """Charge la mémoire des décisions IA. Toute erreur donne une mémoire vide."""
    chemin = chemin_memoire()
    if not chemin.exists():
        return {}
    try:
        df = pd.read_csv(chemin, dtype=str, keep_default_na=False)
        return {
            ligne["cle"]: ligne
            for ligne in df.to_dict("records")
            if ligne.get("cle")
        }
    except Exception as e:
        print(f"⚠️ Mémoire IA illisible ({type(e).__name__}) : repart à vide.")
        return {}


def memoire_valide(entree):
    """Vrai si la décision mémorisée peut être réutilisée sans appeler l'IA."""
    if not entree:
        return False
    if entree.get("version_regles") != REGLES_VERSION:
        return False

    decision = entree.get("decision_ia", "")
    if decision not in MEMOIRE_VALIDITE_JOURS:
        return False

    try:
        age = (datetime.now() - datetime.fromisoformat(entree["date_analyse"])).days
    except Exception:
        return False

    validite = (
        MEMOIRE_VALIDITE_JOURS[decision]
        if entree.get("page_lue") == "OUI"
        else MEMOIRE_VALIDITE_SANS_PAGE_JOURS
    )
    return age < validite


def sauvegarder_memoire(memoire):
    """Écrit la mémoire en supprimant les entrées trop anciennes."""
    maintenant = datetime.now()
    gardees = []

    for entree in memoire.values():
        try:
            age = (maintenant - datetime.fromisoformat(entree["date_analyse"])).days
        except Exception:
            continue
        if age <= MEMOIRE_CONSERVATION_JOURS:
            gardees.append({c: entree.get(c, "") for c in MEMOIRE_COLONNES})

    try:
        DATA.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(gardees, columns=MEMOIRE_COLONNES).to_csv(
            chemin_memoire(), index=False, encoding="utf-8-sig"
        )
        print(f"Mémoire IA sauvegardée : {len(gardees)} décisions.")
    except Exception as e:
        print(f"⚠️ Mémoire IA non sauvegardée ({type(e).__name__}) : {e}")


def load_yaml(filename):
    with open(CONFIG / filename, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_yaml_optionnel(filename):
    """Fichier de configuration facultatif : {} s'il est absent ou illisible."""
    chemin = CONFIG / filename
    if not chemin.exists():
        return {}
    try:
        with open(chemin, "r", encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    except Exception as e:
        print(f"⚠️ {filename} illisible ({type(e).__name__}) : ignoré.")
        return {}


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


# Requêtes supplémentaires réservées à 2027 : formats indirects, annonces
# de prochaines éditions et recherche par grandes villes.
# (mot-clé entre guillemets, complément libre). Le complément n'est pas mis
# entre guillemets : il oriente la recherche vers les annonces d'éditions
# futures sans exiger une expression exacte. Les termes génériques seuls
# (« édition 2027 », « save the date ») ramenaient surtout des festivals et
# de la politique : ils sont toujours associés à un contexte emploi/études.
REQUETES_2027_EMPLOI = [
    ("forum entreprises", ""),
    ("rencontres entreprises", ""),
    ("career day", ""),
    ("journée carrières", ""),
    ("recrutement campus", ""),
    ("forum alternance", ""),
    ("semaine de l'emploi", ""),
    ("forum emploi", "prochaine édition"),
    ("salon emploi", "prochaine édition"),
    ("job dating", "prochaine édition"),
    ("forum entreprises", "prochaine édition"),
]

REQUETES_2027_SUPERIEUR = [
    ("salons et forums", ""),
    ("forum alumni", ""),
    ("salon étudiant", "prochaine édition"),
    ("portes ouvertes", "université école"),
]

# Mots-clés combinés aux deux principales villes de chaque département.
REQUETES_2027_VILLES = [
    "forum emploi",
    "forum entreprises",
    "salon étudiant",
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
    "dirigeant",
    "candidats",
    "métiers",
    "carrière",
    "alternance",
    "apprentissage",
    "cadre",
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
    """Années actuelles ou futures (2026 à 2028).

    - « 26/27 », « 2026-2027 », « 2026/27 » comptent pour les deux années ;
    - une période commençant avant 2025 (« 2023-2027 » : plan, contrat,
      programme pluriannuel) n'est pas une édition 2027 et est ignorée.
    """
    t = normaliser(texte)
    t = re.sub(r"\b20(?:1\d|2[0-4])\s*[-–/]\s*(?:20)?2[6-9]\b", " ", t)
    annees = set(re.findall(r"\b20(?:26|27|28)\b", t))
    for debut, fin in re.findall(
        r"\b(?:20)?(2[6-8])\s*[/\-–]\s*(?:20)?(2[6-8])\b", t
    ):
        annees.add("20" + debut)
        annees.add("20" + fin)
    return sorted(annees)


def extraire_annees_passees(texte):
    """Années 2010 à 2025 : informations considérées comme périmées."""
    return sorted(set(re.findall(r"\b20(?:1\d|2[0-5])\b", normaliser(texte))))


def date_publication_dt(date_publication):
    """Date de publication RSS en datetime naïf, ou None si illisible."""
    texte = str(date_publication or "").strip()
    if not texte:
        return None
    try:
        return parsedate_to_datetime(texte).replace(tzinfo=None)
    except Exception:
        trouve = re.search(r"\b(20\d{2})\b", texte)
        return datetime(int(trouve.group(1)), 1, 1) if trouve else None


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

    # Source dont le département est connu (université, ville, OpenAgenda) :
    # la zone est démontrée par la source elle-même.
    positif = (
        departement_present
        or bool(villes_presentes)
        or resultat.get("geographie_source") == "OUI"
    )

    return positif, departement_present, villes_presentes, hors_cible


# ============================================================
# CONSTRUCTION DES REQUÊTES
# ============================================================

# ============================================================
# OPERATEURS WEB V2 — recherche élargie mais maîtrisée
# ============================================================
# Les requêtes historiques restent inchangées. Cette seconde passe ajoute
# des combinaisons documentées de recherche : expressions exactes, exclusions
# et domaines institutionnels. Elles servent à découvrir des événements que
# les requêtes génériques ratent, sans transformer le résultat en simple bruit.
OPERATEURS_WEB_V2_DEFAUT = [
    '"forum entreprises" université', '"forum entreprises" master',
    '"forum entreprises" ingénieurs', '"forum entreprises" diplômés',
    '"rencontre entreprises" université', '"rencontre entreprises" master',
    '"rencontres entreprises" ingénieurs', '"job dating" cadres',
    '"job dating" "jeunes diplômés"', '"job dating" "bac+5"',
    '"forum emploi" université', '"forum emploi" master',
    '"forum emploi" "grande école"', '"forum emploi" ingénieurs',
    'recrutement "jeunes diplômés"', 'recrutement cadres université',
    'recrutement ingénieurs université', 'recrutement master université',
    '"career day" université', '"career forum" université',
    'alumni entreprises recrutement', 'alumni carrière entreprises',
    'doctorat entreprises recrutement', 'doctorants entreprises',
    'ingénieurs recrutement', '"premier emploi" diplômés',
    '"objectif premier emploi"', '"relations entreprises" université',
    '"rencontre employeurs" université', '"forum alternance" master',
]

DOMAINES_INSTITUTIONNELS_V2_DEFAUT = [
    "site:apec.fr", "site:francetravail.fr", "site:messervices.francetravail.fr",
    "site:univ-poitiers.fr", "site:univ-larochelle.fr", "site:unilim.fr",
    "site:data.gouv.fr",
]

EXCLUSIONS_WEB_V2_DEFAUT = '-"job étudiant" -"emploi étudiant" -"job saisonnier" -"emploi saisonnier" -"job d’été" -"jobs d’été" -lycée -collège -"post-bac"'


def charger_operateurs_web_v2():
    cfg = load_yaml_optionnel("operators.yml")
    web = cfg.get("web", {}) if isinstance(cfg, dict) else {}
    combinaisons = cfg.get("combinaisons", []) if isinstance(cfg, dict) else []
    domaines = cfg.get("institutional_domains", DOMAINES_INSTITUTIONNELS_V2_DEFAUT) if isinstance(cfg, dict) else DOMAINES_INSTITUTIONNELS_V2_DEFAUT
    exclusions = cfg.get("exclusions", EXCLUSIONS_WEB_V2_DEFAUT) if isinstance(cfg, dict) else EXCLUSIONS_WEB_V2_DEFAUT
    ops = []
    if combinaisons:
        for paire in combinaisons:
            if isinstance(paire, (list, tuple)) and len(paire) == 2:
                ops.append(f'"{paire[0]}" "{paire[1]}"')
    if not ops:
        ops = list(OPERATEURS_WEB_V2_DEFAUT)
    return ops, list(domaines), exclusions


def construire_requetes(departements):
    requetes = []
    operateurs_web, domaines_institutionnels, exclusions_web = charger_operateurs_web_v2()

    # Les années antérieures à ANNEE_MIN ne sont plus recherchées.
    annees = sorted(
        {int(a) for a in departements["annees"] if int(a) >= ANNEE_MIN}
        | {ANNEE_PRIORITAIRE}
    )

    for code, nom in departements["departements"].items():
        for annee in annees:

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

            if annee != ANNEE_PRIORITAIRE:
                continue

            # Recherche élargie pour 2027.
            for mot, complement in REQUETES_2027_EMPLOI:
                requetes.append({
                    "departement": code,
                    "nom_departement": nom,
                    "annee": annee,
                    "categorie_numero": 2,
                    "categorie": "Emploi / recrutement / formation",
                    "mot_cle": mot,
                    "complement": complement,
                })

            for mot, complement in REQUETES_2027_SUPERIEUR:
                requetes.append({
                    "departement": code,
                    "nom_departement": nom,
                    "annee": annee,
                    "categorie_numero": 1,
                    "categorie": "Salon étudiant / enseignement supérieur",
                    "mot_cle": mot,
                    "complement": complement,
                })

            for ville in VILLES_DEPARTEMENTS.get(code, [])[:2]:
                for mot in REQUETES_2027_VILLES:
                    requetes.append({
                        "departement": code,
                        "nom_departement": nom,
                        "annee": annee,
                        "categorie_numero": 1 if "étudiant" in mot else 2,
                        "categorie": (
                            "Salon étudiant / enseignement supérieur"
                            if "étudiant" in mot
                            else "Emploi / recrutement / formation"
                        ),
                        "mot_cle": mot,
                        "lieu": ville,
                    })

    # Seconde passe opérateurs Web : quelques combinaisons à fort rendement.
    # On garde une requête par département/année/opérateur et on mélange
    # ensuite les sources institutionnelles dans un sous-ensemble limité.
    for code, nom in departements["departements"].items():
        for annee in annees:
            for operateur in operateurs_web:
                requetes.append({
                    "departement": code,
                    "nom_departement": nom,
                    "annee": annee,
                    "categorie_numero": 2,
                    "categorie": "Recherche Web V2 / opérateurs",
                    "mot_cle": operateur,
                    "complement": exclusions_web,
                    "operateur_web": "V2",
                    "famille_operateur": "combinaison",
                })

            # Passe institutionnelle très ciblée : 4 domaines prioritaires.
            for domaine in domaines_institutionnels:
                requetes.append({
                    "departement": code,
                    "nom_departement": nom,
                    "annee": annee,
                    "categorie_numero": 2,
                    "categorie": "Source institutionnelle / opérateur",
                    "mot_cle": f'"forum entreprises" {domaine}',
                    "complement": str(annee),
                    "operateur_web": "INSTITUTIONNEL",
                    "famille_operateur": "source_institutionnelle",
                })

    return requetes


# ============================================================
# GOOGLE NEWS RSS
# ============================================================

def rechercher_google_news(requete):

    texte = (
        f'"{requete["mot_cle"]}" '
        f'"{requete.get("lieu") or requete["nom_departement"]}" '
        + (f'{requete["complement"]} ' if requete.get("complement") else "")
        + f'{requete["annee"]}'
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
                    "sources_trouvees": source,
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


def est_lien_google_news(url):
    return "news.google.com" in str(url or "").lower()


def decoder_liens_google(candidats):
    """
    Remplace les liens Google News (redirections chiffrées) par l'URL réelle
    de l'article, afin que l'IA lise la vraie page et non la page d'accueil
    de Google News. Un seul appel par lot ; tout échec est non bloquant.
    """
    cibles = [
        c for c in candidats
        if est_lien_google_news(c.get("url")) and not c.get("url_reelle")
    ]
    if not cibles:
        return 0

    try:
        from googlenewsdecoder import gnewsdecoder
    except Exception as e:
        print(f"⚠️ googlenewsdecoder indisponible ({type(e).__name__}) : "
              "liens Google News non décodés. Ajoutez-le à requirements.txt.")
        return 0

    decodes = 0
    TAILLE_LOT = 20

    for debut in range(0, len(cibles), TAILLE_LOT):
        lot = cibles[debut:debut + TAILLE_LOT]
        try:
            reponses = gnewsdecoder([c["url"] for c in lot], timeout=15)
        except Exception as e:
            print(f"⚠️ Décodage Google News en échec sur un lot : "
                  f"{type(e).__name__} : {str(e)[:120]}")
            continue

        for candidat, rep in zip(lot, reponses):
            if isinstance(rep, dict) and rep.get("success") and rep.get("decoded_url"):
                candidat["url_reelle"] = rep["decoded_url"]
                decodes += 1

        time.sleep(1)

    print(f"Liens Google News décodés : {decodes}/{len(cibles)}", flush=True)
    return decodes


def extraire_texte_page(soup, limite=2500):
    """Texte lisible de l'article (balises article/main, sinon paragraphes)."""
    for balise in soup(["script", "style", "nav", "footer", "header",
                        "aside", "form", "noscript", "iframe"]):
        balise.decompose()

    zone = soup.find("article") or soup.find("main") or soup.body or soup
    paragraphes = [
        p.get_text(" ", strip=True)
        for p in zone.find_all(["h1", "h2", "h3", "p", "li"])
    ]
    texte = " ".join(p for p in paragraphes if len(p) > 25)
    return " ".join(texte.split())[:limite]


def enrichir_page(resultat):
    """
    Récupère des informations complémentaires sur la page source.
    L'enrichissement est limité aux candidats IA afin de ne pas
    ralentir la collecte RSS.
    """
    url = resultat.get("url_reelle") or resultat.get("url", "")
    if not url:
        return resultat

    # Lien Google News non décodé : la page récupérée serait celle de Google,
    # sans rapport avec l'article. On n'enrichit pas (et on ne pollue pas
    # la description envoyée à l'IA).
    if est_lien_google_news(url):
        resultat["page_enrichie"] = "NON"
        resultat["erreur_enrichissement"] = "lien Google News non décodé"
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
        # Extrait calculé sur une copie : l'extraction retire des balises.
        extrait_page = extraire_texte_page(
            BeautifulSoup(response.content, "lxml")
        )
        description_page = extraire_meta(
            soup, ["description", "og:description", "twitter:description"]
        )
        date_evenement = extraire_meta(
            soup,
            ["event:start_time", "event:start_date", "startDate"]
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
                            element.get("startDate") or ""
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
        resultat["extrait_page"] = extrait_page
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
    decoder_liens_google(cible)

    with ThreadPoolExecutor(max_workers=min(8, MAX_WORKERS)) as executor:
        futures = [executor.submit(enrichir_page, resultat) for resultat in cible]
        total_cible = len(cible)
        faits = 0
        echecs = 0

        for future in as_completed(futures):
            faits += 1
            try:
                future.result()
            except Exception:
                echecs += 1

            if faits % 10 == 0 or faits == total_cible:
                print(
                    f"Enrichissement : {faits}/{total_cible} "
                    f"| erreurs : {echecs}",
                    flush=True,
                )

    duree = time.time() - debut
    enrichis = sum(1 for r in cible if r.get("page_enrichie") == "OUI")

    print(f"Pages enrichies : {enrichis}/{len(cible)}")
    print(f"Temps enrichissement : {duree:.1f} secondes")

    return candidats


# ============================================================
# DÉDOUBLONNAGE
# ============================================================

def canoniser_url(url):
    """Retire les paramètres de tracking pour fiabiliser le dédoublonnage."""
    valeur = str(url or "").strip()
    if not valeur:
        return ""
    valeur = re.sub(r"[?#].*$", "", valeur)
    valeur = valeur.rstrip("/")
    return valeur.lower()


def cle_titre_evenement(titre):
    """Titre simplifié pour comparer des événements issus de plusieurs sources."""
    t = normaliser(titre)
    t = re.sub(r"[^a-z0-9àâäéèêëîïôöùûüç ]", " ", t)
    mots = [
        "le", "la", "les", "un", "une", "des", "de", "du", "et", "en",
        "à", "a", "pour", "sur", "édition", "edition", "2026", "2027", "2028",
    ]
    return " ".join(m for m in t.split() if m not in mots)


def fusionner_resultats(principal, doublon):
    """Conserve la source primaire et complète les champs manquants."""
    for cle, valeur in doublon.items():
        if not principal.get(cle) and valeur:
            principal[cle] = valeur
    sources = [x for x in [
        principal.get("source", ""), doublon.get("source", "")
    ] if x]
    principal["sources_trouvees"] = " | ".join(dict.fromkeys(sources))
    return principal


def dedoublonner(resultats):
    """
    V9.5 : dédoublonnage par URL canonique + rapprochement de titres.
    Les résultats externes arrivent en premier dans main(), donc la source
    primaire reste prioritaire lorsqu'elle existe.
    """
    par_url = {}
    uniques = []
    titres = []

    for resultat in resultats:
        url = canoniser_url(resultat.get("url_finale") or resultat.get("url"))
        titre = cle_titre_evenement(resultat.get("titre_page") or resultat.get("titre"))
        if not url:
            continue

        if url in par_url:
            fusionner_resultats(par_url[url], resultat)
            continue

        doublon_index = None
        code = str(resultat.get("departement", ""))
        annee = str(resultat.get("annee_recherchee", ""))

        # Comparaison prudente : même département + même année + titre très proche.
        if titre:
            for idx, (ancien_titre, ancien_code, ancienne_annee) in enumerate(titres):
                if code != ancien_code or annee != ancienne_annee:
                    continue
                if len(titre) < 18 or len(ancien_titre) < 18:
                    continue
                ratio = difflib.SequenceMatcher(None, titre, ancien_titre).ratio()
                if ratio >= SIMILARITE_TITRE_DEDOUBLONNAGE:
                    doublon_index = idx
                    break

        if doublon_index is not None:
            principal = uniques[doublon_index]
            fusionner_resultats(principal, resultat)
            par_url[url] = principal
        else:
            par_url[url] = resultat
            uniques.append(resultat)
            titres.append((titre, code, annee))

    return uniques


def classer_fiabilite_source(resultat):
    """V9.6 : attribue un niveau de fiabilité à partir de l'URL et du nom source."""
    texte = normaliser(
        " ".join(str(resultat.get(k, "")) for k in
                 ("url_finale", "url_reelle", "url", "source", "organisateur_page"))
    )
    if any(domaine in texte for domaine in DOMAINES_SOURCE_OFFICIELLE):
        return 1
    noms_officiels = [
        "apec", "france travail", "université de poitiers",
        "la rochelle université", "université de limoges",
    ]
    if any(nom in texte for nom in noms_officiels):
        return 1
    if any(domaine in texte for domaine in DOMAINES_SOURCE_INSTITUTIONNELLE):
        return 2
    if resultat.get("niveau_source"):
        try:
            return int(resultat.get("niveau_source"))
        except Exception:
            pass
    return 3 if resultat.get("source") else 4


# ============================================================
# SCORING RENFORCÉ APRÈS AUDIT
# ============================================================

def _signaux(texte, patterns):
    texte = normaliser(texte)
    return [p for p in patterns if p in texte]


# ============================================================
# SCORING CIBLE APEC — LES OPERATEURS DE RECHERCHE RESTENT INCHANGES
# ============================================================

APEC_DIRECT = [
    "apec", "apec connect", "cadre", "cadres", "cadre dirigeant",
    "jeune diplômé", "jeunes diplômés", "jeune diplômée", "jeunes diplômées",
    "diplômés", "diplômées", "recrutement de cadres", "recrutement cadre",
    "mobilité professionnelle", "évolution professionnelle", "reconversion",
]

APEC_QUALIFIE = [
    "bac+3", "bac+4", "bac+5", "master", "mastère", "mba", "doctorant",
    "doctorants", "docteur", "doctorat", "ingénieur", "ingénieurs",
    "école d'ingénieurs", "école d’ingénieurs", "école de commerce",
    "grande école", "grandes écoles", "iae", "université", "universitaire",
    "alumni", "anciens élèves", "jeunes talents", "talents", "graduate",
]

RELATION_ENTREPRISE = [
    "rencontre entreprises", "rencontres entreprises", "entreprises et étudiants",
    "entreprises / étudiants", "entreprises et diplômés", "employeurs",
    "recruteurs", "recruteur", "job dating", "forum entreprises",
    "forum entreprise", "forum emploi", "salon emploi", "salon de l'emploi",
    "recrutement", "recruter", "offres d'emploi", "relations entreprises",
    "relation entreprise", "career", "carrière",
]

SUPERIEUR_FORT = [
    "université", "université de", "la rochelle université", "poitiers université",
    "université de limoges", "iaE", "master", "bac+5", "école d'ingénieurs",
    "école d’ingénieurs", "école de commerce", "grande école", "alumni",
    "anciens élèves", "doctorat", "doctorants", "ingénieurs",
]

CONTEXTE_APEC = [
    "université", "universitaire", "master", "mastère", "bac+3", "bac+4",
    "bac+5", "doctorat", "doctorant", "docteur", "ingénieur", "ingénieurs",
    "école d'ingénieurs", "école d’ingénieurs", "école de commerce",
    "grande école", "grandes écoles", "iae", "alumni", "anciens élèves",
    "forum entreprises", "forum entreprise", "rencontre entreprises",
    "rencontres entreprises", "entreprises et étudiants",
    "entreprises et diplômés", "relations entreprises", "recruteurs",
    "employeurs", "career", "carrière", "insertion professionnelle",
    "mobilité professionnelle", "évolution professionnelle", "reconversion",
]

EXCLUSIONS_FORTES = [
    "job étudiant", "jobs étudiants", "emploi étudiant", "emplois étudiants",
    "job saisonnier", "jobs saisonniers", "emploi saisonnier", "emplois saisonniers",
    "travail saisonnier", "intérim", "missions d'intérim", "petits boulots",
    "job d'été", "jobs d'été", "vendanges", "baby-sitting", "restauration",
    "serveur", "serveuse", "plongeur", "plongeuse", "caissier", "caissière",
    "manutentionnaire", "préparateur de commande", "préparatrice de commande",
    "sans diplôme", "sans qualification", "niveau cap", "niveau bep", "collège",
    "lycée", "post-bac", "après le bac", "orientation des lycéens",
]

EXCLUSIONS_EDITORIAL = [
    "actualités", "actualité", "portrait", "interview", "tribune", "édito",
    "palmarès", "retour sur", "en images", "recensement", "vie locale",
]

# Formats événementiels indirects : ils ne s'appellent pas « salon emploi »
# mais restent de vrais événements Apec potentiels (réflexion Apec, point 10).
EVENTS_INDIRECTS = [
    "forum entreprises", "forum entreprise", "forum des entreprises",
    "rencontre entreprises", "rencontres entreprises",
    "rencontre avec les entreprises", "rencontres avec les entreprises",
    "salons et forums", "salons & forums", "salon et forum", "salons, forums",
    "career day", "career days", "career fair",
    "job fair", "job meeting", "job days",
    "journée carrières", "journée carrière", "journée des carrières",
    "journées carrières", "semaine de l'emploi", "semaine de l’emploi",
    "semaine carrière", "semaine carrières", "semaine entreprises",
    "forum des métiers", "rencontres métiers", "rencontre métiers",
    "journée professionnelle", "journée entreprises", "journée des entreprises",
    "speed recruiting", "speed dating emploi", "recrutement campus",
    "campus recruitment", "recruitment day", "graduate programme",
    "forum alumni", "networking", "forum de l'alternance",
    "forum de l’alternance", "forum de l'emploi", "forum de l’emploi",
    "forum de recrutement", "forum du recrutement",
]

EVENT_SIGNAL_CIBLE = EVENTS_EMPLOI + EVENTS_SUPERIEUR + EVENTS_INDIRECTS

# V9.1 : signaux qui prouvent davantage qu'une page décrit un événement
# réel, plutôt qu'une simple actualité ou une page institutionnelle générale.
EVENT_PREUVE_FORTE = [
    "inscription", "inscrivez-vous", "inscrivez vous", "programme",
    "date et lieu", "date :", "lieu :", "entrée libre", "gratuit",
    "réservation", "billetterie", "save the date", "rendez-vous",
    "se tiendra", "aura lieu", "a lieu le", "le 15", "le 16", "le 17",
    "le 18", "le 19", "le 20", "le 21", "le 22", "le 23", "le 24",
    "le 25", "le 26", "le 27", "le 28", "le 29", "le 30", "le 31",
]

EVENT_MOTS_STRUCTURELS = [
    "forum", "salon", "job dating", "career day", "career fair",
    "rencontre", "journée", "atelier", "conférence", "table ronde",
    "portes ouvertes", "webinaire", "matinale", "semaine",
]

# V9.6 : classement des sources. Les domaines officiels sont prioritaires.
DOMAINES_SOURCE_OFFICIELLE = [
    "apec.fr", "francetravail.fr", "univ-poitiers.fr",
    "univ-larochelle.fr", "unilim.fr", "data.gouv.fr",
    "poitiers.fr", "larochelle.fr", "limoges.fr",
]

DOMAINES_SOURCE_INSTITUTIONNELLE = [
    "enseignementsup-recherche.gouv.fr", "education.gouv.fr",
    "nouvelle-aquitaine.fr", "grandpoitiers.fr", "agglo-larochelle.fr",
    "grandangouleme.fr", "grand-chatellerault.fr", "niortagglo.fr",
    "limoges-metropole.fr", "openagenda.com",
]



def parser_date_evenement(valeur):
    """Extrait une date ISO/FR simple depuis les métadonnées de page."""
    t = str(valeur or "").strip()
    m = re.search(r"\b(20\d{2})-(\d{2})-(\d{2})\b", t)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})[/-](202\d)\b", t)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def calculer_score(resultat):
    """V9 : scoring Apec centré sur événement réel + cible + période + source."""
    titre = normaliser(resultat.get("titre_page") or resultat.get("titre", ""))
    description = normaliser(resultat.get("description_page") or resultat.get("description", ""))
    source = normaliser(resultat.get("source", ""))
    page = normaliser(
        " ".join(str(resultat.get(k, "")) for k in (
            "titre_page", "description_page", "extrait_page",
            "lieu_page", "organisateur_page", "date_evenement_page"
        ))
    )
    texte = f"{titre} {description} {page}"
    categorie = int(resultat.get("categorie_numero", 0) or 0)
    annee = str(resultat.get("annee_recherchee", ""))

    direct = _signaux(texte, APEC_DIRECT)
    qualifie = _signaux(texte, APEC_QUALIFIE)
    entreprise = _signaux(texte, RELATION_ENTREPRISE)
    superieur = _signaux(texte, SUPERIEUR_FORT)
    exclusions = _signaux(texte, EXCLUSIONS_FORTES)
    editorial = _signaux(texte, EXCLUSIONS_EDITORIAL)
    event_emploi = _signaux(texte, EVENTS_EMPLOI)
    event_superieur = _signaux(texte, EVENTS_SUPERIEUR)
    event_indirect = _signaux(texte, EVENTS_INDIRECTS)
    preuve_event = _signaux(texte, EVENT_PREUVE_FORTE)
    structure_event = _signaux(texte, EVENT_MOTS_STRUCTURELS)

    geo_ok, dept_present, villes_presentes, hors_geo = detecter_geographie(resultat)
    date_evenement = parser_date_evenement(resultat.get("date_evenement_page", ""))
    annees = extraire_annees(texte)

    # V9.1 : un événement doit être identifiable, pas seulement suggéré par
    # « emploi », « master », « université », etc.
    event_dans_titre = bool(_signaux(titre, EVENT_SIGNAL_CIBLE))
    event_meta = bool(date_evenement or resultat.get("lieu_page") or resultat.get("organisateur_page"))
    evenement = bool(
        (event_emploi or event_superieur or event_indirect)
        and (event_dans_titre or preuve_event or event_meta or structure_event)
    )

    # Une page officielle contenant un intitulé événementiel reste acceptable
    # même si elle ne contient pas les mots « inscription » ou « programme ».
    niveau_source = classer_fiabilite_source(resultat)
    if not evenement and niveau_source in (1, 2) and event_dans_titre:
        evenement = True

    score = 0
    raisons = []

    if evenement:
        score += 10
        raisons.append("vrai événement identifié")
        if preuve_event:
            score += 4
            raisons.append("preuve événementielle")
        if date_evenement:
            score += 4
            raisons.append("date événement détectée")
        if resultat.get("lieu_page"):
            score += 2
            raisons.append("lieu détecté")
    else:
        score -= 18
        raisons.append("page sans événement suffisamment démontré")

    # Cible Apec.
    if direct:
        score += min(24, 8 * len(direct))
        raisons.append("cible Apec directe: " + ", ".join(direct[:4]))

    if qualifie:
        score += min(16, 4 * len(qualifie))
        raisons.append("profil qualifié: " + ", ".join(qualifie[:5]))

    if superieur:
        score += min(10, 3 * len(superieur))
        raisons.append("enseignement supérieur: " + ", ".join(superieur[:4]))

    if entreprise:
        score += min(16, 4 * len(entreprise))
        raisons.append("relation entreprises/recrutement: " + ", ".join(entreprise[:5]))

    # Relation particulièrement intéressante pour Apec.
    if superieur and entreprise:
        score += 8
        raisons.append("supérieur + entreprises")

    # Source.
    if niveau_source == 1:
        score += 10
        raisons.append("source officielle prioritaire")
    elif niveau_source == 2:
        score += 7
        raisons.append("source institutionnelle")
    elif niveau_source == 3:
        score += 2
        raisons.append("source secondaire")
    else:
        score -= 2
        raisons.append("source faible/non qualifiée")

    # Catégorie de recherche : bonus faible.
    if categorie == 1 and superieur:
        score += 3
    if categorie == 2 and (event_emploi or entreprise):
        score += 3

    # Géographie.
    if geo_ok:
        score += 5
        raisons.append("géographie cible vérifiée")
    else:
        score -= 12
        raisons.append("géographie cible non démontrée")
    if hors_geo:
        score -= 25
        raisons.append("hors géographie: " + ", ".join(hors_geo[:3]))

    # Exclusions fortes.
    if exclusions:
        malus = min(35, 12 * len(exclusions))
        score -= malus
        raisons.append("hors cible probable: " + ", ".join(exclusions[:5]))

    if editorial and not (direct or entreprise or superieur):
        score -= 20
        raisons.append("contenu éditorial/actualité")

    # Année recherchée.
    if annee and annee in annees:
        score += 3
        raisons.append(f"année {annee} détectée")

    est_2027 = str(ANNEE_PRIORITAIRE) in annees
    if est_2027:
        score += BONUS_2027
        raisons.append("2027 prioritaire")
    elif "2028" in annees:
        score += 6
        raisons.append("2028 future")
    elif "2026" in annees:
        score += 2

    # Date d'événement passée : beaucoup plus fiable qu'une simple année
    # trouvée dans le texte.
    evenement_passe = bool(date_evenement and date_evenement.date() < datetime.now().date())
    if evenement_passe:
        score -= MALUS_HORS_PERIODE
        raisons.append(f"événement passé : {date_evenement:%d/%m/%Y}")

    # Anciennes éditions dans une page sans année actuelle/future.
    annees_passees = extraire_annees_passees(texte)
    date_pub = date_publication_dt(resultat.get("date_publication", ""))
    hors_periode = (
        not annees
        and not date_evenement
        and (
            bool(annees_passees)
            or (date_pub is not None and date_pub < DATE_PUBLICATION_MIN)
        )
    )
    if hors_periode:
        score -= MALUS_HORS_PERIODE
        origine = ", ".join(annees_passees[:3]) if annees_passees else f"publication {date_pub:%d/%m/%Y}"
        raisons.append(f"hors période : {origine}")

    # Contexte Apec, sans laisser « université » seul décider.
    signaux_contexte = contient_terme(texte, CONTEXTE_APEC)
    if signaux_contexte:
        score += min(6, 2 * len(signaux_contexte))
        raisons.append("contexte qualifié: " + ", ".join(signaux_contexte[:4]))

    # V9.7 : priorité 2027.
    if est_2027 and evenement:
        priorite_2027 = "FORTE" if date_evenement or preuve_event else "MOYENNE"
    elif "2027" in annees:
        priorite_2027 = "MOYENNE"
    else:
        priorite_2027 = "FAIBLE"

    exclusion_forte = bool(exclusions)

    if hors_periode or evenement_passe:
        statut = "Non pertinent"
    elif not evenement or not geo_ok or hors_geo:
        # Filet de sécurité : un cas incomplet reste À vérifier si le score
        # est suffisamment haut, au lieu d'être silencieusement supprimé.
        statut = "À vérifier" if score >= 8 and not exclusion_forte else "Non pertinent"
    elif exclusion_forte and score < 32:
        statut = "Non pertinent"
    elif score >= SEUIL_AUTO_PERTINENT:
        statut = "Pertinent"
    elif score >= 8:
        statut = "À vérifier"
    else:
        statut = "Non pertinent"

    # Un format indirect sans source fiable ne devient jamais « Pertinent »
    # uniquement à cause du score.
    if (
        statut == "Pertinent"
        and event_indirect
        and not (event_emploi or event_superieur)
        and niveau_source not in (1, 2)
    ):
        statut = "À vérifier"
        raisons.append("format indirect : vérification nécessaire")

    raison = {
        "Pertinent": "Cible Apec + événement + période suffisamment démontrés",
        "À vérifier": "Signal utile mais preuve ou ciblage à confirmer",
        "Non pertinent": "Événement/cible/période hors périmètre",
    }[statut]

    return {
        "score": score,
        "score_cible_apec": score,
        "score_emploi": 10 if event_emploi else 0,
        "score_formation": 3 if _signaux(texte, TERMES_FORMATION) else 0,
        "score_superieur": min(10, 3 * len(superieur)),
        "score_apec": min(24, 8 * len(direct)),
        "score_entreprise": min(16, 4 * len(entreprise)),
        "score_qualifie": min(16, 4 * len(qualifie)),
        "score_geographie": 5 if geo_ok else -12,
        "score_date": 4 if date_evenement else 0,
        "score_bruit": -20 if editorial else 0,
        "score_complementaire": 0,
        "bonus_titre": 4 if event_dans_titre else 0,
        "signaux_cible_apec": ", ".join(direct),
        "signaux_superieur": ", ".join(superieur + qualifie[:5]),
        "signaux_entreprise": ", ".join(entreprise),
        "signaux_exclusion": ", ".join(exclusions),
        "statut": statut,
        "raison": raison,
        "motif": "Événement passé" if evenement_passe else "",
        "hors_periode": "OUI" if hors_periode or evenement_passe else "NON",
        "periode": (
            "2027 (prioritaire)" if est_2027
            else "2028" if "2028" in annees
            else "2026" if "2026" in annees
            else "Année non détectée"
        ),
        "details_score": " | ".join(raisons),
        "evenement_detecte": "OUI" if evenement else "NON",
        "preuve_evenement": "OUI" if preuve_event or event_dans_titre else "NON",
        "geographie_verifiee": "OUI" if geo_ok else "NON",
        "annee_detectee": ", ".join(annees),
        "hors_perimetre": ", ".join(exclusions),
        "hors_geographie": ", ".join(hors_geo),
        "niveau_source_v9": niveau_source,
        "priorite_2027": priorite_2027,
    }


def traiter(resultats):
    resultats = dedoublonner(resultats)

    pertinents = []
    a_verifier = []
    non_pertinents = []
    ecartes_hors_periode = 0

    for resultat in resultats:
        corriger_departement_par_ville(resultat)
        score_data = calculer_score(resultat)
        resultat.update(score_data)

        if resultat.get("hors_periode") == "OUI":
            ecartes_hors_periode += 1
            continue

        if resultat["statut"] == "Pertinent":
            pertinents.append(resultat)
        elif resultat["statut"] == "À vérifier":
            a_verifier.append(resultat)
        else:
            non_pertinents.append(resultat)

    pertinents.sort(key=lambda x: x["score"], reverse=True)
    a_verifier.sort(key=lambda x: x["score"], reverse=True)
    non_pertinents.sort(key=lambda x: x["score"], reverse=True)

    nb_2027 = sum(
        1 for x in pertinents + a_verifier
        if str(ANNEE_PRIORITAIRE) in str(x.get("annee_detectee", ""))
    )
    print()
    print(f"Résultats hors période (≤ {ANNEE_MIN - 1}) écartés : {ecartes_hors_periode}")
    print(f"Candidats mentionnant {ANNEE_PRIORITAIRE} : {nb_2027}")

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
    """Analyse Gemini uniquement les candidats ayant passé le premier filtre Python."""

    tous_candidats = pertinents + a_verifier

    # Premier test Python : Gemini ne voit que les candidats ayant atteint
    # le seuil défini par SEUIL_GEMINI. Les signaux indirects ne forcent plus
    # un appel IA : ils doivent d'abord passer ce premier filtre.
    candidats = []
    non_eligibles = []
    for candidat in tous_candidats:
        score = int(candidat.get("score", 0) or 0)
        if score >= GEMINI_SCORE_MIN:
            candidats.append(candidat)
        else:
            candidat["decision_ia"] = "IA_NON_REQUISE"
            candidat["confiance_ia"] = 0
            candidat["raison_ia"] = (
                f"Premier filtre Python non atteint : score {score} < {GEMINI_SCORE_MIN}."
            )
            candidat["ia_analyse"] = "NON"
            candidat["statut_final"] = candidat.get("statut", "À vérifier")
            non_eligibles.append(candidat)

    # Les plus pertinents passent en premier.
    candidats.sort(
        key=lambda x: (
            str(ANNEE_PRIORITAIRE) in str(x.get("annee_detectee", "")),
            x.get("score", 0),
        ),
        reverse=True,
    )

    # Protection quota : même si 100+ candidats passent le premier filtre,
    # seuls les N meilleurs sont envoyés à Gemini sur cette exécution.
    candidats = candidats[:MAX_IA_ANALYSES]
    limite_non_analyse = []
    for candidat in (pertinents + a_verifier):
        if candidat in non_eligibles or candidat in candidats:
            continue
        candidat["decision_ia"] = "IA_NON_ANALYSE"
        candidat["confiance_ia"] = 0
        candidat["raison_ia"] = (
            f"Candidat éligible Gemini mais hors quota local : top {MAX_IA_ANALYSES} uniquement."
        )
        candidat["ia_analyse"] = "NON"
        candidat["statut_final"] = candidat.get("statut", "À vérifier")
        limite_non_analyse.append(candidat)

    memoire = charger_memoire()
    en_memoire = {
        id(c) for c in candidats
        if memoire_valide(memoire.get(cle_memoire(c)))
    }
    nouveaux = [c for c in candidats if id(c) not in en_memoire]
    enrichir_candidats_ia(nouveaux)

    # L'enrichissement apporte date/lieu/organisateur et permet un second
    # calcul avant l'appel Gemini, sans élargir la sélection initiale.
    for candidat in nouveaux:
        score_data = calculer_score(candidat)
        candidat.update(score_data)

    analyses = 0
    ia_resultats = []
    nouveaux_pertinents = []
    nouveaux_a_verifier = []
    quota_bloque = False
    erreurs_consecutives = 0
    dernier_appel = 0.0
    motif_arret = ""

    print()
    print("========================================")
    print("COUCHE IA GEMINI")
    print("========================================")
    print(f"Premier filtre Python : score >= {GEMINI_SCORE_MIN}")
    print(f"Candidats ayant passé le filtre : {len(pertinents) + len(a_verifier) - len(non_eligibles)}")
    print(f"Candidats envoyés à Gemini : {len(candidats)}")
    print(f"Non éligibles Gemini : {len(non_eligibles)}")
    print(f"Éligibles mais hors top {MAX_IA_ANALYSES} : {len(limite_non_analyse)}")
    print(f"Déjà en mémoire (réutilisés sans appel) : {len(en_memoire)}")
    print(f"Nouveaux appels prévus : {len(nouveaux)}")
    print(f"Modèle Gemini : {os.getenv('GEMINI_MODEL', 'gemini-3.8-flash')}")
    print(f"Cooldown : {COOLDOWN_IA_SECONDS} s")
    total_prevu = len(nouveaux)
    print(
        f"Durée maximale théorique de la partie cooldown : "
        f"{max(0, total_prevu - 1) * COOLDOWN_IA_SECONDS / 60:.1f} min"
    )
    print(flush=True)

    for resultat in candidats:
        statut_python = str(resultat.get("statut", "À vérifier"))
        cle = cle_memoire(resultat)
        depuis_memoire = id(resultat) in en_memoire

        if not depuis_memoire and analyses >= MAX_IA_ANALYSES:
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

        if not depuis_memoire and quota_bloque:
            resultat["decision_ia"] = "IA_NON_ANALYSE"
            resultat["confiance_ia"] = 0
            resultat["raison_ia"] = (
                f"Non analysé : appels IA arrêtés ({motif_arret}). "
                "Candidat conservé pour vérification."
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

        if depuis_memoire:
            entree = memoire[cle]
            analyse = {
                "decision_ia": entree.get("decision_ia", ""),
                "confiance_ia": int(float(entree.get("confiance_ia") or 0)),
                "raison_ia": entree.get("raison_ia", ""),
                "categorie_ia": entree.get("categorie_ia", ""),
                "date_ia": entree.get("date_ia", ""),
                "departement_ia": entree.get("departement_ia", ""),
                "erreur_ia_type": "",
            }
        else:
            if dernier_appel:
                attente = COOLDOWN_IA_SECONDS - (time.time() - dernier_appel)
                if attente > 0:
                    time.sleep(attente)

            position = analyses + 1
            print(
                f"[IA] {position}/{total_prevu} — "
                f"score {resultat.get('score', 0)} — {court(resultat.get('titre', ''))}",
                flush=True,
            )
            analyse = analyser_resultat(resultat)
            dernier_appel = time.time()
            analyses += 1

            if analyse.get("decision_ia") in MEMOIRE_VALIDITE_JOURS:
                memoire[cle] = {
                    "cle": cle,
                    "url": resultat.get("url", ""),
                    "titre": resultat.get("titre", ""),
                    "decision_ia": analyse.get("decision_ia", ""),
                    "confiance_ia": analyse.get("confiance_ia", 0),
                    "raison_ia": analyse.get("raison_ia", ""),
                    "categorie_ia": analyse.get("categorie_ia", ""),
                    "date_ia": analyse.get("date_ia", ""),
                    "departement_ia": analyse.get("departement_ia", ""),
                    "page_lue": resultat.get("page_enrichie", "NON"),
                    "date_analyse": datetime.now().isoformat(timespec="seconds"),
                    "version_regles": REGLES_VERSION,
                }

        resultat.update(analyse)
        resultat["ia_analyse"] = "MEMOIRE" if depuis_memoire else "OUI"
        decision = str(analyse.get("decision_ia", "IA_INDISPONIBLE"))
        resultat["date_ia"] = normaliser_date_ia(resultat.get("date_ia", ""))

        type_erreur = str(analyse.get("erreur_ia_type", ""))
        if not depuis_memoire:
            if decision == "IA_INDISPONIBLE":
                erreurs_consecutives += 1
            else:
                erreurs_consecutives = 0

        quota_detecte = False
        if type_erreur in {"QUOTA_JOUR", "QUOTA_MINUTE", "CLE_ABSENTE"}:
            quota_detecte = True
            motif_arret = {
                "QUOTA_JOUR": "quota quotidien Gemini atteint",
                "QUOTA_MINUTE": "quota par minute persistant",
                "CLE_ABSENTE": "clé GEMINI_API_KEY absente",
            }[type_erreur]
        elif erreurs_consecutives >= MAX_ERREURS_IA_CONSECUTIVES:
            quota_detecte = True
            motif_arret = f"{erreurs_consecutives} erreurs Gemini consécutives"

        icone = {
            "PERTINENT": "🟢",
            "A_VERIFIER": "🟠",
            "NON_PERTINENT": "🔴",
            "IA_INDISPONIBLE": "⚠️",
        }.get(decision, "•")
        detail = f"{resultat.get('confiance_ia', 0)}%" if decision != "IA_INDISPONIBLE" else (type_erreur or "erreur")
        etiquette = "MÉMOIRE" if depuis_memoire else f"IA {analyses}/{total_prevu}"
        print(
            f"[{etiquette}] {icone} {decision} ({detail}) | dép. "
            f"{resultat.get('departement', '?')} | score {resultat.get('score', 0)} | "
            f"{court(resultat.get('titre', ''))}",
            flush=True,
        )

        if quota_detecte:
            quota_bloque = True
            print(f"⛔ Arrêt des appels IA : {motif_arret}", flush=True)
            resultat["statut_final"] = statut_python
            ia_resultats.append(resultat)
            if statut_python == "Pertinent":
                nouveaux_pertinents.append(resultat)
            else:
                nouveaux_a_verifier.append(resultat)
            continue

        classer_statut_final(resultat, decision)
        if decision == "PERTINENT":
            nouveaux_pertinents.append(resultat)
        elif decision == "NON_PERTINENT":
            pass
        else:
            nouveaux_a_verifier.append(resultat)
        ia_resultats.append(resultat)

    sauvegarder_memoire(memoire)

    # Réintégration : les candidats hors IA restent disponibles dans la sortie
    # et conservent leur classement Python initial.
    hors_ia = non_eligibles + limite_non_analyse
    ia_resultats.extend(hors_ia)
    for resultat in hors_ia:
        if resultat.get("statut") == "Pertinent":
            nouveaux_pertinents.append(resultat)
        else:
            nouveaux_a_verifier.append(resultat)

    print()
    non_traites = sum(1 for x in ia_resultats if x.get("decision_ia") == "IA_NON_ANALYSE")
    print(f"Analyses IA effectuées : {analyses}")
    print(f"Candidats non analysés (conservés) : {non_traites}")
    print(f"Blocage quota Gemini : {'OUI' if quota_bloque else 'NON'}" + (f" ({motif_arret})" if quota_bloque else ""))
    print(
        "Décisions IA : "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'PERTINENT')} pertinent / "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'A_VERIFIER')} à vérifier / "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'NON_PERTINENT')} non pertinent / "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'IA_INDISPONIBLE')} indisponible / "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'IA_NON_REQUIS' or x.get('decision_ia') == 'IA_NON_REQUISE')} non requis / "
        f"{sum(1 for x in ia_resultats if x.get('decision_ia') == 'IA_NON_ANALYSE')} non analysé"
    )

    return nouveaux_pertinents, nouveaux_a_verifier, ia_resultats


# ============================================================
# SUIVI PRÉCIS DES TÂCHES
# ============================================================

SUIVI_TACHES_JSON = DATA / "suivi_taches.json"
SUIVI_TACHES_CSV = DATA / "suivi_taches.csv"


def _nouvelle_tache(nom, description=""):
    return {
        "id": f"T{len(_SUIVI_TACHES)+1:02d}",
        "tache": nom,
        "description": description,
        "statut": "EN ATTENTE",
        "debut": "",
        "fin": "",
        "duree_s": 0,
        "details": "",
        "erreur": "",
    }


_SUIVI_TACHES = []


def initialiser_suivi_taches():
    global _SUIVI_TACHES
    noms = [
        ("Collecte RSS", "Interrogation des requêtes legacy + V2"),
        ("Collecte sources externes", "Sites institutionnels, OpenAgenda et autres sources"),
        ("Dédoublonnage", "Fusion des résultats identiques ou quasi identiques"),
        ("Scoring / qualification", "Détection événement, cible, géographie et période"),
        ("Enrichissement pages", "Récupération des dates, lieux et organisateurs"),
        ("Analyse IA", "Analyse Gemini des cas ambigus"),
        ("Delta / historique", "Comparaison avec la veille précédente"),
        ("Statistiques opérateurs", "Mesure du rendement des requêtes"),
        ("Sauvegarde CSV", "Écriture des résultats et audits"),
        ("Synchronisation Excel", "Mise à jour du classeur"),
        ("Préparation Teams", "Construction de la synthèse quotidienne"),
        ("Envoi Teams", "Notification Teams si configurée et autorisée"),
    ]
    _SUIVI_TACHES = [_nouvelle_tache(n, d) for n, d in noms]


def demarrer_tache(index, details=""):
    t = _SUIVI_TACHES[index]
    t["statut"] = "EN COURS"
    t["debut"] = datetime.now().isoformat(timespec="seconds")
    t["details"] = details
    return t


def terminer_tache(index, statut="TERMINEE", details="", erreur=""):
    t = _SUIVI_TACHES[index]
    t["fin"] = datetime.now().isoformat(timespec="seconds")
    t["statut"] = statut
    if details:
        t["details"] = details
    if erreur:
        t["erreur"] = erreur
    if t["debut"]:
        try:
            debut = datetime.fromisoformat(t["debut"])
            fin = datetime.fromisoformat(t["fin"])
            t["duree_s"] = round((fin - debut).total_seconds(), 2)
        except Exception:
            pass
    print(f"[TÂCHE {t['id']}] {t['statut']} — {t['tache']} — {t['details']}", flush=True)
    return t


def echouer_tache(index, erreur):
    return terminer_tache(index, "ERREUR", erreur=erreur, details="La tâche a échoué")


def sauvegarder_suivi_taches():
    DATA.mkdir(exist_ok=True)
    SUIVI_TACHES_JSON.write_text(
        json.dumps({
            "date_generation": datetime.now().isoformat(timespec="seconds"),
            "taches": _SUIVI_TACHES,
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    pd.DataFrame(_SUIVI_TACHES).to_csv(
        SUIVI_TACHES_CSV, index=False, encoding="utf-8-sig"
    )


def resume_suivi_taches():
    total = len(_SUIVI_TACHES)
    terminees = sum(t["statut"] == "TERMINEE" for t in _SUIVI_TACHES)
    erreurs = sum(t["statut"] == "ERREUR" for t in _SUIVI_TACHES)
    en_cours = sum(t["statut"] == "EN COURS" for t in _SUIVI_TACHES)
    return {"total": total, "terminees": terminees, "erreurs": erreurs, "en_cours": en_cours}


# ============================================================
# PRIORITÉS 8-10 : DELTA / OPÉRATEURS / TEAMS
# ============================================================

def _cle_resultat_delta(resultat):
    """Clé stable pour comparer deux veilles successives."""
    url = canoniser_url(
        resultat.get("url_finale")
        or resultat.get("url_reelle")
        or resultat.get("url")
    )
    if url:
        return "url:" + url
    titre = cle_titre_evenement(resultat.get("titre_page") or resultat.get("titre"))
    return "titre:" + titre


def charger_ancienne_veille():
    chemin = DATA / "resultats_rss.csv"
    if not chemin.exists():
        return []
    try:
        df = pd.read_csv(chemin, dtype=str, keep_default_na=False)
        return df.to_dict("records")
    except Exception as e:
        print(f"⚠️ Ancienne veille illisible pour le delta : {type(e).__name__} : {e}")
        return []


def calculer_delta_veille(ancienne, actuels):
    """Priorité 8 : détecte les mouvements réellement utiles."""
    anciens = {_cle_resultat_delta(x): x for x in ancienne}
    courants = {_cle_resultat_delta(x): x for x in actuels}
    nouveaux_pertinents = []
    promus_pertinents = []
    nouveaux_a_verifier_2027 = []
    evolutions = []

    for cle, courant in courants.items():
        ancien = anciens.get(cle)
        statut_c = courant.get("statut_final") or courant.get("statut", "")
        statut_a = (ancien or {}).get("statut_final") or (ancien or {}).get("statut", "")

        if ancien is None:
            if statut_c == "Pertinent":
                nouveaux_pertinents.append(courant)
            elif statut_c == "À vérifier" and "2027" in str(courant.get("annee_detectee", "")):
                nouveaux_a_verifier_2027.append(courant)
            continue

        if statut_c == "Pertinent" and statut_a != "Pertinent":
            promus_pertinents.append(courant)

        score_c = int(float(courant.get("score", 0) or 0))
        score_a = int(float(ancien.get("score", 0) or 0))
        priorite_c = courant.get("priorite_2027", "")
        priorite_a = ancien.get("priorite_2027", "")
        if statut_c != statut_a or abs(score_c - score_a) >= 8 or priorite_c != priorite_a:
            evolutions.append(courant)

    nouveaux_pertinents.sort(key=lambda x: (x.get("priorite_2027") != "FORTE", -int(x.get("score", 0) or 0)))
    promus_pertinents.sort(key=lambda x: -int(x.get("score", 0) or 0))
    nouveaux_a_verifier_2027.sort(key=lambda x: -int(x.get("score", 0) or 0))

    return {
        "nouveaux_pertinents": nouveaux_pertinents,
        "promus_pertinents": promus_pertinents,
        "nouveaux_a_verifier_2027": nouveaux_a_verifier_2027,
        "evolutions_utiles": len(evolutions),
        "evolutions": evolutions,
        "premiere_veille": not bool(ancienne),
    }


def generer_stats_operateurs(resultats_bruts, actuels):
    """Priorité 9 : mesure les opérateurs sur les résultats uniques et utiles."""
    statut_par_cle = {
        _cle_resultat_delta(r): (r.get("statut_final") or r.get("statut", ""))
        for r in actuels
    }
    groupes = {}
    for r in resultats_bruts:
        operateur = r.get("mot_cle") or "non défini"
        famille = r.get("famille_operateur") or r.get("operateur_web") or "legacy"
        cle = _cle_resultat_delta(r)
        statut = statut_par_cle.get(cle, "")
        g = groupes.setdefault(operateur, {
            "operateur": operateur,
            "famille": famille,
            "resultats_bruts": 0,
            "uniques": set(),
            "pertinents": 0,
            "a_verifier": 0,
            "non_pertinents": 0,
        })
        g["resultats_bruts"] += 1
        g["uniques"].add(cle)
        if statut == "Pertinent":
            g["pertinents"] += 1
        elif statut == "À vérifier":
            g["a_verifier"] += 1
        elif statut == "Non pertinent":
            g["non_pertinents"] += 1

    lignes = []
    for g in groupes.values():
        uniques = len(g["uniques"])
        utiles = g["pertinents"] + g["a_verifier"]
        lignes.append({
            "operateur": g["operateur"],
            "famille": g["famille"],
            "resultats_bruts": g["resultats_bruts"],
            "uniques": uniques,
            "pertinents": g["pertinents"],
            "a_verifier": g["a_verifier"],
            "non_pertinents": g["non_pertinents"],
            "utiles": utiles,
            "taux_utilite": round((utiles / uniques) * 100, 1) if uniques else 0,
            "taux_pertinent": round((g["pertinents"] / uniques) * 100, 1) if uniques else 0,
        })
    lignes.sort(key=lambda x: (x["pertinents"], x["a_verifier"], x["utiles"], x["taux_utilite"]), reverse=True)
    return lignes


def executer_sorties_v8_v10(delta, stats_operateurs, pertinents, a_verifier, ia_resultats):
    """Génère les fichiers d'audit et, si configuré, la notification Teams."""
    DATA.mkdir(exist_ok=True)
    OUTPUT.mkdir(exist_ok=True)
    lignes_delta = []
    for type_delta, groupe in (
        ("nouveau_pertinent", delta["nouveaux_pertinents"]),
        ("promu_pertinent", delta["promus_pertinents"]),
        ("nouveau_2027_a_verifier", delta["nouveaux_a_verifier_2027"]),
    ):
        for r in groupe:
            lignes_delta.append({
                "type_delta": type_delta,
                "titre": r.get("titre", ""),
                "url": r.get("url", ""),
                "source": r.get("source", ""),
                "departement": r.get("departement", ""),
                "score": r.get("score", ""),
                "statut_final": r.get("statut_final", r.get("statut", "")),
                "priorite_2027": r.get("priorite_2027", ""),
                "date_evenement_page": r.get("date_evenement_page", ""),
            })
    pd.DataFrame(lignes_delta, columns=[
        "type_delta", "titre", "url", "source", "departement", "score",
        "statut_final", "priorite_2027", "date_evenement_page"
    ]).to_csv(DATA / "delta_veille.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(stats_operateurs).to_csv(DATA / "stats_operateurs.csv", index=False, encoding="utf-8-sig")

    resume_taches = resume_suivi_taches()
    sauvegarder_suivi_taches()

    meta = {
        "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "pertinents": len(pertinents),
        "a_verifier": len(a_verifier),
        "ia_analysees": sum(1 for r in ia_resultats if r.get("ia_analyse") == "OUI"),
        "ia_memoire": sum(1 for r in ia_resultats if r.get("ia_analyse") == "MEMOIRE"),
        "ia_non_analyse": sum(1 for r in ia_resultats if r.get("decision_ia") == "IA_NON_ANALYSE"),
        "quota_bloque": "OUI" if any(str(r.get("decision_ia", "")) == "IA_INDISPONIBLE" for r in ia_resultats) else "NON",
        "taches_terminees": resume_taches["terminees"],
        "taches_total": resume_taches["total"],
        "taches_erreurs": resume_taches["erreurs"],
    }
    card = construire_resume_teams(delta, stats_operateurs, meta)
    sauvegarder_payload(card, OUTPUT / "teams_payload.json")
    resultat_teams = envoyer_teams(card, DATA / "teams_last_sent.txt")
    print(f"Teams : {'OUI' if resultat_teams['envoye'] else 'NON'} — {resultat_teams['raison']}")
    return resultat_teams

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
        "sources_trouvees",
        "description",
        "url_finale",
        "url_reelle",
        "titre_page",
        "description_page",
        "extrait_page",
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
        "origine",
        "niveau_source",
        "score",
        "score_emploi",
        "score_formation",
        "score_superieur",
        "score_apec",
        "score_cible_apec",
        "score_entreprise",
        "score_qualifie",
        "signaux_cible_apec",
        "signaux_superieur",
        "signaux_entreprise",
        "signaux_exclusion",
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
        "preuve_evenement",
        "niveau_source_v9",
        "priorite_2027",
        "geographie_verifiee",
        "annee_detectee",
        "periode",
        "hors_periode",
        "hors_perimetre",
        "hors_geographie",
        "decision_ia",
        "confiance_ia",
        "raison_ia",
        "categorie_ia",
        "date_ia",
        "departement_ia",
        "erreur_ia_type",
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
        if r.get("ia_analyse") in ("OUI", "MEMOIRE")
        or r.get("decision_ia") == "IA_INDISPONIBLE"
        or r.get("decision_ia") == "IA_NON_ANALYSE"
    ]

    def cle_tri_ia(resultat):
        decision = resultat.get("decision_ia", "")
        # 2027 d'abord à l'intérieur de chaque groupe.
        rang_annee = 0 if str(ANNEE_PRIORITAIRE) in str(resultat.get("annee_detectee", "")) else 1

        if decision == "PERTINENT":
            return (0, rang_annee, -int(resultat.get("confiance_ia", 0) or 0))

        if decision == "A_VERIFIER":
            return (1, rang_annee, -int(resultat.get("score", 0) or 0))

        if decision == "IA_INDISPONIBLE":
            return (2, rang_annee, 0)

        if decision == "IA_NON_ANALYSE":
            return (3, rang_annee, -int(resultat.get("score", 0) or 0))

        return (4, rang_annee, 0)

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
    print(f"Version du programme : {VERSION_PROGRAMME} | règles IA : {REGLES_VERSION}")
    print(
        f"Années : ≥ {ANNEE_MIN} (priorité {ANNEE_PRIORITAIRE}) | "
        f"Gemini : seuil >= {GEMINI_SCORE_MIN} | cooldown : {COOLDOWN_IA_SECONDS} s | max analyses : {MAX_IA_ANALYSES}"
    )

    initialiser_suivi_taches()
    sauvegarder_suivi_taches()

    ancienne_veille = charger_ancienne_veille()

    departements = load_yaml("departements.yml")
    requetes = construire_requetes(departements)

    print()
    print(f"Requêtes générées : {len(requetes)}")

    demarrer_tache(0, f"{len(requetes)} requêtes")
    try:
        resultats_bruts, erreurs = collecter(requetes)
        terminer_tache(0, details=f"{len(resultats_bruts)} résultats, {len(erreurs)} erreurs")
    except Exception as e:
        echouer_tache(0, str(e))
        sauvegarder_suivi_taches()
        raise

    # Sources externes (sites officiels, OpenAgenda, Bing News, flux directs).
    # En tête de liste : le dédoublonnage garde la première occurrence, donc
    # la source primaire l'emporte sur l'article de presse.
    demarrer_tache(1)
    try:
        # mot de la zone (nom de département ou ville) -> code département
        zone = {}
        for code, nom in departements["departements"].items():
            zone[nom.lower()] = code
        for code, villes in VILLES_DEPARTEMENTS.items():
            for ville in villes:
                zone.setdefault(ville.lower(), code)
        annees_ext = sorted(
            {int(a) for a in departements["annees"] if int(a) >= ANNEE_MIN}
            | {ANNEE_PRIORITAIRE}
        )
        externes = collecter_sources_externes(
            sources_yml=load_yaml_optionnel("sources.yml"),
            extra=load_yaml_optionnel("sources_collecte.yml"),
            departements=departements["departements"],
            annees=annees_ext,
            session_factory=get_session,
            timeout=TIMEOUT,
            zone=zone,
        )
    except Exception as e:
        print(f"⚠️ Sources externes ignorées : {type(e).__name__} : {e}")
        externes = []
        terminer_tache(1, "ERREUR", erreur=f"{type(e).__name__}: {e}")
    else:
        terminer_tache(1, details=f"{len(externes)} résultats externes")

    resultats_bruts = externes + resultats_bruts

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

    demarrer_tache(2, f"{len(resultats_bruts)} résultats bruts")
    pertinents, a_verifier, non_pertinents = traiter(resultats_bruts)
    terminer_tache(2, details=f"{len(pertinents)+len(a_verifier)+len(non_pertinents)} résultats conservés après traitement")

    demarrer_tache(3)
    terminer_tache(3, details=f"{len(pertinents)} pertinents / {len(a_verifier)} à vérifier / {len(non_pertinents)} non pertinents")

    # --------------------------------------------------------
    # COUCHE IA
    # --------------------------------------------------------
    demarrer_tache(5, f"{len(pertinents) + len(a_verifier)} candidats")
    pertinents_ia, a_verifier_ia, ia_resultats = appliquer_ia(pertinents, a_verifier)
    terminer_tache(5, details=f"{len(ia_resultats)} candidats traités")

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

    # Priorités 8-10 : delta, performance des opérateurs et sortie Teams.
    actuels_utiles = pertinents + a_verifier
    demarrer_tache(6)
    delta = calculer_delta_veille(ancienne_veille, actuels_utiles)
    terminer_tache(6, details=f"{delta['evolutions_utiles']} évolutions détectées")

    demarrer_tache(7)
    stats_operateurs = generer_stats_operateurs(resultats_bruts, pertinents + a_verifier + non_pertinents)
    terminer_tache(7, details=f"{len(stats_operateurs)} opérateurs mesurés")

    print()
    print("========================================")
    print("PRIORITÉS 8-10")
    print("========================================")
    print(f"Nouveaux pertinents : {len(delta['nouveaux_pertinents'])}")
    print(f"Promus pertinents : {len(delta['promus_pertinents'])}")
    print(f"Nouveaux 2027 à vérifier : {len(delta['nouveaux_a_verifier_2027'])}")
    print(f"Évolutions utiles : {delta['evolutions_utiles']}")
    print(f"Opérateurs mesurés : {len(stats_operateurs)}")

    demarrer_tache(8)
    sauvegarde_ok = sauvegarder(pertinents, a_verifier, non_pertinents, resultats_bruts)
    if sauvegarde_ok:
        terminer_tache(8, details="CSV principaux + audits enregistrés")
    else:
        echouer_tache(8, "Sauvegarde refusée ou aucun résultat")

    demarrer_tache(10)
    sorties_v10 = executer_sorties_v8_v10(
        delta, stats_operateurs, pertinents, a_verifier, ia_resultats
    )

    if sorties_v10.get("envoye"):
        terminer_tache(11, details=sorties_v10.get("raison", ""))
    else:
        terminer_tache(11, "TERMINEE", details=sorties_v10.get("raison", ""))

    synchronisation_ok = False

    demarrer_tache(9)
    if sauvegarde_ok:
        try:
            synchronisation_ok = synchroniser_excel()
            if synchronisation_ok:
                terminer_tache(9, details="Classeur Excel mis à jour")
            else:
                terminer_tache(9, "ERREUR", details="La synchronisation n'a pas été validée")
        except Exception as e:
            echouer_tache(9, str(e))
    else:
        terminer_tache(9, "NON EXECUTEE", details="Sauvegarde CSV non validée")

    terminer_tache(10, details="Payload Teams préparé et notification traitée")
    sauvegarder_suivi_taches()
    duree_total = time.time() - debut_total

    print()
    print("========================================")
    print("STATISTIQUES")
    print("========================================")
    print(f"Erreurs : {erreurs}")
    print(f"Temps total : {duree_total:.1f} secondes")
    suivi = resume_suivi_taches()
    print(f"Tâches : {suivi['terminees']}/{suivi['total']} terminées | {suivi['erreurs']} erreur(s) | {suivi['en_cours']} en cours")
    print(f"Suivi détaillé : {SUIVI_TACHES_JSON}")
    print(
        "Sauvegarde validée : "
        f"{'OUI' if sauvegarde_ok else 'NON'}"
    )
    print(
        "Synchronisation Excel : "
        f"{'OUI' if synchronisation_ok else 'NON'}"
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
