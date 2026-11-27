from pathlib import Path
from urllib.parse import quote_plus, urljoin, urlparse
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
import unicodedata
from html import unescape

# ============================================================
# RAPPORT PDF UNIQUE PAR EXECUTION
# ============================================================
try:
    from reportlab.lib import colors
    from xml.sax.saxutils import escape as xml_escape
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    REPORTLAB_DISPONIBLE = True
except ImportError:
    REPORTLAB_DISPONIBLE = False

import feedparser
import pandas as pd
import yaml
import requests
from requests.adapters import HTTPAdapter
from bs4 import BeautifulSoup

from ia import (analyser_resultat_31, analyser_resultat_31_tour2, analyser_resultat_31_tour3, analyser_resultat_38, preanalyser_geographie)
from sources_externes import collecter_sources_externes
from teams import construire_resume_teams, envoyer_teams, sauvegarder_payload

# Teams : envoi privé uniquement via TEAMS_PRIVATE_WEBHOOK_URL.
# Le webhook est fourni par un workflow Microsoft Teams/Power Automate
# configuré pour publier dans la conversation privée de son propriétaire.


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
IA_MEMORY_FILE = DATA / "ia_memoire.json"
IA_MEMORY_VERSION = "v2-global"
GLOBAL_MEMORY_FILE = DATA / "memoire_veille_globale.json"
MEMORY_REFRESH_DAYS = int(os.getenv("MEMORY_REFRESH_DAYS", "7"))
# BDD persistante des candidats déjà analysés
CANDIDATS_BDD_DIR = DATA / "bdd"
CANDIDATS_BDD_FILE = CANDIDATS_BDD_DIR / "candidats_analyses.json"
CANDIDATS_BDD_VERSION = "v1"
# Historique des runs précédentes : utilisé pour enrichir/amorcer la BDD
# sans relancer automatiquement les recherches ou Gemini.
HISTORIQUE_RUNS_DIR = OUTPUT
HISTORIQUE_MAX_FICHIERS = int(os.getenv("HISTORIQUE_MAX_FICHIERS", "100"))
HISTORIQUE_MAX_LIGNES_PAR_FICHIER = int(os.getenv("HISTORIQUE_MAX_LIGNES_PAR_FICHIER", "5000"))


SUIVI_TACHES_JSON = DATA / "suivi_taches.json"
SUIVI_TACHES_CSV = DATA / "suivi_taches.csv"

MAX_WORKERS = 16
TIMEOUT = 10
MAX_RETRIES = 2
COOLDOWN_31_SECONDS = int(os.getenv("GEMINI_31_COOLDOWN_SECONDS", "10"))
COOLDOWN_38_SECONDS = int(os.getenv("GEMINI_38_COOLDOWN_SECONDS", "60"))
MAX_IA_38 = int(os.getenv("MAX_GEMINI_38_ANALYSES", "3"))
IA_TEST_MODE = os.getenv("IA_TEST_MODE", "false").lower() == "true"
IA_TEST_MAX_CANDIDATS = int(os.getenv("IA_TEST_MAX_CANDIDATS", "5"))
IA_TEST_NO_PERSIST = os.getenv("IA_TEST_NO_PERSIST", "false").lower() == "true"
IA_RULES_VERSION = "v42-preuves-multisources-3x31-plus1x38"
GEMINI_SCORE_MIN = int(os.getenv("SEUIL_GEMINI", "0"))  # V36 : pas de seuil bloquant par défaut
MAX_ERREURS_IA_CONSECUTIVES = 3  # disjoncteur : on stoppe après 3 échecs d'affilée
DEPARTEMENTS_CIBLES = {"16", "17", "19", "23", "79", "86", "87"}

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

# Statistiques IA du dernier traitement, exposées au rapport PDF.
# Cela évite de perdre les compteurs car appliquer_ia() possède sa propre portée locale.
LAST_IA_STATS = {
    "appels_31": 0,
    "appels_38": 0,
    "candidats_analyzes": 0,
    "memoire_reutilisee": 0,
    "sources_decouverte": 0,
    "candidats_decouverts": 0,
    "candidats_perdus_reconciliation": 0,
    "geo_bloques": 0,
    "quota_bloque": False,
}

# Audit complet du pipeline du run courant. Les compteurs sont séparés afin
# que le PDF distingue clairement collecte, filtrage, traitement et IA.
AUDIT_PIPELINE = {
    "requetes_generees": 0,
    "rss_brut": 0,
    "externes_brut": 0,
    "total_brut": 0,
    "apres_pre_filtre": 0,
    "rejets_pre_filtre": 0,
    "motifs_rejets": {},
    "apres_traitement": 0,
    "pertinents_python": 0,
    "a_verifier_python": 0,
    "non_pertinents_python": 0,
    "candidats_ia_entree": 0,
    "candidats_ia_reels": 0,
    "resultats_ia": 0,
    "pertinents_final": 0,
    "a_verifier_final": 0,
    "non_pertinents_final": 0,
    "preuves_verifiees": 0,
    "preuves_fortes": 0,
    "preuves_moyennes": 0,
    "preuves_faibles": 0,
    "preuves_absentes": 0,
    "sources_secondaires_trouvees": 0,
}

# Journal détaillé de la phase de recherche RSS.
JOURNAL_RECHERCHE = []
JOURNAL_RECHERCHE_LOCK = threading.Lock()

def enregistrer_log_recherche(requete, **infos):
    """Ajoute une ligne au journal de recherche, de façon thread-safe."""
    ligne = {
        "horodatage": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "departement": requete.get("departement", ""),
        "nom_departement": requete.get("nom_departement", ""),
        "annee": requete.get("annee", ""),
        "categorie": requete.get("categorie", ""),
        "mot_cle": requete.get("mot_cle", ""),
        "complement": requete.get("complement", ""),
        "lieu": requete.get("lieu", ""),
        "operateur_web": requete.get("operateur_web", ""),
        "famille_operateur": requete.get("famille_operateur", ""),
    }
    ligne.update(infos)
    with JOURNAL_RECHERCHE_LOCK:
        JOURNAL_RECHERCHE.append(ligne)

def nettoyer_texte_rapport(valeur):
    return " ".join(unescape(str(valeur or "")).replace("\n", " ").split())

def generer_rapports_recherche(requetes, resultats_bruts, erreurs, duree_collecte,
                               resultats_externes=0, erreurs_externes=0):
    """Produit un journal Word + PDF + CSV de la phase de recherche."""
    try:
        from docx import Document
        from docx.shared import Pt
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    except Exception as exc:
        print(f"⚠️ Export Word/PDF des logs indisponible : {exc}")
        return None

    maintenant = datetime.now()
    horodatage = maintenant.strftime("%Y%m%d_%H%M%S")
    docx_path = OUTPUT / f"journal_recherche_{horodatage}.docx"
    pdf_path = OUTPUT / f"journal_recherche_{horodatage}.pdf"
    csv_path = OUTPUT / f"journal_recherche_{horodatage}.csv"

    with JOURNAL_RECHERCHE_LOCK:
        journal = list(JOURNAL_RECHERCHE)
    journal.sort(key=lambda x: x.get("horodatage", ""))

    ok = sum(1 for x in journal if x.get("statut") == "OK")
    ko = sum(1 for x in journal if x.get("statut") != "OK")
    total_resultats = sum(int(x.get("nb_resultats", 0) or 0) for x in journal)
    pd.DataFrame(journal).to_csv(csv_path, index=False, encoding="utf-8-sig")

    # ---------- Word ----------
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(9)
    doc.add_heading("Journal de recherche – Flux RSS Salon APEC", 0)
    doc.add_paragraph(f"Exécution : {maintenant.strftime('%d/%m/%Y %H:%M:%S')}")
    doc.add_paragraph(f"Version : {VERSION_PROGRAMME} | Règles IA : {REGLES_VERSION}")
    doc.add_paragraph(f"Période : ≥ {ANNEE_MIN} | priorité {ANNEE_PRIORITAIRE}")
    doc.add_paragraph(f"Collecte RSS : {len(requetes)} requêtes | {total_resultats} résultats bruts | {ok} OK | {ko} erreurs | {duree_collecte:.1f} s")
    doc.add_paragraph(f"Sources externes intégrées : {resultats_externes} résultats | erreurs externes : {erreurs_externes}")

    doc.add_heading("1. Résumé", level=1)
    resume=[("Requêtes générées",len(requetes)),("Requêtes journalisées",len(journal)),("Résultats RSS bruts",total_resultats),
            ("Requêtes avec résultats",sum(1 for x in journal if int(x.get("nb_resultats",0) or 0)>0)),
            ("Requêtes en erreur",ko),("Workers",MAX_WORKERS),("Timeout",f"{TIMEOUT} s"),("Tentatives max",MAX_RETRIES)]
    t=doc.add_table(rows=1,cols=2); t.style="Table Grid"
    t.rows[0].cells[0].text="Indicateur"; t.rows[0].cells[1].text="Valeur"
    for a,b in resume:
        c=t.add_row().cells; c[0].text=str(a); c[1].text=str(b)

    doc.add_heading("2. Journal détaillé des requêtes", level=1)
    t=doc.add_table(rows=1,cols=8); t.style="Table Grid"
    for i,h in enumerate(["Heure","Dept.","Année","Catégorie","Mot-clé","Lieu","Statut","Résultats"]): t.rows[0].cells[i].text=h
    for x in journal:
        vals=[x.get("horodatage",""),x.get("departement",""),x.get("annee",""),nettoyer_texte_rapport(x.get("categorie","")),
              nettoyer_texte_rapport(x.get("mot_cle","")),nettoyer_texte_rapport(x.get("lieu","")),x.get("statut",""),x.get("nb_resultats",0)]
        c=t.add_row().cells
        for i,v in enumerate(vals): c[i].text=str(v)

    doc.add_page_break(); doc.add_heading("3. Résultats RSS bruts", level=1)
    doc.add_paragraph("Ces résultats sont ceux obtenus pendant la recherche, avant filtrage métier, scoring et déduplication.")
    t=doc.add_table(rows=1,cols=7); t.style="Table Grid"
    for i,h in enumerate(["Dept.","Année","Source","Publication","Titre","URL","Mot-clé"]): t.rows[0].cells[i].text=h
    for r in resultats_bruts:
        vals=[r.get("departement",""),r.get("annee_recherchee",""),nettoyer_texte_rapport(r.get("source","")),nettoyer_texte_rapport(r.get("date_publication","")),
              nettoyer_texte_rapport(r.get("titre","")),r.get("url",""),nettoyer_texte_rapport(r.get("mot_cle",""))]
        c=t.add_row().cells
        for i,v in enumerate(vals): c[i].text=str(v)
    doc.add_heading("4. Périmètre du journal", level=1)
    doc.add_paragraph("Le présent fichier documente la partie recherche/collecte. Les décisions de pertinence, exclusions, déduplication et analyses Gemini sont appliquées ensuite.")
    doc.save(docx_path)

    # ---------- PDF ----------
    styles=getSampleStyleSheet(); styles.add(ParagraphStyle(name="SmallApec",parent=styles["BodyText"],fontSize=6.2,leading=7.2))
    pdf=SimpleDocTemplate(str(pdf_path),pagesize=landscape(A4),rightMargin=7*mm,leftMargin=7*mm,topMargin=7*mm,bottomMargin=7*mm)
    story=[Paragraph("Journal de recherche – Flux RSS Salon APEC",styles["Title"]),
           Paragraph(f"{maintenant.strftime('%d/%m/%Y %H:%M:%S')} | {len(requetes)} requêtes | {total_resultats} résultats bruts | {ko} erreurs",styles["SmallApec"]),Spacer(1,5)]
    data=[["Heure","Dept","Année","Catégorie","Mot-clé","Lieu","Statut","Résultats"]]
    for x in journal:
        data.append([str(x.get("horodatage","")),str(x.get("departement","")),str(x.get("annee","")),nettoyer_texte_rapport(x.get("categorie",""))[:30],nettoyer_texte_rapport(x.get("mot_cle",""))[:42],nettoyer_texte_rapport(x.get("lieu",""))[:20],str(x.get("statut","")),str(x.get("nb_resultats",0))])
    tbl=Table(data,repeatRows=1,colWidths=[27*mm,11*mm,11*mm,36*mm,55*mm,26*mm,18*mm,15*mm])
    tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.grey),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTSIZE",(0,0),(-1,-1),5.2),("LEADING",(0,0),(-1,-1),6.2),("GRID",(0,0),(-1,-1),0.25,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(tbl); story.append(PageBreak()); story.append(Paragraph("Résultats RSS bruts",styles["Heading2"]))
    data2=[["Dept","Année","Source","Publication","Titre","URL"]]
    for r in resultats_bruts:
        data2.append([str(r.get("departement","")),str(r.get("annee_recherchee","")),nettoyer_texte_rapport(r.get("source",""))[:26],nettoyer_texte_rapport(r.get("date_publication",""))[:22],nettoyer_texte_rapport(r.get("titre",""))[:75],str(r.get("url",""))[:115]])
    tbl2=Table(data2,repeatRows=1,colWidths=[12*mm,13*mm,32*mm,28*mm,80*mm,88*mm])
    tbl2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.grey),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTSIZE",(0,0),(-1,-1),5),("LEADING",(0,0),(-1,-1),6),("GRID",(0,0),(-1,-1),0.25,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(tbl2); story.append(Spacer(1,5)); story.append(Paragraph("Les résultats sont bruts et peuvent être rejetés lors des étapes ultérieures.",styles["SmallApec"]))
    pdf.build(story)
    print(f"📄 Journal recherche Word : {docx_path}"); print(f"📄 Journal recherche PDF : {pdf_path}"); print(f"🧾 Journal recherche CSV : {csv_path}")
    return {"docx":docx_path,"pdf":pdf_path,"csv":csv_path}


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

VERSION_PROGRAMME = "2026-11-27-v42-preuves-multisources"

# À changer quand le prompt de ia.py ou les règles de décision évoluent :
# toutes les décisions mémorisées avec une autre version sont ignorées.
REGLES_VERSION = "2026-11-27-v33-geo-prequeries-strictes"

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
    # Villes hors des 7 départements cibles : uniquement des lieux potentiels.
    "nantes", "saint-nazaire", "angers", "tours", "bordeaux", "agen",
    "nice", "mulhouse", "montpellier", "lille", "chambéry", "saint-malo",
    "langon", "colmar", "condrieu", "alès", "ales", "paris", "toulouse",
    "rennes", "orléans", "orleans", "clermont-ferrand", "dijon", "grenoble",
    "lyon", "marseille", "strasbourg", "metz", "nancy", "rouen", "caen",
    "amiens", "le mans", "le havre", "mont-de-marsan", "pau", "bayonne",
    "périgueux", "perigueux", "bergerac", "libourne", "montauban", "avignon",
    "troyes", "reims", "pyrénées-atlantiques", "gironde", "alsace",
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
# GÉOGRAPHIE STRICTE — PÉRIMÈTRE ABSOLU
# ============================================================
# Aucun candidat ne doit pouvoir rester dans la veille si son lieu
# d'événement est explicitement situé hors des 7 départements cibles.
# Les départements cibles restent exclusivement : 16, 17, 19, 23, 79, 86, 87.
DEPARTEMENTS_CIBLES = {"16", "17", "19", "23", "79", "86", "87"}

# Noms des départements français métropolitains utilisés pour détecter
# une localisation explicite hors périmètre, même lorsque le code postal
# n'est pas présent.
DEPARTEMENTS_FRANCAIS = {
    "01": "ain", "02": "aisne", "03": "allier", "04": "alpes-de-haute-provence",
    "05": "hautes-alpes", "06": "alpes-maritimes", "07": "ardeche", "08": "ardennes",
    "09": "ariege", "10": "aube", "11": "aude", "12": "aveyron", "13": "bouches-du-rhone",
    "14": "calvados", "15": "cantal", "16": "charente", "17": "charente-maritime",
    "18": "cher", "19": "correze", "2a": "corse-du-sud", "2b": "haute-corse",
    "21": "cote-d-or", "22": "cotes-d-armor", "23": "creuse", "24": "dordogne",
    "25": "doubs", "26": "drome", "27": "eure", "28": "eure-et-loir", "29": "finistere",
    "30": "gard", "31": "haute-garonne", "32": "gers", "33": "gironde", "34": "herault",
    "35": "ille-et-vilaine", "36": "indre", "37": "indre-et-loire", "38": "isere",
    "39": "jura", "40": "landes", "41": "loir-et-cher", "42": "loire", "43": "haute-loire",
    "44": "loire-atlantique", "45": "loiret", "46": "lot", "47": "lot-et-garonne",
    "48": "lozere", "49": "maine-et-loire", "50": "manche", "51": "marne", "52": "haute-marne",
    "53": "mayenne", "54": "meurthe-et-moselle", "55": "meuse", "56": "morbihan",
    "57": "moselle", "58": "nievre", "59": "nord", "60": "oise", "61": "orne",
    "62": "pas-de-calais", "63": "puy-de-dome", "64": "pyrenees-atlantiques", "65": "hautes-pyrenees",
    "66": "pyrenees-orientales", "67": "bas-rhin", "68": "haut-rhin", "69": "rhone",
    "70": "haute-saone", "71": "saone-et-loire", "72": "sarthe", "73": "savoie",
    "74": "haute-savoie", "75": "paris", "76": "seine-maritime", "77": "seine-et-marne",
    "78": "yvelines", "79": "deux-sevres", "80": "somme", "81": "tarn", "82": "tarn-et-garonne",
    "83": "var", "84": "vaucluse", "85": "vendee", "86": "vienne", "87": "haute-vienne",
    "88": "vosges", "89": "yonne", "90": "territoire de belfort", "91": "essonne",
    "92": "hauts-de-seine", "93": "seine-saint-denis", "94": "val-de-marne", "95": "val-d-oise",
}

def _normaliser_geo_fort(texte):
    """Normalisation plus robuste pour comparer noms de départements/villes."""
    texte = unicodedata.normalize("NFKD", str(texte or "").lower())
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    texte = re.sub(r"[^a-z0-9\s-]", " ", texte)
    texte = re.sub(r"\s+", " ", texte)
    return texte.strip()



# ============================================================
# OUTILS TEXTE
# ============================================================

def normaliser(texte):
    texte = str(texte or "").lower()
    texte = re.sub(r"\s+", " ", texte)
    return texte.strip()


DEPARTEMENTS_CIBLES_NOMS = {normaliser(v): k for k, v in NOMS_DEPARTEMENTS.items()}

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



def verifier_geographie_inviolable(resultat):
    """Contrôle géographique indépendant, haute priorité.

    Cette couche ne fait confiance ni au département porté par la requête,
    ni à un champ lieu éventuellement mal extrait. Elle recherche d'abord
    les preuves géographiques explicites dans les champs de localisation et
    dans le titre de l'événement. Une preuve hors périmètre explicite prime
    sur toute attribution automatique à un département cible.

    Retourne (ok, motif, code).
    """
    champs_forts = {
        "lieu_page": str(resultat.get("lieu_page", "") or ""),
        "adresse_page": str(resultat.get("adresse_page", "") or ""),
        "lieu": str(resultat.get("lieu", "") or ""),
        "titre_page": str(resultat.get("titre_page", "") or ""),
        "titre": str(resultat.get("titre", "") or ""),
    }
    lieu = _normaliser_geo_fort(" ".join(
        champs_forts[k] for k in ("lieu_page", "adresse_page", "lieu")
    ))
    titres = _normaliser_geo_fort(" ".join(
        champs_forts[k] for k in ("titre_page", "titre")
    ))
    contexte = _normaliser_geo_fort(" ".join(champs_forts.values()))

    codes_hors = set()
    codes_cibles = set()

    # 1. Codes postaux : preuve très forte.
    for source in (lieu, titres):
        for cp in re.findall(r"(?<!\d)(?:0[1-9]|[1-8]\d|9[0-5])\d{3}(?!\d)", source):
            code = cp[:2]
            (codes_cibles if code in DEPARTEMENTS_CIBLES else codes_hors).add(code)

    # 2. Noms de départements dans les champs de lieu/adresse.
    for code, nom in DEPARTEMENTS_FRANCAIS.items():
        nn = _normaliser_geo_fort(nom)
        if not nn:
            continue
        if re.search(r"(?<![a-z])" + re.escape(nn) + r"(?![a-z])", lieu):
            (codes_cibles if code in DEPARTEMENTS_CIBLES else codes_hors).add(code)

    # 3. Villes hors cible : uniquement lorsqu'elles sont dans un contexte
    #    manifestement localisant. Un simple nom d'organisateur ne suffit pas.
    marqueurs_lieu = (
        r"(?:\ba\s+|\bà\s+|\bdans\s+|\bsur\s+|\bchez\s+|"
        r"\blieu\s*[:\-]|\badresse\s*[:\-]|\bsite\s*[:\-]|\bcampus\s*[:\-]|"
        r"\bse tiendra\s+(?:a|à)\s+|\baura lieu\s+(?:a|à)\s+|"
        r"\bmairie\s+|[-–—]\s*)"
    )
    for ville in VILLES_HORS_CIBLE:
        vn = _normaliser_geo_fort(ville)
        if not vn:
            continue
        pattern = rf"{marqueurs_lieu}[^\n]{{0,80}}\b{re.escape(vn)}\b"
        if re.search(pattern, contexte, flags=re.I):
            codes_hors.add(
                next((c for c, n in DEPARTEMENTS_FRANCAIS.items()
                      if _normaliser_geo_fort(n) == vn), "")
            )
            # Ville hors cible connue : le code n'est pas nécessaire pour le rejet.
            resultat["_geo_ville_hors_fort"] = ville
            return False, f"Ville explicitement hors périmètre : {ville}", ""

        # Cas particuliers d'arrondissements / quartiers, ex. « Lyon 3 ».
        if vn == "lyon" and re.search(r"\blyon\s+(?:[1-9]|1er|2e|3e|4e|5e|6e|7e|8e|9e)\b", contexte, flags=re.I):
            resultat["_geo_ville_hors_fort"] = "Lyon"
            return False, "Localisation explicitement hors périmètre : Lyon", "69"

    # 4. Ville cible explicite : elle peut confirmer la zone, mais ne doit
    #    jamais écraser une preuve hors cible.
    for code, villes in VILLES_DEPARTEMENTS.items():
        for ville in villes:
            vn = _normaliser_geo_fort(ville)
            if vn and re.search(r"(?<![a-z])" + re.escape(vn) + r"(?![a-z])", lieu):
                codes_cibles.add(code)
            elif vn and re.search(rf"{marqueurs_lieu}[^\n]{{0,80}}\b{re.escape(vn)}\b", titres, flags=re.I):
                codes_cibles.add(code)

    if codes_hors:
        code = sorted(codes_hors)[0]
        nom = DEPARTEMENTS_FRANCAIS.get(code, code)
        return False, f"Département explicitement hors périmètre : {nom}", code

    if codes_cibles:
        code = sorted(codes_cibles)[0]
        return True, f"Localisation cible explicite : département {code}", code

    # Pas de preuve locale contradictoire : on ne rejette pas.
    return True, "Aucune preuve géographique hors cible détectée", str(resultat.get("departement", "")).strip()

def verifier_geographie_stricte(resultat):
    """Filtre géographique ABSOLU.

    Règle : le lieu de l'événement doit appartenir aux départements
    16/17/19/23/79/86/87. Une localisation explicitement hors périmètre
    entraîne un rejet immédiat. Si une localisation cible est identifiable,
    elle devient la référence. À défaut, le département de la requête est
    conservé seulement s'il n'existe aucun indice géographique contradictoire.

    Retourne (ok, motif, code_detecte).
    """
    code_requete = str(resultat.get("departement", "")).strip()
    if code_requete not in DEPARTEMENTS_CIBLES:
        return False, f"Département de recherche hors périmètre : {code_requete or 'inconnu'}", code_requete

    # Pour déterminer le lieu réel, on privilégie les champs liés au lieu
    # de l'événement. On n'utilise volontairement pas l'organisateur seul.
    texte_lieu = _normaliser_geo_fort(" ".join(str(resultat.get(k, "")) for k in (
        "lieu_page", "lieu", "titre_page", "titre", "description_page", "description"
    )))

    # 1) Code postal explicite dans le contexte géographique.
    codes_postaux = re.findall(r"(?<!\d)(?:0[1-9]|[1-8]\d|9[0-5])\d{3}(?!\d)", texte_lieu)
    for cp in codes_postaux:
        prefixe = cp[:2]
        if prefixe in DEPARTEMENTS_CIBLES:
            resultat["departement"] = prefixe
            resultat["nom_departement"] = NOMS_DEPARTEMENTS[prefixe]
            resultat["geographie_corrigee"] = "OUI"
            resultat["geographie_stricte"] = "OUI"
            return True, f"Code postal cible détecté : {cp}", prefixe
        # Un code postal hors cible est une preuve forte d'exclusion.
        resultat["geographie_stricte"] = "NON"
        return False, f"Code postal hors périmètre : {cp}", prefixe

    # 2) Nom explicite d'un département.
    departements_trouves = []
    for code, nom in DEPARTEMENTS_FRANCAIS.items():
        nom_norm = _normaliser_geo_fort(nom)
        if nom_norm and re.search(r"(?<![a-z])" + re.escape(nom_norm) + r"(?![a-z])", texte_lieu):
            departements_trouves.append(code)

    departements_trouves = sorted(set(departements_trouves))
    if departements_trouves:
        cibles = [c for c in departements_trouves if c in DEPARTEMENTS_CIBLES]
        hors = [c for c in departements_trouves if c not in DEPARTEMENTS_CIBLES]
        if hors and not cibles:
            resultat["geographie_stricte"] = "NON"
            noms = ", ".join(DEPARTEMENTS_FRANCAIS[c] for c in hors[:3])
            return False, f"Département explicitement hors périmètre : {noms}", hors[0]
        if cibles:
            code = cibles[0]
            resultat["departement"] = code
            resultat["nom_departement"] = NOMS_DEPARTEMENTS[code]
            resultat["geographie_corrigee"] = "OUI"
            resultat["geographie_stricte"] = "OUI"
            return True, f"Département cible explicitement détecté : {code}", code

    # 3) Ville cible connue.
    codes_villes = []
    for code, villes in VILLES_DEPARTEMENTS.items():
        for ville in villes:
            ville_norm = _normaliser_geo_fort(ville)
            if ville_norm and re.search(r"(?<![a-z])" + re.escape(ville_norm) + r"(?![a-z])", texte_lieu):
                codes_villes.append(code)
                break
    codes_villes = sorted(set(codes_villes))
    if len(codes_villes) == 1:
        code = codes_villes[0]
        resultat["departement"] = code
        resultat["nom_departement"] = NOMS_DEPARTEMENTS[code]
        resultat["geographie_corrigee"] = "OUI"
        resultat["geographie_stricte"] = "OUI"
        return True, f"Ville cible détectée dans le département {code}", code

    # 4) Aucun indice contradictoire : on conserve le département porté par
    # la requête, mais on marque que la preuve géographique reste indirecte.
    # Il pourra être renforcé par l'enrichissement de page.
    resultat["geographie_stricte"] = "INDIRECTE"
    return True, f"Département de requête cible : {code_requete}, aucune contradiction géographique détectée", code_requete


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
# GÉOGRAPHIE DES PRÉ-REQUÊTES — PÉRIMÈTRE STRICT
# ============================================================

def construire_bloc_geo_recherche(requete):
    """Construit une contrainte géographique explicite pour chaque requête.

    La géographie n'est plus seulement appliquée au filtrage final : elle est
    injectée directement dans la requête Google News. Une requête
    départementale est ancrée sur le département demandé et toutes les villes
    connues de ce département. Une requête ville est ancrée sur cette ville
    uniquement.

    Les villes hors périmètre connues du moteur sont ajoutées en exclusions.
    """
    code = str(requete.get("departement", "")).strip()
    if code not in DEPARTEMENTS_CIBLES:
        return '"__DEPARTEMENT_HORS_CIBLE__"'

    lieu = _normaliser_geo_fort(requete.get("lieu", ""))
    nom_dept = _normaliser_geo_fort(NOMS_DEPARTEMENTS.get(code, ""))

    if lieu:
        positifs = [lieu]
    else:
        # Le département est toujours obligatoire. Les villes servent de
        # points d'ancrage supplémentaires sans remplacer le département.
        positifs = [nom_dept]
        positifs.extend(
            _normaliser_geo_fort(v)
            for v in VILLES_DEPARTEMENTS.get(code, [])
            if v
        )

    positifs = list(dict.fromkeys(x for x in positifs if x))
    if not positifs:
        return '"__GEOGRAPHIE_INVALIDE__"'

    # Pour une requête départementale, on exige le département OU une ville
    # appartenant à ce département. Pour une requête ville, la ville est seule
    # utilisée afin d'éviter que Google News élargisse le territoire.
    if lieu:
        bloc = f'"{lieu}"'
    else:
        bloc = '(' + ' OR '.join(f'"{v}"' for v in positifs) + ')'

    exclusions = []
    for ville_hors in VILLES_HORS_CIBLE:
        ville_norm = _normaliser_geo_fort(ville_hors)
        if ville_norm:
            exclusions.append(f'-"{ville_norm}"')

    return bloc + (" " + " ".join(exclusions) if exclusions else "")


def valider_requete_geographique(requete):
    """Empêche la création d'une requête dont le territoire est hors cible."""
    code = str(requete.get("departement", "")).strip()
    if code not in DEPARTEMENTS_CIBLES:
        return False

    lieu = _normaliser_geo_fort(requete.get("lieu", ""))
    if not lieu:
        return True

    villes_cibles = {
        _normaliser_geo_fort(v)
        for v in VILLES_DEPARTEMENTS.get(code, [])
        if v
    }
    return lieu in villes_cibles

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
                    "lieu": nom,
                })

            for mot in REQUETES_SUPERIEUR:
                requetes.append({
                    "departement": code,
                    "nom_departement": nom,
                    "annee": annee,
                    "categorie_numero": 1,
                    "categorie": "Salon étudiant / enseignement supérieur",
                    "mot_cle": mot,
                    "lieu": nom,
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
                    "lieu": nom,
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
                    "lieu": nom,
                })

            for ville in VILLES_DEPARTEMENTS.get(code, []):
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
                    "lieu": nom,
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
                    "lieu": nom,
                    "operateur_web": "INSTITUTIONNEL",
                    "famille_operateur": "source_institutionnelle",
                })

    # ------------------------------------------------------------
    # NORMALISATION + DEDOUBLONNAGE DES PRE-REQUETES
    # ------------------------------------------------------------
    # Une même recherche pouvait être générée plusieurs fois par les
    # différentes familles d'opérateurs. On déduplique avant tout appel
    # réseau, tout en conservant la première occurrence (ordre prioritaire).
    requetes_valides = []
    signatures_vues = set()

    for requete in requetes:
        if not valider_requete_geographique(requete):
            continue

        requete["geo_recherche"] = construire_bloc_geo_recherche(requete)

        signature = (
            str(requete.get("departement", "")).strip(),
            str(requete.get("annee", "")).strip(),
            normaliser(requete.get("mot_cle", "")),
            normaliser(requete.get("complement", "")),
            normaliser(requete.get("lieu", "")),
            normaliser(requete.get("geo_recherche", "")),
            str(requete.get("operateur_web", "")).strip(),
            str(requete.get("famille_operateur", "")).strip(),
        )

        if signature in signatures_vues:
            continue

        signatures_vues.add(signature)
        requete["id_requete"] = f"Q{len(requetes_valides) + 1:05d}"
        requetes_valides.append(requete)

    print(
        f"Pré-requêtes : {len(requetes)} générées -> "
        f"{len(requetes_valides)} uniques et géographiquement valides"
    )

    return requetes_valides


# ============================================================
# GOOGLE NEWS RSS
# ============================================================

def rechercher_google_news(requete):

    if not valider_requete_geographique(requete):
        enregistrer_log_recherche(
            requete, statut="REJET_GEO", nb_resultats=0,
            duree_ms=0, url_recherche="", tentative=0,
            erreur="Requête hors périmètre géographique"
        )
        return {
            "ok": False,
            "resultats": [],
            "erreur": "Requête hors périmètre géographique",
        }

    bloc_geo = construire_bloc_geo_recherche(requete)
    texte = (
        f'"{requete["mot_cle"]}" '
        f'{bloc_geo} '
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
    debut_recherche = time.time()

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
                    "lieu": requete.get("lieu", ""),
                    "geo_recherche": requete.get("geo_recherche") or construire_bloc_geo_recherche(requete),
                    "id_requete": requete.get("id_requete", ""),
                })

            enregistrer_log_recherche(
                requete, statut="OK", nb_resultats=len(resultats),
                duree_ms=round((time.time() - debut_recherche) * 1000),
                url_recherche=url, tentative=tentative + 1, erreur=""
            )
            return {
                "ok": True,
                "resultats": resultats,
                "erreur": "",
            }

        except Exception as erreur:
            derniere_erreur = str(erreur)
            if tentative < MAX_RETRIES - 1:
                time.sleep(0.5)

    enregistrer_log_recherche(
        requete, statut="ERREUR", nb_resultats=0,
        duree_ms=round((time.time() - debut_recherche) * 1000),
        url_recherche=url, tentative=MAX_RETRIES, erreur=derniere_erreur
    )
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



def _signature_memoire(resultat):
    """Signature du contenu entrant AVANT enrichissement de la page.

    On ne met volontairement pas les champs issus de l'enrichissement HTTP
    dans la signature : sinon une page déjà enrichie ne pourrait jamais être
    reconnue au lancement suivant, puisque ces champs sont absents au départ.
    """
    base = {
        # L'URL RSS/origine est privilégiée car elle reste stable même si la
        # redirection finale du site change.
        "url": canoniser_url(resultat.get("url") or resultat.get("url_reelle") or resultat.get("url_finale")),
        "titre": normaliser(resultat.get("titre", "")),
        "description": normaliser(resultat.get("description", ""))[:5000],
        "date_publication": str(resultat.get("date_publication", "")),
        "departement": str(resultat.get("departement", "")),
        "annee_recherchee": str(resultat.get("annee_recherchee", "")),
        "annee_detectee": str(resultat.get("annee_detectee", "")),
        "source": normaliser(resultat.get("source", "")),
        "mot_cle": normaliser(resultat.get("mot_cle", "")),
        "categorie": normaliser(resultat.get("categorie", "")),
        "categorie_numero": str(resultat.get("categorie_numero", "")),
        "complement": normaliser(resultat.get("complement", "")),
        "lieu": normaliser(resultat.get("lieu", "")),
        "operateur_web": normaliser(resultat.get("operateur_web", "")),
        "famille_operateur": normaliser(resultat.get("famille_operateur", "")),
    }
    brut = json.dumps(base, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(brut.encode("utf-8")).hexdigest()


def _cle_memoire_globale(resultat):
    """Clé stable d'un résultat, indépendante de l'analyse IA."""
    url = canoniser_url(resultat.get("url") or resultat.get("url_reelle") or resultat.get("url_finale"))
    if url:
        return "url:" + url
    titre = cle_titre_evenement(resultat.get("titre_page") or resultat.get("titre"))
    dept = str(resultat.get("departement", "")).strip()
    annee = str(resultat.get("annee_detectee") or resultat.get("annee_recherchee") or "").strip()
    return f"titre:{dept}:{annee}:{titre}" if titre else ""


def charger_memoire_globale():
    """Charge la mémoire de veille pour tous les résultats."""
    if not GLOBAL_MEMORY_FILE.exists():
        return {}
    try:
        data = json.loads(GLOBAL_MEMORY_FILE.read_text(encoding="utf-8"))
        return data.get("items", {}) if isinstance(data, dict) and isinstance(data.get("items", {}), dict) else {}
    except Exception as e:
        print(f"⚠️ Mémoire globale illisible : {type(e).__name__} : {e}")
        return {}


def appliquer_memoire_globale(resultats, memoire):
    """Réutilise les données récentes avant les enrichissements lourds."""
    reutilises = 0
    a_rafraichir = []
    maintenant = datetime.now()
    for r in resultats:
        cle = _cle_memoire_globale(r)
        item = memoire.get(cle) if cle else None
        if not item:
            a_rafraichir.append(r)
            continue
        try:
            date_mem = datetime.fromisoformat(str(item.get("date_memorisation", "")).replace("Z", ""))
            age = (maintenant - date_mem).days
        except Exception:
            age = MEMORY_REFRESH_DAYS + 1
        if item.get("signature") == _signature_memoire(r) and age < MEMORY_REFRESH_DAYS:
            for k, v in (item.get("champs") or {}).items():
                if v not in ("", None):
                    r[k] = v
            r["memoire_globale"] = "OUI"
            r["date_memoire_globale"] = item.get("date_memorisation", "")
            reutilises += 1
        else:
            r["memoire_globale"] = "NON"
            a_rafraichir.append(r)
    print(f"🧠 Mémoire globale : {reutilises} résultats réutilisés / {len(resultats)}")
    print(f"🔎 Nouveaux/modifiés/à rafraîchir : {len(a_rafraichir)}")
    return a_rafraichir, reutilises


def sauvegarder_memoire_globale(resultats, ancienne_memoire=None):
    """Conserve la veille pour éviter les retraitements lors des runs suivantes."""
    memoire = dict(ancienne_memoire or {})
    champs = [
        "titre", "url", "url_reelle", "url_finale", "source", "description",
        "date_publication", "departement", "nom_departement", "annee_recherchee",
        "annee_detectee", "categorie_numero", "categorie", "mot_cle", "complement",
        "lieu", "operateur_web", "famille_operateur",
        # Mémoire complète du scoring Python : le score et chacun de ses
        # composants sont réutilisés au prochain passage si la signature
        # du résultat n'a pas changé.
        "score", "score_cible_apec", "score_emploi", "score_formation",
        "score_superieur", "score_apec", "score_entreprise", "score_qualifie",
        "score_geographie", "score_date", "score_bruit", "score_complementaire",
        "bonus_titre", "signaux_cible_apec", "signaux_superieur",
        "signaux_entreprise", "signaux_exclusion", "details_score",
        "statut", "raison", "motif", "hors_periode", "periode",
        "evenement_detecte", "preuve_evenement", "geographie_verifiee",
        "hors_perimetre", "hors_geographie", "geographie_stricte", "geographie_stricte_motif", "erreur_geographique", "annee_detectee",
        "priorite_2027", "niveau_source_v9",
        "statut_final", "decision_ia", "confiance_ia", "raison_ia",
        "page_enrichie", "date_evenement_page", "lieu_page", "organisateur_page",
    ]
    for r in resultats:
        if not isinstance(r, dict):
            continue
        cle = _cle_memoire_globale(r)
        if not cle:
            continue
        memoire[cle] = {
            "signature": _signature_memoire(r),
            "date_memorisation": datetime.now().isoformat(timespec="seconds"),
            "champs": {k: r.get(k, "") for k in champs if k in r},
        }
    GLOBAL_MEMORY_FILE.parent.mkdir(exist_ok=True)
    GLOBAL_MEMORY_FILE.write_text(
        json.dumps({"version": "v1", "updated_at": datetime.now().isoformat(timespec="seconds"), "items": memoire}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return memoire


def charger_memoire_ia():
    """Charge la mémoire des enrichissements IA des runs précédentes."""
    if not IA_MEMORY_FILE.exists():
        return {}
    try:
        data = json.loads(IA_MEMORY_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {}
        if data.get("version") != IA_MEMORY_VERSION:
            return {}
        return data.get("items", {}) if isinstance(data.get("items", {}), dict) else {}
    except Exception as e:
        print(f"⚠️ Mémoire IA illisible : {type(e).__name__} : {e}")
        return {}


def sauvegarder_memoire_ia(resultats, ancienne_memoire=None):
    """Sauvegarde les enrichissements 3.1/3.8 et les données de page.

    La mémoire est cumulative : une run ne supprime pas les événements vus
    précédemment. Cela permet de reconnaître un événement lorsqu'il ressort
    d'un flux plusieurs jours plus tard.
    """
    memoire = dict(ancienne_memoire or {})
    champs_page = [
        "url_finale", "url_reelle", "titre_page", "description_page",
        "extrait_page", "date_evenement_page", "lieu_page",
        "organisateur_page", "page_enrichie", "erreur_enrichissement",
    ]
    champs_ia = [
        "modele_ia_1", "decision_ia_1", "confiance_ia_1", "raison_ia_1",
        "modele_ia_2", "decision_ia_2", "confiance_ia_2", "raison_ia_2",
        "modele_ia_31", "decision_ia_31", "confiance_ia_31", "raison_ia_31",
        "modele_ia_38", "decision_ia_38", "confiance_ia_38", "raison_ia_38",
        "enrichissement_38", "escalade_38", "motif_escalade_38",
        "decision_ia", "confiance_ia", "raison_ia", "categorie_ia",
        "date_ia", "departement_ia", "erreur_ia_type", "geographie_corrigee",
        "statut_final", "ia_analyse", "modele_ia",
    ]
    for r in resultats:
        if not isinstance(r, dict):
            continue
        if r.get("ia_analyse") not in {"OUI", "MEMOIRE"} and not r.get("modele_ia_1"):
            continue
        key = _cle_memoire_evenement(r)
        if not key:
            continue
        item = {
            "signature": _signature_memoire(r),
            "date_memorisation": datetime.now().isoformat(),
            "titre": r.get("titre", ""),
            "url": r.get("url", ""),
            "departement": r.get("departement", ""),
            "page": {k: r.get(k, "") for k in champs_page},
            "ia": {k: r.get(k, "") for k in champs_ia if k in r},
        }
        memoire[key] = item
    IA_MEMORY_FILE.parent.mkdir(exist_ok=True)
    IA_MEMORY_FILE.write_text(
        json.dumps({"version": IA_MEMORY_VERSION, "updated_at": datetime.now().isoformat(), "items": memoire}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return memoire


def _cle_memoire_evenement(resultat):
    """Clé persistante : URL canonique prioritaire, sinon titre+département+année."""
    url = canoniser_url(resultat.get("url") or resultat.get("url_reelle") or resultat.get("url_finale"))
    if url:
        return "url:" + url
    titre = cle_titre_evenement(resultat.get("titre_page") or resultat.get("titre"))
    dept = str(resultat.get("departement", "")).strip()
    annee = str(resultat.get("annee_detectee") or resultat.get("annee_recherchee") or "").strip()
    if titre:
        return f"titre:{dept}:{annee}:{titre}"
    return ""


def appliquer_memoire_ia(candidats, memoire):
    """Réinjecte les enrichissements déjà connus et retourne les candidats à traiter."""
    reutilises = []
    a_traiter = []
    for r in candidats:
        key = _cle_memoire_evenement(r)
        item = memoire.get(key) if key else None
        signature = _signature_memoire(r)
        if item and item.get("signature") == signature:
            for k, v in (item.get("page") or {}).items():
                if v not in ("", None):
                    r[k] = v
            for k, v in (item.get("ia") or {}).items():
                if v not in ("", None):
                    r[k] = v
            r["ia_analyse"] = "MEMOIRE"
            r["memoire_ia"] = "OUI"
            r["date_ia_memoire"] = item.get("date_memorisation", "")
            r["statut_final"] = r.get("statut_final") or r.get("statut", "À vérifier")
            reutilises.append(r)
        else:
            r["memoire_ia"] = "NON"
            a_traiter.append(r)
    print(f"🧠 Mémoire IA : {len(reutilises)} enrichissements réutilisés / {len(candidats)} candidats")
    print(f"🆕 À enrichir/analyser réellement : {len(a_traiter)}")
    return a_traiter, reutilises

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
    print(f"Candidats à enrichir : {len(candidats)}")

    debut = time.time()
    cible = [r for r in candidats if str(r.get("page_enrichie", "")).upper() != "OUI"]
    decoder_liens_google(cible)

    if not cible:
        print("Toutes les pages candidates sont déjà présentes en mémoire : aucun nouvel enrichissement HTTP.")
        return candidats

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
        # Le score représente uniquement des preuves positives.
        # Une absence de preuve ne produit jamais de score négatif.
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
        # Pas de malus : une source faible n'ajoute simplement aucun point.
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
        # La géographie est traitée comme une condition de décision, pas comme
        # une pénalité. Aucun score négatif ne peut être généré.
        raisons.append("géographie cible non démontrée")
    if hors_geo:
        raisons.append("hors géographie: " + ", ".join(hors_geo[:3]))

    # Exclusions fortes.
    if exclusions:
        # Les exclusions fortes font basculer le statut, mais ne créent pas
        # de score négatif. Le score reste lisible comme un niveau de preuve.
        raisons.append("hors cible probable: " + ", ".join(exclusions[:5]))

    if editorial and not (direct or entreprise or superieur):
        # Un contenu éditorial n'est pas récompensé, mais n'est pas pénalisé
        # artificiellement non plus. Le statut final gère l'exclusion.
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

    # Invariant V31 : un score affiché ou enregistré ne peut jamais être négatif.
    score = max(0, int(score or 0))

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
        "score_geographie": 5 if geo_ok else 0,
        "score_date": 4 if date_evenement else 0,
        "score_bruit": 0,
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




# ============================================================
# SOURCES DE DÉCOUVERTE — V40
# ============================================================
MOTS_SOURCE_DECOUVERTE = {
    "insertion professionnelle", "orientation et insertion", "orientation & insertion",
    "stages et emplois", "emploi et carrière", "emploi et formation",
    "rencontrer et recruter", "rencontrer nos étudiants", "recruter nos étudiants",
    "entreprises, partenaires", "entreprises partenaires", "club des partenaires",
    "réseau alumni", "association des anciens", "débouchés métiers",
    "débouchés et insertion", "accompagner les entreprises", "participer à des salons",
    "salons, manifestations et évènements", "carrière et compétences",
    "événements orientation et insertion", "orientation et insertion archives",
}

def _url_http(resultat):
    url = str(resultat.get("url_finale") or resultat.get("url_reelle") or resultat.get("url") or "").strip()
    return url if url.startswith(("http://", "https://")) else ""

def _est_source_decouverte(resultat):
    texte = normaliser(" ".join(str(resultat.get(k, "")) for k in (
        "titre", "description", "titre_page", "description_page", "extrait_page"
    )))
    titre = normaliser(resultat.get("titre_page") or resultat.get("titre", ""))
    if not texte or not any(x in texte for x in MOTS_SOURCE_DECOUVERTE):
        return False
    signaux_event = [x for x in EVENT_SIGNAL_CIBLE if x in titre]
    preuve = any(x in texte for x in EVENT_PREUVE_FORTE)
    if signaux_event and (preuve or resultat.get("date_evenement_page") or resultat.get("lieu_page")):
        return False
    return True

def extraire_evenements_sources_decouverte(sources):
    """Utilise les pages institutionnelles comme index, pas comme candidats."""
    decouverte = [dict(r, type_source="SOURCE_DECOUVERTE", source_decouverte="OUI")
                  for r in sources if _est_source_decouverte(r) and _url_http(r)]
    if not decouverte:
        return [], []
    mots_event = set(EVENT_SIGNAL_CIBLE) | set(EVENT_PREUVE_FORTE) | {
        "career day", "career forum", "forum entreprises", "rencontres entreprises",
        "rencontre entreprises", "job dating", "forum alternance", "journée entreprises",
        "journée carrières", "semaine de l'emploi", "premier emploi", "alumni",
        "portes ouvertes", "salon étudiant", "forum étudiant", "2027", "2028",
    }
    def worker(parent):
        url = _url_http(parent)
        out = []
        try:
            response = requests.get(url, timeout=min(TIMEOUT, 12),
                                    headers={"User-Agent": "Mozilla/5.0 (compatible; ApecVeille/40)"})
            if response.status_code >= 400:
                return out
            soup = BeautifulSoup(response.text, "html.parser")
            for a in soup.find_all("a", href=True):
                label = normaliser(a.get_text(" ", strip=True))
                href = str(a.get("href", "")).strip()
                blob = normaliser(label + " " + href)
                if len(label) < 6 or not any(m in blob for m in mots_event):
                    continue
                if any(x in blob for x in EXCLUSIONS_FORTES):
                    continue
                lien = urljoin(response.url, href)
                if not lien.startswith(("http://", "https://")) or lien.rstrip("/") == url.rstrip("/"):
                    continue
                enfant = dict(parent)
                enfant.update({
                    "titre": label[:500], "titre_page": label[:500],
                    "url": lien, "url_reelle": lien, "url_finale": lien,
                    "origine_collecte": "decouverte",
                    "type_source": "EVENEMENT_DECOUVERT",
                    "source_decouverte_parent": url,
                    "description": (str(parent.get("description", "")) + " | Découvert depuis : " + url)[:5000],
                })
                out.append(enfant)
                if len(out) >= 30:
                    break
        except Exception:
            return []
        return out
    enfants=[]
    with ThreadPoolExecutor(max_workers=min(16, max(4, len(decouverte)))) as pool:
        futures=[pool.submit(worker,r) for r in decouverte]
        for f in futures:
            try: enfants.extend(f.result())
            except Exception: pass
    uniques={}
    for r in enfants:
        cle=canoniser_url(r.get("url")) or _cle_stable_candidat(r)
        if cle and cle not in uniques: uniques[cle]=r
    return decouverte, list(uniques.values())


def filtrer_candidats_avant_scoring(resultats):
    """V36 : filtre de sécurité à haute précision sans tuer le rappel.

    Principe : cette couche ne doit supprimer que les faux positifs évidents.
    Elle ne décide PAS à elle seule de la pertinence Apec.

    BLOQUANTS AVANT IA :
      1) géographie explicitement hors périmètre ;
      2) exclusions fortes manifestes (job étudiant, saisonnier, etc.) ;
      3) contenu manifestement non événementiel ;
      4) événement explicitement passé.

    IMPORTANT : l'absence de signal Apec explicite n'est PLUS un motif de
    rejet. Un événement réel mais ambigu est conservé comme candidat et sera
    évalué par le scoring puis par les 3 tours Gemini 3.1.
    """
    acceptes = []
    rejetes = []
    motifs = {}

    cible_directe = set(APEC_DIRECT)
    cible_qualifiee = set(APEC_QUALIFIE)
    relation = set(RELATION_ENTREPRISE)
    evenements = set(EVENT_SIGNAL_CIBLE)
    structures = set(EVENT_MOTS_STRUCTURELS)
    exclusions = set(EXCLUSIONS_FORTES)
    editoriaux = set(EXCLUSIONS_EDITORIAL)

    for resultat in resultats:
        texte = normaliser(" ".join(str(resultat.get(k, "")) for k in (
            "titre", "description", "titre_page", "description_page",
            "extrait_page", "lieu_page", "organisateur_page",
            "date_evenement_page"
        )))
        titre = normaliser(resultat.get("titre_page") or resultat.get("titre", ""))

        # ============================================================
        # 1. GÉOGRAPHIE : SEUL UN HORS-PÉRIMÈTRE EXPLICITE BLOQUE
        # ============================================================
        corriger_departement_par_ville(resultat)
        geo_ok_fort, geo_motif_fort, geo_code_fort = verifier_geographie_inviolable(resultat)
        if not geo_ok_fort:
            resultat.update({
                "pre_filtre": "REJETE",
                "pre_filtre_motif": "HORS_GEOGRAPHIE",
                "erreur_geographique": "ERREUR_GEOGRAPHIQUE",
                "hors_geographie": geo_motif_fort,
                "score": 0,
                "statut": "Non pertinent",
                "statut_final": "Non pertinent",
                "geo_controle_fort": "BLOQUE",
            })
            rejetes.append(resultat)
            motifs["HORS_GEOGRAPHIE"] = motifs.get("HORS_GEOGRAPHIE", 0) + 1
            continue

        geo_ok, geo_motif, geo_code = verifier_geographie_stricte(resultat)
        if not geo_ok:
            resultat.update({
                "pre_filtre": "REJETE",
                "pre_filtre_motif": "HORS_GEOGRAPHIE",
                "erreur_geographique": "ERREUR_GEOGRAPHIQUE",
                "hors_geographie": geo_motif,
                "score": 0,
                "statut": "Non pertinent",
                "statut_final": "Non pertinent",
            })
            rejetes.append(resultat)
            motifs["HORS_GEOGRAPHIE"] = motifs.get("HORS_GEOGRAPHIE", 0) + 1
            continue

        resultat["pre_filtre_geo"] = "OK"
        if geo_code:
            resultat["departement"] = str(geo_code)
            resultat["nom_departement"] = NOMS_DEPARTEMENTS.get(
                str(geo_code), resultat.get("nom_departement", "")
            )

        # ============================================================
        # 2. EXCLUSIONS FORTES : BLOCAGE IMMÉDIAT
        # ============================================================
        signaux_exclusion = [x for x in exclusions if x in texte]

        # Contextes manifestement hors veille Apec, même si la page utilise
        # un mot événementiel générique comme « journée » ou « atelier ».
        signaux_hors_apec = [
            x for x in (
                "crèche", "petite enfance", "bal des pompiers", "festival",
                "concert", "spectacle", "tourisme", "destination france à vélo",
                "lauréat", "lauréate", "prix touristique", "emploi-retraite",
                "information retraite", "droits à la retraite"
            ) if x in texte
        ]
        direct_ctx = [x for x in cible_directe if x in texte]
        qual_ctx = [x for x in cible_qualifiee if x in texte]
        rel_ctx = [x for x in relation if x in texte]
        sup_ctx = [x for x in SUPERIEUR_FORT if x in texte]
        if signaux_hors_apec and not (direct_ctx or qual_ctx or rel_ctx or sup_ctx):
            resultat.update({
                "pre_filtre": "REJETE",
                "pre_filtre_motif": "CONTEXTE_HORS_APEC",
                "motif_exclusion": ", ".join(signaux_hors_apec[:8]),
                "score": 0,
                "statut": "Non pertinent",
                "statut_final": "Non pertinent",
            })
            rejetes.append(resultat)
            motifs["CONTEXTE_HORS_APEC"] = motifs.get("CONTEXTE_HORS_APEC", 0) + 1
            continue

        if signaux_exclusion:
            resultat.update({
                "pre_filtre": "REJETE",
                "pre_filtre_motif": "EXCLUSION_FORTE",
                "motif_exclusion": ", ".join(signaux_exclusion[:8]),
                "score": 0,
                "statut": "Non pertinent",
                "statut_final": "Non pertinent",
            })
            rejetes.append(resultat)
            motifs["EXCLUSION_FORTE"] = motifs.get("EXCLUSION_FORTE", 0) + 1
            continue

        # ============================================================
        # 3. ÉVÉNEMENT : EXIGER UN SIGNAL ÉVÉNEMENTIEL, MAIS PAS UN
        #    SIGNAL APEC. C'EST LE POINT CENTRAL DE LA V36.
        # ============================================================
        signaux_event = [x for x in evenements if x in texte]
        signaux_structure = [x for x in structures if x in texte]
        preuve = [x for x in EVENT_PREUVE_FORTE if x in texte]
        event_titre = [x for x in evenements if x in titre]

        event_reel = bool(
            signaux_event
            or event_titre
            or (signaux_structure and (preuve or resultat.get("date_evenement_page")
                                       or resultat.get("lieu_page")
                                       or resultat.get("organisateur_page")))
        )

        if not event_reel:
            resultat.update({
                "pre_filtre": "REJETE",
                "pre_filtre_motif": "PAS_UN_EVENEMENT",
                "score": 0,
                "statut": "Non pertinent",
                "statut_final": "Non pertinent",
            })
            rejetes.append(resultat)
            motifs["PAS_UN_EVENEMENT"] = motifs.get("PAS_UN_EVENEMENT", 0) + 1
            continue

        # ============================================================
        # 4. SIGNAL APEC : DEVENU UN NIVEAU DE CONFIANCE, PAS UN FILTRE
        #    DUR. L'IA doit pouvoir arbitrer les cas ambigus.
        # ============================================================
        direct = [x for x in cible_directe if x in texte]
        qualifie = [x for x in cible_qualifiee if x in texte]
        rel = [x for x in relation if x in texte]
        superieur = [x for x in SUPERIEUR_FORT if x in texte]

        if direct or rel:
            niveau_cible = "FORT"
        elif qualifie or superieur:
            niveau_cible = "MOYEN"
        else:
            niveau_cible = "A_VERIFIER"

        resultat["cible_apec_pre"] = "OUI" if (direct or rel or qualifie or superieur) else "A_VERIFIER"
        resultat["niveau_cible_apec_pre"] = niveau_cible
        resultat["signaux_cible_apec_pre"] = ", ".join(
            (direct + qualifie[:8] + rel[:8] + superieur[:8])[:20]
        )

        # ============================================================
        # 5. ÉDITORIAL : ON NE REJETTE QUE SI LE TEXTE EST ÉDITORIAL
        #    ET NE CONTIENT AUCUNE PREUVE ÉVÉNEMENTIELLE.
        # ============================================================
        editorial = [x for x in editoriaux if x in texte]
        if editorial and not (event_titre or preuve or resultat.get("date_evenement_page")
                              or resultat.get("lieu_page")):
            resultat.update({
                "pre_filtre": "REJETE",
                "pre_filtre_motif": "ARTICLE_EDITORIAL",
                "score": 0,
                "statut": "Non pertinent",
                "statut_final": "Non pertinent",
            })
            rejetes.append(resultat)
            motifs["ARTICLE_EDITORIAL"] = motifs.get("ARTICLE_EDITORIAL", 0) + 1
            continue

        # ============================================================
        # 6. DATE : UN ÉVÉNEMENT DONT LA DATE EST EXPLICITEMENT PASSÉE
        #    EST ÉCARTÉ. Une simple année 2026 dans un titre ne suffit pas.
        # ============================================================
        date_evt = parser_date_evenement(resultat.get("date_evenement_page", ""))
        if date_evt and date_evt.date() < datetime.now().date():
            resultat.update({
                "pre_filtre": "REJETE",
                "pre_filtre_motif": "EVENEMENT_PASSE",
                "score": 0,
                "statut": "Non pertinent",
                "statut_final": "Non pertinent",
            })
            rejetes.append(resultat)
            motifs["EVENEMENT_PASSE"] = motifs.get("EVENEMENT_PASSE", 0) + 1
            continue

        # ============================================================
        # 7. CONSERVATION : TOUT ÉVÉNEMENT PLAUSIBLE ENTRE DANS LE
        #    SCORING. Les cas ambigus seront "À vérifier" / Gemini.
        # ============================================================
        resultat.update({
            "pre_filtre": "ACCEPTE",
            "pre_filtre_motif": "CANDIDAT_EVENEMENT",
            "preuve_evenement_pre": "OUI" if (
                preuve or event_titre or resultat.get("date_evenement_page")
                or resultat.get("lieu_page")
            ) else "A_VERIFIER",
            "signaux_evenement_pre": ", ".join(signaux_event[:12]),
        })
        acceptes.append(resultat)

    print()
    print("🛡️ FILTRE V36 — HAUTE PRÉCISION / RAPPEL PRÉSERVÉ")
    print(f"Entrées : {len(resultats)}")
    print(f"Candidats conservés : {len(acceptes)}")
    print(f"Rejetés avant scoring/IA : {len(rejetes)}")
    if motifs:
        print("Motifs : " + " | ".join(f"{k}={v}" for k, v in sorted(motifs.items())))

    return acceptes, rejetes, motifs


def traiter(resultats, memoire_globale=None):
    resultats = dedoublonner(resultats)
    memoire_globale = memoire_globale or {}

    pertinents = []
    a_verifier = []
    non_pertinents = []
    ecartes_hors_periode = 0
    scores_reutilises = 0
    scores_recalcules = 0

    champs_score = [
        "score", "score_cible_apec", "score_emploi", "score_formation",
        "score_superieur", "score_apec", "score_entreprise", "score_qualifie",
        "score_geographie", "score_date", "score_bruit", "score_complementaire",
        "bonus_titre", "signaux_cible_apec", "signaux_superieur",
        "signaux_entreprise", "signaux_exclusion", "statut", "raison",
        "motif", "details_score", "evenement_detecte", "preuve_evenement",
        "geographie_verifiee", "annee_detectee", "hors_periode",
        "hors_perimetre", "hors_geographie", "niveau_source_v9",
        "priorite_2027", "periode",
    ]

    for resultat in resultats:
        corriger_departement_par_ville(resultat)

        # FILTRE GÉOGRAPHIQUE ABSOLU : aucun événement explicitement hors
        # des 16/17/19/23/79/86/87 ne doit atteindre le scoring ou Gemini.
        geo_stricte_ok, geo_motif, geo_code = verifier_geographie_stricte(resultat)
        if not geo_stricte_ok:
            resultat["hors_perimetre"] = "OUI"
            resultat["hors_geographie"] = geo_motif
            resultat["erreur_geographique"] = "ERREUR_GEOGRAPHIQUE"
            resultat["score"] = 0
            resultat["score_geographie"] = 0
            resultat["statut"] = "Non pertinent"
            resultat["statut_final"] = "Non pertinent"
            resultat["decision_ia"] = "IA_NON_REQUISE"
            resultat["raison_ia"] = "Exclusion géographique absolue : " + geo_motif
            non_pertinents.append(resultat)
            ecartes_hors_periode += 0
            continue

        resultat["geographie_stricte_motif"] = geo_motif
        resultat["departement"] = str(resultat.get("departement") or geo_code)
        resultat["nom_departement"] = NOMS_DEPARTEMENTS.get(
            resultat["departement"], resultat.get("nom_departement", "")
        )

        cle = _cle_memoire_globale(resultat)
        item = memoire_globale.get(cle) if cle else None
        reutiliser_score = False
        if item and item.get("signature") == _signature_memoire(resultat):
            try:
                date_mem = datetime.fromisoformat(
                    str(item.get("date_memorisation", "")).replace("Z", "")
                )
                age = (datetime.now() - date_mem).days
            except Exception:
                age = MEMORY_REFRESH_DAYS + 1
            if age < MEMORY_REFRESH_DAYS:
                champs_mem = item.get("champs") or {}
                # On ne réutilise le score que si le cache contient bien le
                # score principal et le statut. Cela protège les anciennes
                # mémoires créées avant cette évolution.
                if "score" in champs_mem and "statut" in champs_mem:
                    for champ in champs_score:
                        valeur = champs_mem.get(champ)
                        if valeur not in ("", None):
                            resultat[champ] = valeur
                    # Compatibilité avec les anciennes mémoires : aucun score
                    # négatif ne doit ressortir d'un cache historique.
                    try:
                        resultat["score"] = max(0, int(float(resultat.get("score", 0) or 0)))
                    except Exception:
                        resultat["score"] = 0
                    try:
                        resultat["score_geographie"] = max(0, int(float(resultat.get("score_geographie", 0) or 0)))
                    except Exception:
                        resultat["score_geographie"] = 0
                    resultat["score_bruit"] = max(0, int(float(resultat.get("score_bruit", 0) or 0)))
                    resultat["memoire_scoring"] = "OUI"
                    reutiliser_score = True
                    scores_reutilises += 1

        if not reutiliser_score:
            score_data = calculer_score(resultat)
            resultat.update(score_data)
            resultat["memoire_scoring"] = "NON"
            scores_recalcules += 1

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
    print(f"🧮 Mémoire scoring : {scores_reutilises} scores réutilisés / {len(resultats)}")
    print(f"🆕 Scorings Python réellement calculés : {scores_recalcules}")

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

def arbitrer_coherence_decision(resultat, decision):
    """Sécurise les contradictions manifestes entre justification et décision IA."""
    d = str(decision or "A_VERIFIER").strip().upper()
    if d not in {"PERTINENT", "A_VERIFIER", "NON_PERTINENT"}:
        return "A_VERIFIER", "DECISION_IA_INVALIDE"
    texte = " ".join(str(resultat.get(k, "") or "") for k in (
        "raison_ia", "analyse_humaine", "action_recommandee",
        "raison_ia_1", "raison_ia_31_t2", "raison_ia_31_t3",
        "raison_ia_38", "preuve_decisive_38", "apport_38",
    )).lower()
    negatifs = [x for x in [
        "en décalage avec le positionnement de l'apec",
        "en decalage avec le positionnement de l'apec",
        "pas adapté à l'apec", "pas adapte a l'apec",
        "non adapté à l'apec", "non adapte a l'apec",
        "ne correspond pas au positionnement de l'apec",
        "ne correspond pas à l'apec", "ne correspond pas a l'apec",
        "pas de recrutement de cadres", "pas de recrutement cadre",
        "orientation initiale", "formation initiale",
        "lycéens/parents", "lyceens/parents",
        "lycéens et parents", "lyceens et parents",
        "public lycéen", "public lyceen",
        "métiers de l'éducation nationale", "metiers de l'education nationale",
        "préparation aux concours", "preparation aux concours",
    ] if x in texte]
    if d == "PERTINENT" and negatifs:
        return "NON_PERTINENT", "ARBITRAGE_COHERENCE_NEGATIF:" + negatifs[0]
    positifs = [x for x in [
        "recrutement de cadres", "recrutement cadre", "jeunes diplômés",
        "jeunes diplomes", "bac+5", "bac +5", "bac+3", "bac +3",
        "relations entreprises", "forum entreprises", "rencontre entreprises",
        "career day", "job dating", "recruteurs", "employeurs",
    ] if x in texte]
    if d == "NON_PERTINENT" and len(positifs) >= 2 and not negatifs:
        return "A_VERIFIER", "ARBITRAGE_COHERENCE_POSITIF"
    return d, ""



def _domaine_url(url):
    try:
        return urlparse(str(url or "").strip()).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def _preuve_secondaire_google_news(candidat):
    """Recherche une seconde source indépendante pour le même événement.

    La recherche est volontairement étroite : titre + lieu + année. Elle ne
    sert pas à créer un candidat ; elle sert uniquement à confirmer un
    événement déjà passé le filtre V42. La source primaire est exclue lorsque
    son domaine est identifiable.
    """
    titre = normaliser(candidat.get("titre_page") or candidat.get("titre", ""))
    if len(titre) < 10:
        return []
    lieu = normaliser(candidat.get("lieu_page") or candidat.get("lieu", ""))
    annee = str(candidat.get("annee_detectee") or candidat.get("annee_recherchee") or ANNEE_PRIORITAIRE)
    domaine_primaire = _domaine_url(candidat.get("url_finale") or candidat.get("url"))
    morceaux = [f'"{titre[:180]}"']
    if lieu and lieu not in titre:
        morceaux.append(f'"{lieu[:80]}"')
    if annee:
        morceaux.append(annee)
    if domaine_primaire:
        morceaux.append(f'-site:{domaine_primaire}')
    q = " ".join(morceaux)
    url = f"https://news.google.com/rss/search?q={quote_plus(q)}&hl=fr&gl=FR&ceid=FR:fr"
    try:
        response = requests.get(
            url, timeout=min(TIMEOUT, 12),
            headers={"User-Agent": "Mozilla/5.0 (compatible; ApecVeille/42)"}
        )
        if response.status_code >= 400:
            return []
        feed = feedparser.parse(response.content)
        out=[]
        titre_cle = cle_titre_evenement(titre)
        for entry in getattr(feed, "entries", [])[:8]:
            et = normaliser(str(entry.get("title", "")))
            if not et:
                continue
            ratio = difflib.SequenceMatcher(None, titre_cle, cle_titre_evenement(et)).ratio()
            if ratio < 0.45:
                continue
            link = str(entry.get("link", "")).strip()
            domaine = _domaine_url(link)
            if domaine and domaine == domaine_primaire:
                continue
            out.append({
                "titre": str(entry.get("title", ""))[:400],
                "url": link,
                "source": domaine or "Google News",
                "similarite": round(ratio, 3),
            })
        return out[:3]
    except Exception:
        return []


def verifier_preuves_multisources(candidats):
    """V42 : transforme les informations disponibles en niveau de preuve.

    La preuve n'est jamais un malus et ne décide pas seule de la pertinence.
    Elle sert à distinguer :
      4 = preuve explicite + source officielle + seconde source indépendante
      3 = preuve explicite sur une source fiable
      2 = plusieurs indices cohérents, sans confirmation indépendante
      1 = signal faible
      0 = aucune preuve exploitable
    """
    if not candidats:
        return candidats

    def worker(c):
        titre = normaliser(c.get("titre_page") or c.get("titre", ""))
        desc = normaliser(c.get("description_page") or c.get("description", ""))
        date_evt = str(c.get("date_evenement_page") or "").strip()
        lieu_evt = normaliser(c.get("lieu_page") or c.get("adresse_page") or c.get("lieu", ""))
        org = normaliser(c.get("organisateur_page") or "")
        url = c.get("url_finale") or c.get("url") or ""
        source = str(c.get("source", ""))
        texte = " ".join(x for x in (titre, desc, date_evt, lieu_evt, org) if x)

        signaux = []
        if titre and any(x in titre for x in EVENT_SIGNAL_CIBLE):
            signaux.append("INTITULE_EVENEMENT")
        if any(x in texte for x in EVENT_PREUVE_FORTE):
            signaux.append("FORMULATION_EVENEMENTIELLE")
        if date_evt:
            signaux.append("DATE_EXPLICITE")
        if lieu_evt:
            signaux.append("LIEU_EXPLICITE")
        if org:
            signaux.append("ORGANISATEUR_EXPLICITE")
        if c.get("preuve_evenement") == "OUI" or c.get("evenement_detecte") == "OUI":
            signaux.append("PREUVE_EXISTANTE")

        niveau_source = classer_fiabilite_source(c)
        secondaire = []
        # On ne dépense une recherche secondaire que si le candidat est
        # suffisamment structuré pour que la vérification soit utile.
        if len(signaux) >= 2 and (date_evt or lieu_evt):
            secondaire = _preuve_secondaire_google_news(c)

        if secondaire:
            signaux.append("SOURCE_SECONDAIRE")

        explicite = bool(date_evt and lieu_evt and (org or c.get("preuve_evenement") == "OUI"))
        if explicite and niveau_source <= 2 and secondaire:
            niveau = 4
            libelle = "FORTE"
        elif explicite and niveau_source <= 2:
            niveau = 3
            libelle = "FORTE"
        elif len(set(signaux)) >= 3:
            niveau = 2
            libelle = "MOYENNE"
        elif len(set(signaux)) >= 1:
            niveau = 1
            libelle = "FAIBLE"
        else:
            niveau = 0
            libelle = "ABSENTE"

        c["niveau_preuve"] = niveau
        c["preuve_cible"] = (
            "FORTE" if any(x in texte for x in APEC_DIRECT + RELATION_ENTREPRISE)
            else "MOYENNE" if any(x in texte for x in APEC_QUALIFIE + SUPERIEUR_FORT)
            else "FAIBLE"
        )
        c["preuve_evenement_v42"] = libelle
        c["preuves_v42"] = " | ".join(dict.fromkeys(signaux))
        c["source_secondaire_v42"] = " | ".join(
            f"{x.get('source','')}::{x.get('titre','')}::{x.get('url','')}" for x in secondaire
        )
        c["nombre_sources_secondaires_v42"] = len(secondaire)
        c["preuve_verifiee_v42"] = "OUI" if niveau >= 3 else "NON"
        c["preuve_source_primaire"] = source
        return c

    with ThreadPoolExecutor(max_workers=min(8, MAX_WORKERS)) as executor:
        futures=[executor.submit(worker,c) for c in candidats]
        for f in as_completed(futures):
            try:
                f.result()
            except Exception:
                pass
    return candidats

def appliquer_ia(pertinents, a_verifier, memoire_ia=None, bdd_candidats=None):
    """Cascade conservatrice : 3 tours Gemini 3.1, puis TOP 3 en 3.8.

    La BDD/mémoire reste prioritaire : les candidats inchangés déjà analysés
    ne repassent pas par Gemini. En mode test, le volume peut être limité et
    la persistance désactivée.
    """
    memoire_ia = memoire_ia or {}
    bdd_candidats = bdd_candidats or preparer_bdd_candidats()
    tous_candidats = pertinents + a_verifier
    candidats, non_eligibles, hors_perimetre = [], [], []

    for candidat in tous_candidats:
        score = max(0, int(candidat.get("score", 0) or 0))
        candidat["score"] = score
        dept = str(candidat.get("departement", "")).strip()
        erreur_geo = str(candidat.get("erreur_geographique", "")).strip()
        if dept and dept not in DEPARTEMENTS_CIBLES:
            candidat.update({"erreur_geographique":"ERREUR_GEOGRAPHIQUE","decision_ia":"IA_NON_REQUISE","confiance_ia":0,"raison_ia":f"Hors périmètre géographique : département {dept}.","ia_analyse":"NON","statut_final":"Non pertinent"})
            hors_perimetre.append(candidat); continue
        if erreur_geo == "ERREUR_GEOGRAPHIQUE":
            candidat.update({"decision_ia":"IA_NON_REQUISE","confiance_ia":0,"raison_ia":"Hors périmètre géographique : exclu avant analyse IA.","ia_analyse":"NON","statut_final":"Non pertinent"})
            hors_perimetre.append(candidat); continue
        # V36 : après le filtre événementiel dur, tout candidat encore éligible
        # est envoyé à Gemini 3.1. Le score Python ne sert plus de barrière
        # métier : il reste un signal explicable pour l'arbitrage.
        candidats.append(candidat)

    # ------------------------------------------------------------
    # PRÉ-ANALYSE GÉOGRAPHIQUE IA
    # ------------------------------------------------------------
    # Étape locale, gratuite et conservatrice : elle inspecte le lieu de
    # l'événement avant toute consommation Gemini. Une localisation
    # explicitement hors 16/17/19/23/79/86/87 est pré-exclue.
    geo_preexclus = []
    candidats_geo_ok = []
    for candidat in candidats:
        geo = preanalyser_geographie(candidat)
        candidat.update(geo)
        candidat["ia_preanalyse"] = "OUI"
        if geo.get("geo_preanalyse") == "GEO_HORS_CIBLE":
            candidat.update({
                "ia_preexclusion_geo": "OUI",
                "decision_ia": "NON_PERTINENT",
                "confiance_ia": int(geo.get("geo_confiance", 0) or 0),
                "raison_ia": geo.get("geo_motif", "Hors périmètre géographique."),
                "motif_exclusion": "HORS_GEOGRAPHIE",
                "ia_analyse": "GEO_PREEXCLUSION",
                "statut_final": "Non pertinent",
            })
            geo_preexclus.append(candidat)
        else:
            candidat["ia_preexclusion_geo"] = "NON"
            candidats_geo_ok.append(candidat)

    # Même les exclusions géographiques sont identifiées par la BDD :
    # une exclusion déjà connue ne ressort pas comme « nouvelle » chaque jour.
    geo_nouveaux_bdd, geo_modifies_bdd, geo_reutilises_bdd = classifier_candidats_bdd(geo_preexclus, bdd_candidats) if geo_preexclus else ([], [], [])
    geo_preexclus = geo_nouveaux_bdd + geo_modifies_bdd

    candidats = sorted(candidats_geo_ok, key=lambda x: (str(ANNEE_PRIORITAIRE) in str(x.get("annee_detectee", "")), x.get("score", 0)), reverse=True)
    nouveaux_bdd, modifies_bdd, reutilises_bdd = classifier_candidats_bdd(candidats, bdd_candidats)
    a_analyser_bdd = nouveaux_bdd + modifies_bdd
    # V41 : la mémoire globale est un cache de scoring, pas une mémoire IA.
    # Un candidat NOUVEAU ou MODIFIÉ selon la BDD doit impérativement
    # continuer vers la couche IA. Le scoring global est déjà géré dans
    # traiter(); il ne doit donc pas retirer de candidats du flux IA.
    memoire_globale = charger_memoire_globale()
    a_verifier_memoire = a_analyser_bdd
    a_traiter, memorises = appliquer_memoire_ia(a_verifier_memoire, memoire_ia)

    # ------------------------------------------------------------
    # SECOND VERROU GÉOGRAPHIQUE — APRÈS ENRICHISSEMENT DE PAGE
    # ------------------------------------------------------------
    # La première pré-analyse est volontairement rapide. Ici, on attend
    # d'avoir récupéré lieu_page/adresse_page/titre_page/contenu_page avant
    # de consommer le moindre appel Gemini. Cela corrige notamment les cas
    # du type « Campus Nantes » qui ne sont pas présents dans le RSS brut.
    enrichir_candidats_ia(a_traiter)
    geo_preexclus_apres_enrichissement = []
    a_traiter_geo_ok = []
    for c in a_traiter:
        geo2 = preanalyser_geographie(c)
        c.update(geo2)
        geo_fort_ok, geo_fort_motif, geo_fort_code = verifier_geographie_inviolable(c)
        c["ia_preanalyse"] = "OUI"
        # V40 : la pré-analyse géographique IA est informative. Le seul verrou
        # dur ici est la vérification géographique indépendante et explicite.
        if not geo_fort_ok:
            c.update({
                "ia_preexclusion_geo": "OUI",
                "decision_ia": "NON_PERTINENT",
                "confiance_ia": int(geo2.get("geo_confiance", 0) or 0),
                "raison_ia": "Pré-exclusion géographique après enrichissement : " + (
                    geo_fort_motif if not geo_fort_ok else geo2.get("geo_motif", "hors périmètre")
                ),
                "motif_exclusion": "HORS_GEOGRAPHIE",
                "ia_analyse": "GEO_PREEXCLUSION",
                "statut_final": "Non pertinent",
                "geo_controle_final": "BLOQUE_AVANT_GEMINI",
            })
            geo_preexclus_apres_enrichissement.append(c)
        else:
            c["ia_preexclusion_geo"] = "NON"
            a_traiter_geo_ok.append(c)
    a_traiter = a_traiter_geo_ok
    geo_preexclus.extend(geo_preexclus_apres_enrichissement)

    for c in a_traiter:
        try: c.update(calculer_score(c))
        except Exception: pass

    # ------------------------------------------------------------
    # V42 — PREUVES MULTI-SOURCES
    # ------------------------------------------------------------
    # Cette couche intervient après le contrôle géographique et avant Gemini.
    # Elle n'élimine aucun candidat : elle documente la solidité de la preuve.
    if a_traiter:
        print(f"🔎 V42 — vérification des preuves : {len(a_traiter)} candidats")
        verifier_preuves_multisources(a_traiter)

    if IA_TEST_MODE and IA_TEST_MAX_CANDIDATS > 0:
        a_traiter = a_traiter[:IA_TEST_MAX_CANDIDATS]

    preuves = [r for r in a_traiter if r.get("niveau_preuve") is not None]
    AUDIT_PIPELINE["preuves_verifiees"] = len(preuves)
    AUDIT_PIPELINE["preuves_fortes"] = sum(1 for r in preuves if r.get("preuve_evenement_v42") == "FORTE")
    AUDIT_PIPELINE["preuves_moyennes"] = sum(1 for r in preuves if r.get("preuve_evenement_v42") == "MOYENNE")
    AUDIT_PIPELINE["preuves_faibles"] = sum(1 for r in preuves if r.get("preuve_evenement_v42") == "FAIBLE")
    AUDIT_PIPELINE["preuves_absentes"] = sum(1 for r in preuves if r.get("preuve_evenement_v42") == "ABSENTE")
    AUDIT_PIPELINE["sources_secondaires_trouvees"] = sum(int(r.get("nombre_sources_secondaires_v42", 0) or 0) for r in preuves)

    ia_resultats = list(geo_preexclus) + list(geo_reutilises_bdd) + list(reutilises_bdd) + list(memorises)
    quota_bloque = False
    erreurs_consecutives = 0
    appels_31 = appels_38 = 0

    print("\n========================================")
    print("IA V31 — 3 TOURS 3.1 + TOP 3 EN 3.8")
    print("========================================")
    print(f"Règles : {IA_RULES_VERSION}")
    print(f"Candidats filtrés : {len(candidats) + len(geo_preexclus)}")
    print(f"Pré-exclusions géographiques IA : {len(geo_preexclus)} nouvelles/modifiées | {len(geo_reutilises_bdd)} anciennes masquées")
    print(f"BDD : {len(reutilises_bdd)} réutilisés / {len(nouveaux_bdd)} nouveaux / {len(modifies_bdd)} modifiés")
    print(f"Mémoire IA : {len(memorises)} réutilisés")
    print(f"À analyser : {len(a_traiter)}" + (f" | MODE TEST, max {IA_TEST_MAX_CANDIDATS}" if IA_TEST_MODE else ""))
    print(f"3.1 : 3 tours, cooldown {COOLDOWN_31_SECONDS}s | 3.8 : TOP {MAX_IA_38}, cooldown {COOLDOWN_38_SECONDS}s")

    def appel_avec_cooldown(fn, args, cooldown, dernier):
        if dernier:
            attente = cooldown - (time.time() - dernier)
            if attente > 0: time.sleep(attente)
        resultat = fn(*args)
        return resultat, time.time()

    # PHASE 1 — 3.1 TOUR 1
    t1_ok = []
    dernier = 0.0
    for i, r in enumerate(a_traiter, 1):
        print(f"[PHASE 1/4 — 3.1 T1] {i}/{len(a_traiter)} — {court(r.get('titre',''))}", flush=True)
        analyse, dernier = appel_avec_cooldown(analyser_resultat_31, (r,), COOLDOWN_31_SECONDS, dernier)
        appels_31 += 1
        r.update(analyse)
        r["ia_analyse"] = "OUI"; r["memoire_ia"] = "NON"
        r["decision_ia_1"] = analyse.get("decision_ia_1", "IA_INDISPONIBLE")
        r["confiance_ia_1"] = analyse.get("confiance_ia_1", 0)
        r["raison_ia_1"] = analyse.get("raison_ia_1", "")
        if r["decision_ia_1"] == "IA_INDISPONIBLE":
            erreurs_consecutives += 1
        else:
            erreurs_consecutives = 0
        if erreurs_consecutives >= MAX_ERREURS_IA_CONSECUTIVES:
            quota_bloque = True; break
        t1_ok.append(r)

    # PHASE 2 — 3.1 TOUR 2 : contre-analyse
    t2_ok = []
    dernier = 0.0
    for i, r in enumerate(t1_ok, 1):
        print(f"[PHASE 2/4 — 3.1 T2] {i}/{len(t1_ok)} — {court(r.get('titre',''))}", flush=True)
        t1 = {k:v for k,v in r.items() if k.endswith('_31') or k in {"decision_ia_1","confiance_ia_1","raison_ia_1","preuves_31","public_cible_31","type_evenement_31"}}
        analyse, dernier = appel_avec_cooldown(analyser_resultat_31_tour2, (r, t1), COOLDOWN_31_SECONDS, dernier)
        appels_31 += 1; r.update(analyse)
        if analyse.get("decision_ia_31_t2") == "IA_INDISPONIBLE":
            r["raison_ia_31_t2"] = "Tour 2 indisponible : Tour 1 conservé."
        t2_ok.append(r)

    # PHASE 3 — 3.1 TOUR 3 : tranche
    t3_ok = []
    dernier = 0.0
    for i, r in enumerate(t2_ok, 1):
        print(f"[PHASE 3/4 — 3.1 T3] {i}/{len(t2_ok)} — {court(r.get('titre',''))}", flush=True)
        t1 = {k:v for k,v in r.items() if k.endswith('_31') or k in {"decision_ia_1","confiance_ia_1","raison_ia_1","preuves_31","public_cible_31"}}
        t2 = {k:v for k,v in r.items() if '31_t2' in k}
        analyse, dernier = appel_avec_cooldown(analyser_resultat_31_tour3, (r, t1, t2), COOLDOWN_31_SECONDS, dernier)
        appels_31 += 1; r.update(analyse)
        decision = analyse.get("decision_ia_31_t3", "A_VERIFIER")
        if decision == "IA_INDISPONIBLE":
            decision = r.get("decision_ia_1", "A_VERIFIER")
            r["raison_ia_31_t3"] = "Tour 3 indisponible : décision Tour 1 conservée."
        r["decision_ia"] = decision
        r["confiance_ia"] = analyse.get("confiance_ia_31_t3", r.get("confiance_ia_1", 0))
        r["raison_ia"] = analyse.get("raison_ia_31_t3", r.get("raison_ia_1", ""))
        r["modele_ia"] = MODEL_31 if 'MODEL_31' in globals() else "gemini-3.1-flash-lite"
        r["analyse_humaine"] = r.get("raison_ia_31_t3", "")
        decision_coherente, motif_coherence = arbitrer_coherence_decision(r, r["decision_ia"])
        r["decision_ia_brute_31_t3"] = r["decision_ia"]
        r["decision_ia"] = decision_coherente
        r["arbitrage_coherence"] = motif_coherence
        classer_statut_final(r, r["decision_ia"])
        t3_ok.append(r)

    # Sélection intelligente : pertinence + incertitude + désaccord, pas score Python seul.
    def priorite_38(r):
        d1 = r.get("decision_ia_1", ""); d2 = r.get("decision_ia_31_t2", ""); d3 = r.get("decision_ia_31_t3", r.get("decision_ia", ""))
        c3 = int(r.get("confiance_ia_31_t3", 0) or 0)
        py = int(r.get("score", 0) or 0)
        conflit = 35 if len({d1,d2,d3}) > 1 else (15 if d1 != d3 or d2 != d3 else 0)
        incert = max(0, 100-c3)
        fort = 20 if d3 == "PERTINENT" else (12 if d3 == "A_VERIFIER" else 0)
        p2027 = 15 if "FORTE" in str(r.get("priorite_2027", "")).upper() else 0
        return (conflit + incert + fort + p2027, py, c3)

    top_38 = sorted([r for r in t3_ok if r.get("ia_analyse") == "OUI"], key=priorite_38, reverse=True)[:MAX_IA_38]

    # PHASE 4 — 3.8 : expertise senior sur TOP 3
    dernier = 0.0
    for i, r in enumerate(top_38, 1):
        print(f"[PHASE 4/4 — 3.8] {i}/{len(top_38)} — {court(r.get('titre',''))}", flush=True)
        analyse38, dernier = appel_avec_cooldown(analyser_resultat_38, (r, r), COOLDOWN_38_SECONDS, dernier)
        appels_38 += 1; r.update(analyse38); r["escalade_38"] = "OUI"; r["motif_escalade_38"] = "Top candidat après 3 tours Gemini 3.1"
        if analyse38.get("decision_ia_38") != "IA_INDISPONIBLE":
            # 3.8 est un arbitre senior, pas un remplacement aveugle de 3.1.
            # La décision T3 reste la base ; l'arbitrage 3.8 détermine comment
            # elle évolue. La sortie JSON de 3.8 est conservée pour audit.
            decision_t3 = r.get("decision_ia_31_t3", r.get("decision_ia", "A_VERIFIER"))
            arbitrage = str(analyse38.get("arbitrage_38", "AUCUN")).upper()
            if arbitrage == "CONFIRMER":
                decision_finale = decision_t3
            elif arbitrage == "RETROGRADER":
                decision_finale = "A_VERIFIER" if decision_t3 == "PERTINENT" else decision_t3
            elif arbitrage == "REJETER":
                decision_finale = "NON_PERTINENT"
            else:
                # Si l'arbitrage est absent/incohérent, on ne laisse pas 3.8
                # écraser T3 par sa seule décision textuelle.
                decision_finale = decision_t3
            r["decision_ia_31_finale"] = decision_t3
            r["arbitrage_38"] = arbitrage
            r["decision_ia_38_brute"] = analyse38.get("decision_ia_38", "")
            r["decision_ia"] = decision_finale
            decision_coherente, motif_coherence = arbitrer_coherence_decision(r, r["decision_ia"])
            r["decision_ia_brute_38_arbitrage"] = r["decision_ia"]
            r["decision_ia"] = decision_coherente
            if motif_coherence:
                r["arbitrage_coherence"] = motif_coherence
            r["confiance_ia"] = analyse38.get("confiance_ia_38", r.get("confiance_ia", 0))
            r["raison_ia"] = analyse38.get("raison_ia_38", r.get("raison_ia", ""))
            r["analyse_humaine"] = analyse38.get("analyse_humaine_38", r.get("analyse_humaine", ""))
            r["action_recommandee"] = analyse38.get("action_recommandee_38", r.get("action_recommandee_31_t3", ""))
            r["preuve_decisive_38"] = analyse38.get("preuve_decisive_38", "")
            r["apport_38"] = analyse38.get("apport_38", "")
            r["modele_ia"] = "gemini-3.1-flash-lite ×3 + gemini-3.8-flash"
            classer_statut_final(r, r["decision_ia"])
        else:
            r["escalade_38"] = "ECHEC"; r["raison_ia"] = str(r.get("raison_ia", "")) + " | 3.8 indisponible, décision 3.1 T3 conservée."

    # ------------------------------------------------------------
    # CONTRÔLE GÉOGRAPHIQUE FINAL — INVIOLABLE
    # ------------------------------------------------------------
    # Même si Gemini a retenu un candidat, une localisation d'événement
    # explicitement hors périmètre reprend la main. Les mentions d'autres
    # territoires dans le public/organisateur ne sont pas utilisées ici.
    for r in t3_ok:
        geo_final = preanalyser_geographie(r)
        geo_fort_ok, geo_fort_motif, geo_fort_code = verifier_geographie_inviolable(r)
        r.update({k: v for k, v in geo_final.items() if k.startswith("geo_")})
        if (not geo_fort_ok) or geo_final.get("geo_preanalyse") == "GEO_HORS_CIBLE":
            r.update({
                "decision_ia": "NON_PERTINENT",
                "confiance_ia": int(geo_final.get("geo_confiance", 0) or 0),
                "raison_ia": "Contrôle géographique final : " + (
                    geo_fort_motif if not geo_fort_ok else geo_final.get("geo_motif", "hors périmètre")
                ),
                "motif_exclusion": "HORS_GEOGRAPHIE",
                "statut_final": "Non pertinent",
                "geo_controle_final": "BLOQUE",
            })
        else:
            r["geo_controle_final"] = "OK"

    ia_resultats.extend(t3_ok)
    ia_resultats.extend(non_eligibles)
    for r in ia_resultats:
        if r.get("memoire_ia") == "OUI":
            r["decision_ia"] = r.get("decision_ia") or r.get("decision_ia_1", "A_VERIFIER")
            classer_statut_final(r, r["decision_ia"])

    nouveaux_pertinents = sorted([r for r in ia_resultats if r.get("decision_ia") == "PERTINENT"], key=lambda x:x.get("confiance_ia",0), reverse=True)
    nouveaux_a_verifier = sorted([r for r in ia_resultats if r.get("decision_ia") == "A_VERIFIER"], key=lambda x:x.get("score",0), reverse=True)

    print(f"\nAnalyses Gemini 3.1 : {appels_31} ({len(t3_ok)} candidats × 3 si complet)")
    print(f"Gemini 3.8 : {appels_38}/{MAX_IA_38}")
    print(f"Top 3 sélectionné : {len(top_38)}")
    print(f"Réutilisations BDD : {len(reutilises_bdd)} | mémoire IA : {len(memorises)}")
    print(f"Pré-exclusions géographiques : {len(geo_preexclus)}")
    print(f"Blocage 3.1 : {'OUI' if quota_bloque else 'NON'}")

    global LAST_IA_STATS
    LAST_IA_STATS = {
        "appels_31": appels_31,
        "appels_38": appels_38,
        "candidats_analyzes": len(t3_ok),
        "a_traiter_reellement": len(a_traiter),
        "bdd_reutilises": len(reutilises_bdd) + len(geo_reutilises_bdd),
        "memoire_reutilisee": len(memorises),
        "memoire_globale_bypass_ia": 0,
        "geo_bloques": len(geo_preexclus),
        "quota_bloque": quota_bloque,
    }

    if not (IA_TEST_MODE and IA_TEST_NO_PERSIST):
        a_persister = [r for r in ia_resultats if r.get("bdd_reutilise") is not True and r.get("ia_analyse") in {"OUI", "GEO_PREEXCLUSION"}]
        if a_persister:
            enregistrer_candidats_bdd(a_persister, bdd_candidats)
    else:
        print("🧪 MODE TEST : aucune persistance BDD pour cette phase.")

    return nouveaux_pertinents, nouveaux_a_verifier, ia_resultats

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

    # Les ajouts utiles correspondent aux lignes nouvelles que le pipeline
    # est susceptible d'intégrer au classeur Excel. Les non-pertinents ne
    # sont volontairement jamais comptés comme ajouts.
    ajouts_tableur = nouveaux_pertinents + nouveaux_a_verifier_2027
    modifications_tableur = promus_pertinents + evolutions

    return {
        "nouveaux_pertinents": nouveaux_pertinents,
        "promus_pertinents": promus_pertinents,
        "nouveaux_a_verifier_2027": nouveaux_a_verifier_2027,
        "evolutions_utiles": len(evolutions),
        "evolutions": evolutions,
        "ajouts_tableur": ajouts_tableur,
        "modifications_tableur": modifications_tableur,
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


def executer_sorties_v8_v10(delta, stats_operateurs, pertinents, a_verifier, non_pertinents, ia_resultats, resultats_bruts):
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
        "requetes": len(requetes) if "requetes" in globals() else 0,
        "resultats_rss": meta_resultats.get("resultats_rss", 0) if "meta_resultats" in locals() else 0,
        "resultats_externes": meta_resultats.get("resultats_externes", 0) if "meta_resultats" in locals() else 0,
        "resultats_collectes_bruts": nb_resultats_collectes_bruts if "nb_resultats_collectes_bruts" in locals() else len(resultats_bruts),
        "resultats_bruts": len(resultats_bruts),
        "rejets_pretraitement": len(rejets_pretraitement) if "rejets_pretraitement" in locals() else 0,
        "resultats_conserves": len(pertinents) + len(a_verifier) + len(non_pertinents),
        "candidats_ia": len(ia_resultats),
        "pertinents": len(pertinents),
        "a_verifier": len(a_verifier),
        "non_pertinents": len(non_pertinents),
        "excel_synchronise": meta_resultats.get("excel_synchronise", "NON") if "meta_resultats" in globals() else "NON",
        "ia_analysees": sum(1 for r in ia_resultats if r.get("ia_analyse") == "OUI"),
        "ia_memoire": sum(1 for r in ia_resultats if r.get("ia_analyse") == "MEMOIRE"),
        "ia_non_analyse": sum(1 for r in ia_resultats if r.get("decision_ia") == "IA_NON_ANALYSE"),
        "gemini_31": sum(1 for r in ia_resultats if str(r.get("modele_ia_31", "")).strip()),
        "gemini_38": sum(1 for r in ia_resultats if str(r.get("modele_ia_38", "")).strip()),
        "gemini_31_cooldown": 10,
        "gemini_38_cooldown": 60,
        "quota_bloque": "OUI" if any(str(r.get("decision_ia", "")) == "IA_INDISPONIBLE" for r in ia_resultats) else "NON",
        "taches_terminees": resume_taches["terminees"],
        "taches_total": resume_taches["total"],
        "taches_erreurs": resume_taches["erreurs"],
    }
    meta.update({
        "gemini_model": "gemini-3.1-flash-lite → gemini-3.8-flash",
        "gemini_score_min": GEMINI_SCORE_MIN,
        "gemini_max_analyses": "3.1 illimité / 3.8 max 5",
        "gemini_cooldown": "3.1: 10 s / 3.8: 60 s",
        "memoire_ia_fichier": str(IA_MEMORY_FILE.relative_to(ROOT)),
        "memoire_ia_reutilisations": sum(1 for r in ia_resultats if r.get("memoire_ia") == "OUI"),
        "memoire_ia_nouveaux": sum(1 for r in ia_resultats if r.get("ia_analyse") == "OUI" and r.get("memoire_ia") != "OUI"),
        "memoire_globale_fichier": str(GLOBAL_MEMORY_FILE.relative_to(ROOT)),
        "memoire_globale_reutilisations": sum(1 for r in resultats_bruts if r.get("memoire_globale") == "OUI"),
        "memoire_globale_rafraichissement_jours": MEMORY_REFRESH_DAYS,
    })
    card = construire_resume_teams(delta, stats_operateurs, meta, _SUIVI_TACHES)
    sauvegarder_payload(card, OUTPUT / "teams_payload.json")
    resultat_teams = envoyer_teams(card, DATA / "teams_last_sent.txt")
    try:
        (OUTPUT / "teams_result.json").write_text(
            json.dumps(resultat_teams, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
    except Exception as exc:
        print(f"⚠️ Diagnostic Teams non sauvegardé : {type(exc).__name__}: {exc}")
    print(f"Teams : {'OUI' if resultat_teams['envoye'] else 'NON'} — {resultat_teams['raison']}")
    if resultat_teams.get("diagnostic"):
        print(f"Teams diagnostic : {json.dumps(resultat_teams['diagnostic'], ensure_ascii=False)}")
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
        # Aucun score négatif n'est autorisé dans les sorties.
        try:
            resultat["score"] = max(0, int(float(resultat.get("score", 0) or 0)))
        except Exception:
            resultat["score"] = 0

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
        "memoire_scoring",
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
        "memoire_ia",
        "date_ia_memoire",
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
        if r.get("bdd_reutilise") is not True
        and (
            r.get("ia_analyse") in ("OUI", "MEMOIRE")
            or r.get("decision_ia") == "IA_INDISPONIBLE"
            or r.get("decision_ia") == "IA_NON_ANALYSE"
        )
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


# ---------------------------------------------------------------------------
# BDD PERSISTANTE DES CANDIDATS
# ---------------------------------------------------------------------------

def _normaliser_cle_candidat(valeur):
    return re.sub(r"\s+", " ", str(valeur or "").strip().lower())

def _cle_stable_candidat(resultat):
    """
    Identifiant stable :
    - URL canonique si disponible
    - sinon titre + département + année
    """
    url = (
        resultat.get("url_finale")
        or resultat.get("url_reelle")
        or resultat.get("url")
        or ""
    ).strip()
    if url:
        return "url:" + _normaliser_cle_candidat(url)

    titre = _normaliser_cle_candidat(resultat.get("titre"))
    dept = _normaliser_cle_candidat(
        resultat.get("departement") or resultat.get("departement_numero")
    )
    annee = _normaliser_cle_candidat(
        resultat.get("annee_detectee") or resultat.get("annee_recherchee")
    )
    return "fallback:" + "|".join((titre, dept, annee))

def _signature_source_candidat(resultat):
    """
    Empreinte des données disponibles avant analyse approfondie.
    Elle sert à détecter une modification réelle du candidat.
    """
    champs = (
        "titre", "url", "url_reelle", "url_finale", "source", "description",
        "date_publication", "departement", "nom_departement",
        "annee_recherchee", "annee_detectee", "categorie", "mot_cle",
        "complement", "lieu", "operateur_web", "famille_operateur",
    )
    payload = {k: _normaliser_cle_candidat(resultat.get(k)) for k in champs}
    brut = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(brut.encode("utf-8")).hexdigest()

def _signature_contenu_candidat(resultat):
    """
    Empreinte enrichie. Les champs de page sont inclus lorsqu'ils existent.
    """
    champs = (
        "titre", "url_finale", "description", "date_publication",
        "date_evenement_page", "lieu_page", "organisateur_page",
        "extrait_page", "contenu_page", "page_enrichie",
    )
    payload = {k: _normaliser_cle_candidat(resultat.get(k)) for k in champs}
    brut = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(brut.encode("utf-8")).hexdigest()

def _charger_bdd_candidats():
    CANDIDATS_BDD_DIR.mkdir(parents=True, exist_ok=True)
    if not CANDIDATS_BDD_FILE.exists():
        return {"version": CANDIDATS_BDD_VERSION, "updated_at": "", "items": {}}
    try:
        data = json.loads(CANDIDATS_BDD_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("BDD invalide")
        data.setdefault("version", CANDIDATS_BDD_VERSION)
        data.setdefault("updated_at", "")
        data.setdefault("items", {})
        return data
    except Exception as exc:
        print(f"⚠️ BDD candidats illisible : {exc} — nouvelle BDD utilisée")
        return {"version": CANDIDATS_BDD_VERSION, "updated_at": "", "items": {}}

def _sauvegarder_bdd_candidats(bdd):
    CANDIDATS_BDD_DIR.mkdir(parents=True, exist_ok=True)
    bdd["version"] = CANDIDATS_BDD_VERSION
    bdd["updated_at"] = datetime.now().isoformat(timespec="seconds")
    tmp = CANDIDATS_BDD_FILE.with_suffix(".tmp")
    tmp.write_text(
        json.dumps(bdd, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    tmp.replace(CANDIDATS_BDD_FILE)

def _restaurer_resultat_bdd(resultat, entree):
    """
    Réinjecte les résultats précédemment calculés sans relancer
    enrichissement, scoring ou IA.
    """
    champs = entree.get("resultat", {})
    if isinstance(champs, dict):
        resultat.update(champs)
    resultat["bdd_candidat_connu"] = True
    resultat["bdd_reutilise"] = True
    resultat["bdd_modifie"] = False
    resultat["date_derniere_analyse"] = entree.get("last_analyzed", "")
    resultat["nombre_de_reanalyses"] = entree.get("analysis_count", 1)
    resultat["contenu_modifie"] = False
    return resultat

def classifier_candidats_bdd(resultats, bdd):
    """
    Sépare strictement :
    - nouveaux candidats
    - candidats modifiés
    - candidats déjà analysés et inchangés

    Les candidats déjà analysés sont restaurés depuis la BDD et ne doivent
    pas repasser dans les recherches approfondies / Gemini.
    """
    nouveaux = []
    modifies = []
    reutilises = []
    items = bdd.setdefault("items", {})

    for resultat in resultats:
        cle = _cle_stable_candidat(resultat)
        sig_source = _signature_source_candidat(resultat)
        entree = items.get(cle)

        resultat["bdd_cle"] = cle
        resultat["bdd_signature_source"] = sig_source

        if not entree:
            resultat["bdd_candidat_connu"] = False
            resultat["bdd_reutilise"] = False
            resultat["bdd_modifie"] = False
            nouveaux.append(resultat)
            continue

        ancienne_sig = entree.get("signature_source", "")
        if ancienne_sig == sig_source:
            reutilises.append(_restaurer_resultat_bdd(resultat, entree))
        else:
            resultat["bdd_candidat_connu"] = True
            resultat["bdd_reutilise"] = False
            resultat["bdd_modifie"] = True
            resultat["contenu_modifie"] = True
            modifies.append(resultat)

    print(
        f"🗃️ BDD candidats : {len(nouveaux)} nouveaux | "
        f"{len(modifies)} modifiés | {len(reutilises)} réutilisés"
    )
    return nouveaux, modifies, reutilises

def enregistrer_candidats_bdd(resultats, bdd):
    """
    Enregistre uniquement les candidats qui ont effectivement terminé leur
    analyse. Les erreurs non analysées ne sont donc pas figées en BDD.
    """
    items = bdd.setdefault("items", {})
    now = datetime.now().isoformat(timespec="seconds")

    for resultat in resultats:
        # Un candidat déjà réutilisé depuis la BDD ne doit pas être réécrit
        # ni compter comme une nouvelle analyse à chaque run.
        if resultat.get("bdd_reutilise") is True:
            continue
        # On ne fige dans la BDD que les candidats réellement analysés.
        # Une IA indisponible ou une analyse interrompue doit pouvoir être
        # retentée lors de la prochaine run.
        if resultat.get("ia_analyse") not in {"OUI", "MEMOIRE", "GEO_PREEXCLUSION"}:
            continue

        cle = resultat.get("bdd_cle") or _cle_stable_candidat(resultat)
        entree = items.get(cle, {})
        ancienne_signature = entree.get("signature_source", "")
        analyse_count = int(entree.get("analysis_count", 0) or 0) + 1

        # Ne pas stocker de références internes ou données de contrôle inutiles
        resultat_stocke = dict(resultat)
        entree.update({
            "signature_source": resultat.get("bdd_signature_source")
                or _signature_source_candidat(resultat),
            "signature_contenu": _signature_contenu_candidat(resultat),
            "first_seen": entree.get("first_seen", now),
            "last_seen": now,
            "last_analyzed": now,
            "analysis_count": analyse_count,
            "contenu_modifie": bool(
                ancienne_signature
                and ancienne_signature != resultat.get("bdd_signature_source")
            ),
            "status": resultat.get("decision_ia")
                or resultat.get("statut_final")
                or resultat.get("statut")
                or "",
            "resultat": resultat_stocke,
        })
        items[cle] = entree

    _sauvegarder_bdd_candidats(bdd)

def preparer_bdd_candidats():
    CANDIDATS_BDD_DIR.mkdir(parents=True, exist_ok=True)
    return _charger_bdd_candidats()



# ---------------------------------------------------------------------------
# IMPORT DE L'HISTORIQUE DES RUNS
# ---------------------------------------------------------------------------

def _lire_json_historique(path):
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            if isinstance(data.get("items"), dict):
                return list(data["items"].values())
            return [data]
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []

def _lire_csv_historique(path):
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            return [dict(row) for _, row in zip(range(HISTORIQUE_MAX_LIGNES_PAR_FICHIER), reader)]
    except Exception:
        return []

def _normaliser_historique_resultat(row):
    """
    Convertit une ligne historique en candidat exploitable par la BDD.
    Les informations historiques sont conservées, mais aucune analyse IA
    n'est déclenchée par cet import.
    """
    if not isinstance(row, dict):
        return None

    # Certains historiques peuvent avoir encapsulé le résultat.
    resultat = dict(row.get("resultat", row))
    if not isinstance(resultat, dict):
        return None

    # Ne pas créer d'entrées sans identité exploitable.
    if not (
        resultat.get("url")
        or resultat.get("url_reelle")
        or resultat.get("url_finale")
        or resultat.get("titre")
    ):
        return None

    return resultat

def importer_historique_runs_dans_bdd(bdd):
    """
    Analyse les fichiers produits par les runs précédentes présents dans
    output/ et data/. L'objectif est de récupérer le maximum d'informations
    déjà calculées : scores, décisions, enrichissements, dates, sources,
    statuts et résultats Gemini.

    Cet import est une opération de mémoire : il ne lance ni HTTP ni Gemini.
    Les fichiers les plus récents sont prioritaires.
    """
    fichiers = []

    # Résultats persistants principaux
    for base in (DATA, OUTPUT):
        if base.exists():
            for pattern in (
                "resultats_*.csv",
                "resultats_*.json",
                "journal_recherche_*.csv",
                "journal_recherche_*.json",
            ):
                fichiers.extend(base.glob(pattern))

    # Archives de recherches déjà créées
    for base in (
        DATA / "recherches" / "rss",
        DATA / "recherches" / "approfondies",
    ):
        if base.exists():
            fichiers.extend(base.glob("*.json"))
            fichiers.extend(base.glob("*.csv"))

    fichiers = sorted(
        {f for f in fichiers if f.is_file()},
        key=lambda x: x.stat().st_mtime,
        reverse=True,
    )[:HISTORIQUE_MAX_FICHIERS]

    items = bdd.setdefault("items", {})
    importes = 0
    enrichis = 0
    ignores = 0

    # Champs historiques utiles qui peuvent compléter une entrée BDD.
    champs_prioritaires = {
        k for k in (
            "titre", "url", "url_reelle", "url_finale", "source",
            "description", "date_publication", "departement",
            "nom_departement", "annee_recherchee", "annee_detectee",
            "categorie_numero", "categorie", "mot_cle", "complement",
            "lieu", "operateur_web", "famille_operateur",
            "score", "score_cible_apec", "score_emploi", "score_formation",
            "score_superieur", "score_apec", "score_entreprise",
            "score_qualifie", "score_geographie", "score_date",
            "score_bruit", "score_complementaire", "bonus_titre",
            "details_score", "statut", "statut_final", "decision_ia",
            "confiance_ia", "raison_ia", "modele_ia", "modele_ia_1",
            "decision_ia_1", "confiance_ia_1", "raison_ia_1",
            "modele_ia_2", "decision_ia_2", "confiance_ia_2",
            "raison_ia_2", "modele_ia_31", "decision_ia_31",
            "confiance_ia_31", "raison_ia_31", "modele_ia_38",
            "decision_ia_38", "confiance_ia_38", "raison_ia_38",
            "enrichissement_38", "page_enrichie", "titre_page",
            "description_page", "extrait_page", "date_evenement_page",
            "lieu_page", "organisateur_page", "evenement_detecte",
            "preuve_evenement", "geographie_verifiee", "hors_perimetre",
            "hors_geographie", "priorite_2027", "niveau_source_v9",
        )
    }

    for fichier in fichiers:
        if fichier.name == CANDIDATS_BDD_FILE.name:
            continue

        lignes = (
            _lire_csv_historique(fichier)
            if fichier.suffix.lower() == ".csv"
            else _lire_json_historique(fichier)
        )

        for ligne in lignes:
            resultat = _normaliser_historique_resultat(ligne)
            if not resultat:
                ignores += 1
                continue

            cle = _cle_stable_candidat(resultat)
            sig = _signature_source_candidat(resultat)
            now = datetime.fromtimestamp(
                fichier.stat().st_mtime
            ).isoformat(timespec="seconds")

            entree = items.get(cle)
            donnees = {
                k: resultat.get(k)
                for k in champs_prioritaires
                if k in resultat and str(resultat.get(k, "")).strip() != ""
            }

            if entree is None:
                items[cle] = {
                    "signature_source": sig,
                    "signature_contenu": _signature_contenu_candidat(resultat),
                    "first_seen": now,
                    "last_seen": now,
                    "last_analyzed": now,
                    "analysis_count": 1,
                    "status": (
                        resultat.get("decision_ia")
                        or resultat.get("statut_final")
                        or resultat.get("statut")
                        or ""
                    ),
                    "source_historique": fichier.name,
                    "resultat": donnees,
                }
                importes += 1
            else:
                # On enrichit l'entrée existante sans écraser une information
                # plus récente par une ligne historique plus ancienne.
                cible = entree.setdefault("resultat", {})
                for k, v in donnees.items():
                    if str(v).strip() and not str(cible.get(k, "")).strip():
                        cible[k] = v
                        enrichis += 1

                # Une ancienne analyse valide est conservée.
                if not entree.get("signature_source"):
                    entree["signature_source"] = sig
                entree.setdefault("first_seen", now)
                entree.setdefault("last_seen", now)
                entree.setdefault("last_analyzed", now)
                entree.setdefault("analysis_count", 1)

    if importes or enrichis:
        _sauvegarder_bdd_candidats(bdd)

    print(
        f"📚 Historique des runs : {len(fichiers)} fichiers inspectés | "
        f"{importes} candidats importés | {enrichis} informations complémentaires | "
        f"{ignores} lignes ignorées",
        flush=True,
    )
    return {
        "fichiers": len(fichiers),
        "importes": importes,
        "enrichis": enrichis,
        "ignores": ignores,
    }


def generer_rapport_pdf_unique(
    delta=None,
    stats_operateurs=None,
    pertinents=None,
    a_verifier=None,
    non_pertinents=None,
    ia_resultats=None,
    resultats_bruts=None,
    externes=None,
    rejets_pretraitement=None,
    motifs_pretraitement=None,
    audit_pipeline=None,
    erreurs=None,
    duree_total=0,
    appels_31=0,
    appels_38=0,
    sauvegarde_ok=False,
    synchronisation_ok=False,
):
    """Crée EXACTEMENT UN PDF par exécution dans output/.

    Aucun dossier run_... et aucun PDF intermédiaire n'est créé.
    Le fichier est remplacé s'il existe déjà pour la même exécution.
    """
    if not REPORTLAB_DISPONIBLE:
        print("⚠️ ReportLab absent : rapport PDF non généré.")
        return None

    OUTPUT.mkdir(exist_ok=True)

    # Un seul nom de rapport pour l'exécution.
    horodatage = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    pdf_path = OUTPUT / f"rapport_veille_{horodatage}.pdf"

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="RapportTitle", parent=styles["Title"],
        alignment=TA_CENTER, fontSize=19, leading=23, spaceAfter=8
    ))
    styles.add(ParagraphStyle(
        name="RapportSub", parent=styles["Normal"],
        alignment=TA_CENTER, fontSize=9, textColor=colors.grey, spaceAfter=14
    ))
    styles.add(ParagraphStyle(
        name="RapportSection", parent=styles["Heading2"],
        fontSize=13, leading=16, spaceBefore=10, spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        name="RapportSmall", parent=styles["Normal"],
        fontSize=8.2, leading=10.5
    ))

    def valeur(obj, cle, defaut="—"):
        if isinstance(obj, dict):
            v = obj.get(cle, defaut)
            return defaut if v in (None, "") else str(v)
        return defaut

    def nb(x):
        try:
            return len(x)
        except Exception:
            return 0

    # Les erreurs peuvent être un entier ou une liste selon le pipeline.
    erreurs_nb = len(erreurs) if isinstance(erreurs, (list, tuple, dict, set)) else (erreurs or 0)

    story = [
        Paragraph("RAPPORT DE VEILLE APEC", styles["RapportTitle"]),
        Paragraph(
            f"Exécution du {datetime.now().strftime('%d/%m/%Y à %H:%M:%S')} — document unique",
            styles["RapportSub"]
        ),
        Paragraph("1. Synthèse", styles["RapportSection"]),
    ]

    # Styles dédiés aux cellules : Paragraph force le retour à la ligne
    # et empêche tout débordement horizontal. splitLongWords permet aussi
    # de couper les chaînes longues sans espaces.
    styles.add(ParagraphStyle(
        name="RapportCell", parent=styles["Normal"],
        fontSize=7.2, leading=8.6, wordWrap="CJK", splitLongWords=True,
        spaceAfter=0, spaceBefore=0
    ))
    styles.add(ParagraphStyle(
        name="RapportCellSmall", parent=styles["Normal"],
        fontSize=8.2, leading=9.8, wordWrap="CJK", splitLongWords=True,
        spaceAfter=0, spaceBefore=0
    ))
    styles.add(ParagraphStyle(
        name="RapportCellHead", parent=styles["Normal"],
        fontSize=7.2, leading=8.5, fontName="Helvetica-Bold",
        wordWrap="CJK", splitLongWords=True, spaceAfter=0, spaceBefore=0
    ))

    def cell(value, style="RapportCell"):
        txt = "" if value is None else str(value)
        txt = xml_escape(txt).replace("\n", "<br/>")
        return Paragraph(txt or "—", styles[style])

    def header_row(values):
        return [cell(v, "RapportCellHead") for v in values]

    audit = dict(AUDIT_PIPELINE)
    if isinstance(audit_pipeline, dict):
        audit.update(audit_pipeline)
    motifs_audit = audit.get("motifs_rejets") or motifs_pretraitement or {}
    resume = [
        header_row(["Indicateur", "Résultat"]),
        [cell("Requêtes générées"), cell(audit.get("requetes_generees", 0))],
        [cell("Résultats RSS bruts"), cell(audit.get("rss_brut", 0))],
        [cell("Sources externes brutes"), cell(audit.get("externes_brut", 0))],
        [cell("Total brut collecté"), cell(audit.get("total_brut", 0))],
        [cell("Candidats après pré-filtre"), cell(audit.get("apres_pre_filtre", 0))],
        [cell("Rejets avant scoring / IA"), cell(audit.get("rejets_pre_filtre", 0))],
        [cell("Après traitement Python"), cell(audit.get("apres_traitement", 0))],
        [cell("Entrée couche IA"), cell(audit.get("candidats_ia_entree", 0))],
        [cell("À traiter réellement par Gemini"), cell(audit.get("ia_a_traiter_reellement", 0))],
        [cell("Analyses IA réelles"), cell(audit.get("candidats_ia_reels", 0))],
        [cell("Décisions IA réelles"), cell(audit.get("resultats_ia_reels", 0))],
        [cell("Réutilisations BDD"), cell(audit.get("ia_reutilises_bdd", 0))],
        [cell("Réutilisations mémoire IA"), cell(audit.get("ia_reutilises_memoire", 0))],
        [cell("Sources de découverte"), cell(audit.get("sources_decouverte", 0))],
        [cell("Événements découverts"), cell(audit.get("candidats_decouverts", 0))],
        [cell("Candidats perdus — réconciliation"), cell(audit.get("candidats_perdus_reconciliation", 0))],
        [cell("Preuves vérifiées V42"), cell(audit.get("preuves_verifiees", 0))],
        [cell("Preuves fortes"), cell(audit.get("preuves_fortes", 0))],
        [cell("Preuves moyennes"), cell(audit.get("preuves_moyennes", 0))],
        [cell("Sources secondaires trouvées"), cell(audit.get("sources_secondaires_trouvees", 0))],
        [cell("Pertinents finaux"), cell(audit.get("pertinents_final", len(pertinents)))],
        [cell("À vérifier finaux"), cell(audit.get("a_verifier_final", len(a_verifier)))],
        [cell("Non pertinents finaux"), cell(audit.get("non_pertinents_final", len(non_pertinents)))],
        [cell("Appels Gemini 3.1"), cell(appels_31)],
        [cell("Appels Gemini 3.8"), cell(appels_38)],
        [cell("Erreurs"), cell(erreurs_nb)],
        [cell("Durée totale"), cell(f"{float(duree_total):.1f} s")],
        [cell("Sauvegarde"), cell("OUI" if sauvegarde_ok else "NON")],
        [cell("Synchronisation Excel"), cell("OUI" if synchronisation_ok else "NON")],
    ]
    t = Table(resume, colWidths=[75*mm, 100*mm], repeatRows=1, splitByRow=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E8E8E8")),
        ("GRID", (0,0), (-1,-1), 0.4, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    story.append(t)

    story.append(Paragraph("2. Audit du pipeline", styles["RapportSection"]))
    story.append(Paragraph(
        "Cette section explique où les résultats disparaissent pendant le run. "
        "Les rejets sont ceux de cette exécution uniquement ; ils ne sont pas ajoutés "
        "aux résultats historiques affichés.", styles["RapportSmall"]
    ))
    audit_lignes = [header_row(["Étape", "Nombre"])]
    for lib, cle in [
        ("RSS brut", "rss_brut"), ("Sources externes brutes", "externes_brut"),
        ("Total brut", "total_brut"), ("Après pré-filtre", "apres_pre_filtre"),
        ("Rejets pré-filtre", "rejets_pre_filtre"), ("Après traitement Python", "apres_traitement"),
        ("Entrée couche IA", "candidats_ia_entree"),
        ("À traiter réellement par Gemini", "ia_a_traiter_reellement"),
        ("Analyses IA réelles", "candidats_ia_reels"),
        ("Décisions IA réelles", "resultats_ia_reels"),
        ("Réutilisations BDD", "ia_reutilises_bdd"),
        ("Réutilisations mémoire IA", "ia_reutilises_memoire"),
        ("Sources de découverte", "sources_decouverte"),
        ("Événements découverts", "candidats_decouverts"),
        ("Candidats perdus — réconciliation", "candidats_perdus_reconciliation"),
        ("Preuves vérifiées V42", "preuves_verifiees"),
        ("Preuves fortes", "preuves_fortes"),
        ("Preuves moyennes", "preuves_moyennes"),
        ("Sources secondaires trouvées", "sources_secondaires_trouvees"),
    ]:
        audit_lignes.append([cell(lib), cell(audit.get(cle, 0))])
    if motifs_audit:
        for motif, count in sorted(motifs_audit.items(), key=lambda kv: (-int(kv[1]), str(kv[0]))):
            audit_lignes.append([cell("Rejet — " + str(motif)), cell(count)])
    t_audit = Table(audit_lignes, colWidths=[120*mm, 55*mm], repeatRows=1, splitByRow=1)
    t_audit.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E8E8E8")),
        ("GRID", (0,0), (-1,-1), 0.35, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 4), ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 3), ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    story.append(t_audit)

    rejets = list(rejets_pretraitement or [])
    if rejets:
        story.append(Paragraph("2.1 Rejets du run — audit", styles["RapportSection"]))
        rej_lignes = [header_row(["Candidat rejeté", "Motif", "Géographie / preuve"])]
        for r in rejets[:40]:
            preuve = (r.get("hors_geographie") or r.get("motif_exclusion") or
                      r.get("pre_filtre_motif") or "—")
            rej_lignes.append([
                cell(r.get("titre", r.get("title", "Sans titre"))),
                cell(r.get("pre_filtre_motif", r.get("motif_exclusion", "—"))),
                cell(preuve),
            ])
        t_rej = Table(rej_lignes, colWidths=[70*mm, 45*mm, 60*mm], repeatRows=1, splitByRow=1)
        t_rej.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E8E8E8")),
            ("GRID", (0,0), (-1,-1), 0.35, colors.grey), ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 3), ("RIGHTPADDING", (0,0), (-1,-1), 3),
            ("TOPPADDING", (0,0), (-1,-1), 3), ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ]))
        story.append(t_rej)

    story.append(Paragraph("3. Résultats prioritaires", styles["RapportSection"]))
    lignes = [header_row(["Candidat", "Statut", "Score", "Raison / preuve"])]
    tous = list(pertinents or []) + list(a_verifier or [])
    for c in tous[:50]:
        # Aucun score négatif ne doit être affiché dans le PDF.
        try:
            score_affiche = max(0, int(float(c.get("score", 0) or 0)))
        except Exception:
            score_affiche = 0
        lignes.append([
            cell(valeur(c, "titre", valeur(c, "title", "Sans titre"))),
            cell(valeur(c, "statut")),
            cell(score_affiche),
            cell(
                "Preuve: " + valeur(c, "preuve_evenement_v42", "—")
                + " | " + valeur(c, "preuve_cible", "—")
                + " | " + valeur(c, "raison", valeur(c, "motif", ""))
            ),
        ])
    if len(lignes) == 1:
        lignes.append([cell("Aucun candidat"), cell("—"), cell("—"), cell("Aucun résultat prioritaire")])
    t = Table(lignes, colWidths=[55*mm, 30*mm, 18*mm, 72*mm], repeatRows=1, splitByRow=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E8E8E8")),
        ("GRID", (0,0), (-1,-1), 0.35, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    story.append(t)

    story.append(Paragraph("4. Analyse IA et arbitrage", styles["RapportSection"]))
    story.append(Paragraph(
        "Le rapport conserve les décisions IA disponibles à la fin de l'exécution. "
        "Les trois tours Gemini 3.1 et la sélection Gemini 3.8 sont regroupés dans "
        "ce même document, sans génération de rapports PDF séparés.",
        styles["RapportSmall"]
    ))

    ia_lignes = [header_row(["Candidat", "Décision", "Confiance", "Analyse / action"])]
    for c in list(ia_resultats or [])[:50]:
        analyse = (
            "Preuve événement: " + valeur(c, "preuve_evenement_v42", "—")
            + " | Sources secondaires: " + valeur(c, "nombre_sources_secondaires_v42", "0")
            + " | " + valeur(c, "analyse_humaine", valeur(c, "raison", ""))
        )
        ia_lignes.append([
            cell(valeur(c, "titre", valeur(c, "title", "Sans titre"))),
            cell(valeur(c, "statut", valeur(c, "decision", "—"))),
            cell(valeur(c, "confiance", valeur(c, "confiance_31", "—"))),
            cell(analyse),
        ])
    if len(ia_lignes) == 1:
        ia_lignes.append([cell("Aucun résultat IA"), cell("—"), cell("—"), cell("—")])
    t = Table(ia_lignes, colWidths=[50*mm, 32*mm, 23*mm, 70*mm], repeatRows=1, splitByRow=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E8E8E8")),
        ("GRID", (0,0), (-1,-1), 0.35, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    story.append(t)

    story.append(Spacer(1, 8))
    story.append(Paragraph(
        "Règle de sortie : un seul fichier PDF est produit par exécution. "
        "Les données CSV/Excel internes restent inchangées ; elles ne génèrent pas "
        "de PDF supplémentaires.",
        styles["RapportSmall"]
    ))

    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=A4,
        rightMargin=15*mm, leftMargin=15*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title="Rapport de veille APEC — nouveaux et modifiés"
    )
    doc.build(story)

    print(f"📄 Rapport PDF unique créé : {pdf_path}")
    return pdf_path


def main():
    bdd_candidats = preparer_bdd_candidats()
    stats_historique_runs = importer_historique_runs_dans_bdd(bdd_candidats)
    stats_bdd_candidats = {
        "taille_avant_run": len(bdd_candidats.get("items", {}))
    }
    global meta_resultats

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
        f"Gemini : tous les candidats V36 | 3.1: illimité/10s | 3.8: max {MAX_IA_38}/60s | mémoire globale"
    )

    initialiser_suivi_taches()
    sauvegarder_suivi_taches()

    ancienne_veille = charger_ancienne_veille()

    departements = load_yaml("departements.yml")
    requetes = construire_requetes(departements)

    AUDIT_PIPELINE["requetes_generees"] = len(requetes)

    print()
    print(f"Requêtes générées : {len(requetes)}")

    demarrer_tache(0, f"{len(requetes)} requêtes")
    try:
        resultats_bruts, erreurs = collecter(requetes)
        AUDIT_PIPELINE["rss_brut"] = len(resultats_bruts)
        terminer_tache(0, details=f"{len(resultats_bruts)} résultats, {erreurs} erreurs")
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
        AUDIT_PIPELINE["externes_brut"] = len(externes)
        terminer_tache(1, details=f"{len(externes)} résultats externes")

    AUDIT_PIPELINE["rss_brut"] = len(resultats_bruts)
    AUDIT_PIPELINE["externes_brut"] = len(externes)

    # Aucun journal intermédiaire : le rapport PDF unique est généré en fin de run.

    for _r in externes:
        _r.setdefault("origine_collecte", "externe")
    for _r in resultats_bruts:
        _r.setdefault("origine_collecte", "rss")

    # --------------------------------------------------------
    # V40 — SOURCES DE DÉCOUVERTE
    # --------------------------------------------------------
    sources_decouverte, candidats_decouverts = extraire_evenements_sources_decouverte(externes)
    AUDIT_PIPELINE["sources_decouverte"] = len(sources_decouverte)
    AUDIT_PIPELINE["candidats_decouverts"] = len(candidats_decouverts)
    print(f"🔎 Sources de découverte : {len(sources_decouverte)} | événements enfants : {len(candidats_decouverts)}")

    # Une page institutionnelle de navigation n'est plus un candidat.
    # Seuls ses événements enfants éventuels rejoignent la collecte.
    resultats_collectes = candidats_decouverts + resultats_bruts + [
        r for r in externes if not _est_source_decouverte(r)
    ]

    # --------------------------------------------------------
    # FILTRE V40 AVANT TOUT SCORING
    # --------------------------------------------------------
    # Les sources externes et RSS sont filtrées ensemble, mais les compteurs
    # de provenance sont conservés séparément. Les contenus manifestement
    # hors sujet ne deviennent jamais des « candidats » de la veille.
    nb_resultats_collectes_bruts = len(resultats_collectes)
    AUDIT_PIPELINE["total_brut"] = nb_resultats_collectes_bruts
    resultats_bruts, rejets_pretraitement, motifs_pretraitement = filtrer_candidats_avant_scoring(resultats_collectes)
    AUDIT_PIPELINE["apres_pre_filtre"] = len(resultats_bruts)
    AUDIT_PIPELINE["rejets_pre_filtre"] = len(rejets_pretraitement)
    AUDIT_PIPELINE["motifs_rejets"] = dict(motifs_pretraitement)

    print()
    print("========================================")
    print("TRAITEMENT")
    print("========================================")
    print(f"Résultats collectés : {len(resultats_collectes)}")
    print(f"Candidats après filtre V36 : {len(resultats_bruts)}")
    print(f"Rejets avant scoring/IA : {len(rejets_pretraitement)}")

    if not resultats_bruts:
        print()
        print("⚠️ AUCUN CANDIDAT APRÈS LE PRÉ-FILTRE.")
        print("Le run continue pour produire l’audit complet et conserver l’historique.")
        print(f"Erreurs HTTP/RSS : {erreurs}")

    # Mémoire globale chargée AVANT le scoring : un résultat inchangé
    # réutilise directement son score Python et ses composants, sans refaire
    # calculer_score(). Les nouveaux/modifiés sont recalculés puis mémorisés.
    memoire_globale = charger_memoire_globale()

    demarrer_tache(2, f"{len(resultats_bruts)} résultats bruts")
    pertinents, a_verifier, non_pertinents = traiter(resultats_bruts, memoire_globale)
    AUDIT_PIPELINE["apres_traitement"] = len(pertinents) + len(a_verifier) + len(non_pertinents)
    AUDIT_PIPELINE["pertinents_python"] = len(pertinents)
    AUDIT_PIPELINE["a_verifier_python"] = len(a_verifier)
    AUDIT_PIPELINE["non_pertinents_python"] = len(non_pertinents)
    terminer_tache(2, details=f"{AUDIT_PIPELINE['apres_traitement']} résultats conservés après traitement")

    demarrer_tache(3)
    terminer_tache(3, details=f"{len(pertinents)} pertinents / {len(a_verifier)} à vérifier / {len(non_pertinents)} non pertinents")

    # --------------------------------------------------------
    # COUCHE IA
    # --------------------------------------------------------
    demarrer_tache(5, f"{len(pertinents) + len(a_verifier)} candidats")
    memoire_ia = charger_memoire_ia()
    AUDIT_PIPELINE["candidats_ia_entree"] = len(pertinents) + len(a_verifier)
    pertinents_ia, a_verifier_ia, ia_resultats = appliquer_ia(
        pertinents,
        a_verifier,
        memoire_ia,
        bdd_candidats,
    )

    AUDIT_PIPELINE["resultats_ia"] = len(ia_resultats)
    AUDIT_PIPELINE["resultats_ia_reels"] = sum(1 for r in ia_resultats if r.get("bdd_reutilise") is not True and r.get("memoire_ia") != "OUI" and r.get("ia_analyse") in {"OUI", "GEO_PREEXCLUSION"})
    AUDIT_PIPELINE["candidats_ia_reels"] = LAST_IA_STATS.get("candidats_analyzes", 0)
    AUDIT_PIPELINE["ia_reutilises_bdd"] = LAST_IA_STATS.get("bdd_reutilises", 0)
    AUDIT_PIPELINE["ia_reutilises_memoire"] = LAST_IA_STATS.get("memoire_reutilisee", 0)
    AUDIT_PIPELINE["ia_a_traiter_reellement"] = LAST_IA_STATS.get("a_traiter_reellement", 0)

    # V40 — aucune disparition silencieuse entre entrée IA et destinations.
    entree_ia = AUDIT_PIPELINE.get("candidats_ia_entree", 0)
    bdd_r = LAST_IA_STATS.get("bdd_reutilises", 0)
    mem_r = LAST_IA_STATS.get("memoire_reutilisee", 0)
    gem_r = LAST_IA_STATS.get("a_traiter_reellement", 0)
    geo_r = LAST_IA_STATS.get("geo_bloques", 0)
    total_reconcilie = bdd_r + mem_r + gem_r + geo_r
    AUDIT_PIPELINE["candidats_perdus_reconciliation"] = max(0, entree_ia - total_reconcilie)
    if entree_ia == total_reconcilie:
        print(f"✅ V41 RÉCONCILIATION : {entree_ia} = {total_reconcilie}")
    else:
        print(f"⚠️ V41 RÉCONCILIATION : entrée={entree_ia} | BDD={bdd_r} | mémoire IA={mem_r} | Gemini={gem_r} | geo={geo_r} | perdus={AUDIT_PIPELINE['candidats_perdus_reconciliation']}")

    # ========================================================
    # DELTA D'AFFICHAGE — UNIQUEMENT NOUVEAUX / MODIFIÉS
    # ========================================================
    # La BDD conserve tout l'historique, mais aucun candidat simplement
    # réutilisé ne doit réapparaître dans les sorties du run courant.
    # Un candidat modifié est considéré comme nouveau pour l'affichage du
    # jour, car son contenu/source a changé depuis sa dernière analyse.
    # Diagnostic V39 : si beaucoup de candidats entrent dans la couche IA mais
    # aucun ne passe réellement par Gemini, le rapport doit l'indiquer clairement.
    if AUDIT_PIPELINE.get("candidats_ia_entree", 0) and AUDIT_PIPELINE.get("ia_a_traiter_reellement", 0) == 0:
        print("ℹ️ V39 : aucun appel Gemini nécessaire sur ce run : candidats déjà connus/réutilisés.")

    resultats_du_jour = [
        r for r in ia_resultats
        if r.get("bdd_reutilise") is not True
        and r.get("ia_analyse") != "MEMOIRE"
    ]
    urls_du_jour = {str(r.get("url") or "").strip() for r in resultats_du_jour}
    urls_du_jour.discard("")
    resultats_bruts_du_jour = [
        r for r in resultats_bruts
        if str(r.get("url") or "").strip() in urls_du_jour
    ]

    print(
        f"📌 Affichage du jour : {len(resultats_du_jour)} nouveaux/modifiés | "
        f"{len(ia_resultats) - len(resultats_du_jour)} anciens réutilisés masqués"
    )

    # La BDD candidats est alimentée uniquement après une analyse terminée.
    # Les candidats nouveaux/modifiés sont ainsi disponibles à la prochaine
    # run, tandis que les candidats inchangés seront réutilisés sans nouvel
    # appel HTTP/Gemini.
    try:
        enregistrer_candidats_bdd(ia_resultats, bdd_candidats)
        stats_bdd_candidats["taille_apres_run"] = len(bdd_candidats.get("items", {}))
        stats_bdd_candidats["analyses_enregistrees"] = sum(
            1 for r in ia_resultats
            if r.get("ia_analyse") in {"OUI", "MEMOIRE"} and not r.get("bdd_reutilise")
        )
        print(
            f"🗃️ BDD candidats sauvegardée : {stats_bdd_candidats['taille_apres_run']} candidats | "
            f"{stats_bdd_candidats['analyses_enregistrees']} analyses enregistrées"
        )
    except Exception as e:
        print(f"⚠️ BDD candidats non sauvegardée : {type(e).__name__} : {e}")

    try:
        sauvegarder_memoire_ia(ia_resultats, memoire_ia)
        print(f"🧠 Mémoire IA sauvegardée : {IA_MEMORY_FILE}")
    except Exception as e:
        print(f"⚠️ Impossible de sauvegarder la mémoire IA : {type(e).__name__} : {e}")

    try:
        sauvegarder_memoire_globale(resultats_bruts, memoire_globale)
        print(f"🧠 Mémoire globale sauvegardée : {GLOBAL_MEMORY_FILE}")
    except Exception as e:
        print(f"⚠️ Impossible de sauvegarder la mémoire globale : {type(e).__name__} : {e}")
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

    # Sorties quotidiennes : uniquement les candidats NOUVEAUX ou MODIFIÉS.
    # La BDD et les mémoires restent complètes, mais elles ne polluent jamais
    # le CSV/Excel/PDF/affichage quotidien.
    pertinents = [r for r in pertinents_ia if r.get("bdd_reutilise") is not True and r.get("ia_analyse") != "MEMOIRE"]
    a_verifier = [r for r in a_verifier_ia if r.get("bdd_reutilise") is not True and r.get("ia_analyse") != "MEMOIRE"]
    non_pertinents = [r for r in (non_pertinents_python + non_pertinents_ia) if r.get("bdd_reutilise") is not True and r.get("ia_analyse") != "MEMOIRE"]

    AUDIT_PIPELINE["pertinents_final"] = len(pertinents)
    AUDIT_PIPELINE["a_verifier_final"] = len(a_verifier)
    AUDIT_PIPELINE["non_pertinents_final"] = len(non_pertinents)

    total_dedoublonne = (
        len(pertinents)
        + len(a_verifier)
        + len(non_pertinents)
    )

    print()
    print("========================================")
    print("RÉSULTATS DU JOUR — ANCIENS MASQUÉS")
    print("========================================")
    pertinents_jour = [r for r in resultats_du_jour if r.get("decision_ia") == "PERTINENT"]
    a_verifier_jour = [r for r in resultats_du_jour if r.get("decision_ia") == "A_VERIFIER"]
    non_pertinents_jour = [r for r in resultats_du_jour if r.get("decision_ia") == "NON_PERTINENT"]
    print(f"Nouveaux/modifiés affichés : {len(resultats_du_jour)}")
    print(f"Pertinents du jour : {len(pertinents_jour)}")
    print(f"À vérifier du jour : {len(a_verifier_jour)}")
    print(f"Non pertinents du jour : {len(non_pertinents_jour)}")
    print(f"Anciens candidats masqués : {len(ia_resultats) - len(resultats_du_jour)}")
    print("\nAUDIT PIPELINE")
    print(f"  Requêtes : {AUDIT_PIPELINE['requetes_generees']}")
    print(f"  RSS brut : {AUDIT_PIPELINE['rss_brut']}")
    print(f"  Externes brut : {AUDIT_PIPELINE['externes_brut']}")
    print(f"  Total brut : {AUDIT_PIPELINE['total_brut']}")
    print(f"  Après pré-filtre : {AUDIT_PIPELINE['apres_pre_filtre']}")
    print(f"  Rejets pré-filtre : {AUDIT_PIPELINE['rejets_pre_filtre']}")
    print(f"  Entrée IA : {AUDIT_PIPELINE['candidats_ia_entree']}")
    print(f"  Analyses IA réelles : {AUDIT_PIPELINE['candidats_ia_reels']}")
    print("🛡️ Sorties quotidiennes verrouillées : NOUVEAU + MODIFIÉ uniquement")

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

    # Métadonnées de collecte utilisées par le rapport Teams.
    meta_resultats = {
        "resultats_rss": sum(1 for r in resultats_bruts if str(r.get("origine_collecte", "")).lower() not in {"externe", "source_externe"}),
        "resultats_externes": sum(1 for r in resultats_bruts if str(r.get("origine_collecte", "")).lower() in {"externe", "source_externe"}),
        "resultats_collectes_avant_filtre": len(resultats_collectes),
        "rejets_pretraitement": len(rejets_pretraitement),
    }

    # Les compteurs IA proviennent directement de appliquer_ia(), pas de locals()
    # de main(), afin que le PDF affiche les appels réellement effectués.
    meta_resultats["appels_gemini_31"] = LAST_IA_STATS.get("appels_31", 0)
    meta_resultats["appels_gemini_38"] = LAST_IA_STATS.get("appels_38", 0)
    meta_resultats["candidats_analyses_ia"] = LAST_IA_STATS.get("candidats_analyzes", 0)

    synchronisation_ok = False

    # Le rapport Teams est envoyé APRÈS la synchronisation Excel afin de
    # pouvoir annoncer honnêtement ce qui a réellement été intégré au classeur.
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

    meta_resultats["excel_synchronise"] = "OUI" if synchronisation_ok else "NON"

    demarrer_tache(10)
    sorties_v10 = executer_sorties_v8_v10(
        delta, stats_operateurs, pertinents, a_verifier, non_pertinents, resultats_du_jour, resultats_bruts
    )

    terminer_tache(10, details="Payload Teams préparé")
    if sorties_v10.get("envoye"):
        terminer_tache(11, details=sorties_v10.get("raison", ""))
    else:
        terminer_tache(11, "TERMINEE", details=sorties_v10.get("raison", ""))
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

    # ========================================================
    # RAPPORT PDF UNIQUE — UN SEUL DOCUMENT PAR EXECUTION
    # ========================================================
    try:
        generer_rapport_pdf_unique(
            delta=locals().get("delta"),
            stats_operateurs=locals().get("stats_operateurs"),
            pertinents=[r for r in pertinents_ia if r.get("bdd_reutilise") is not True],
            a_verifier=[r for r in a_verifier_ia if r.get("bdd_reutilise") is not True],
            non_pertinents=[r for r in non_pertinents if r.get("bdd_reutilise") is not True],
            ia_resultats=resultats_du_jour,
            resultats_bruts=resultats_bruts_du_jour,
            externes=[r for r in externes if str(r.get("url") or "").strip() in urls_du_jour],
            rejets_pretraitement=rejets_pretraitement,
            motifs_pretraitement=motifs_pretraitement,
            audit_pipeline=AUDIT_PIPELINE,
            erreurs=erreurs,
            duree_total=duree_total,
            appels_31=LAST_IA_STATS.get("appels_31", 0),
            appels_38=LAST_IA_STATS.get("appels_38", 0),
            sauvegarde_ok=sauvegarde_ok,
            synchronisation_ok=synchronisation_ok,
        )
    except Exception as e:
        print(f"⚠️ Génération du rapport PDF impossible : {type(e).__name__}: {e}")

    print()
    print("========================================")

    if sauvegarde_ok:
        print("VEILLE TERMINÉE AVEC SUCCÈS")
    else:
        print("VEILLE NON VALIDÉE")

    print("========================================")


if __name__ == "__main__":
    main()
