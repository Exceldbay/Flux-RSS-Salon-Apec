import os
import json
import re
import unicodedata
from google import genai
from google.genai import errors as genai_errors

MODEL_31 = os.getenv("GEMINI_MODEL_31", "gemini-3.1-flash-lite")
MODEL_38 = os.getenv("GEMINI_MODEL_38", "gemini-3.8-flash")
IA_RULES_VERSION = "2026-11-27-v32-geo-delta-arbitrage-preuves"


# ============================================================
# PRÉ-ANALYSE GÉOGRAPHIQUE IA — SANS APPEL GEMINI
# ============================================================
# Cette étape appartient à la couche IA mais reste locale et gratuite.
# Elle intervient avant les 3 tours Gemini afin d'éviter de consommer
# des appels sur des événements manifestement hors périmètre.

DEPARTEMENTS_CIBLES = {"16", "17", "19", "23", "79", "86", "87"}
DEPARTEMENTS_NOMS = {
    "01":"ain","02":"aisne","03":"allier","04":"alpes-de-haute-provence","05":"hautes-alpes",
    "06":"alpes-maritimes","07":"ardeche","08":"ardennes","09":"ariege","10":"aube","11":"aude",
    "12":"aveyron","13":"bouches-du-rhone","14":"calvados","15":"cantal","16":"charente",
    "17":"charente-maritime","18":"cher","19":"correze","23":"creuse","24":"dordogne",
    "25":"doubs","26":"drome","27":"eure","28":"eure-et-loir","29":"finistere",
    "30":"gard","31":"haute-garonne","32":"gers","33":"gironde","34":"herault",
    "35":"ille-et-vilaine","36":"indre","37":"indre-et-loire","38":"isere","39":"jura",
    "40":"landes","41":"loir-et-cher","42":"loire","43":"haute-loire","44":"loire-atlantique",
    "45":"loiret","46":"lot","47":"lot-et-garonne","48":"lozere","49":"maine-et-loire",
    "50":"manche","51":"marne","52":"haute-marne","53":"mayenne","54":"meurthe-et-moselle",
    "55":"meuse","56":"morbihan","57":"moselle","58":"nievre","59":"nord","60":"oise",
    "61":"orne","62":"pas-de-calais","63":"puy-de-dome","64":"pyrenees-atlantiques",
    "65":"hautes-pyrenees","66":"pyrenees-orientales","67":"bas-rhin","68":"haut-rhin","69":"rhone",
    "70":"haute-saone","71":"saone-et-loire","72":"sarthe","73":"savoie","74":"haute-savoie",
    "75":"paris","76":"seine-maritime","77":"seine-et-marne","78":"yvelines","79":"deux-sevres",
    "80":"somme","81":"tarn","82":"tarn-et-garonne","83":"var","84":"vaucluse","85":"vendee",
    "86":"vienne","87":"haute-vienne","88":"vosges","89":"yonne","90":"territoire de belfort",
    "91":"essonne","92":"hauts-de-seine","93":"seine-saint-denis","94":"val-de-marne","95":"val-d-oise",
}

# Grandes villes hors périmètre utiles pour les cas où seul le nom de ville
# est présent sans département/code postal. Cette liste ne sert jamais à
# exclure une ville mentionnée uniquement comme origine du public.
VILLES_HORS_CIBLE = {
    # Uniquement des villes hors périmètre. Les villes cibles (Limoges,
    # Poitiers, La Rochelle, Niort, etc.) ne doivent JAMAIS figurer ici.
    "nantes", "saint-nazaire", "angers", "tours", "bordeaux", "agen",
    "nice", "paris", "toulouse", "montpellier", "lille", "rennes",
    "orléans", "orleans", "clermont-ferrand", "clermont ferrand", "dijon",
    "grenoble", "lyon", "marseille", "strasbourg", "mulhouse", "metz",
    "nancy", "rouen", "caen", "amiens", "le mans", "le havre",
    "saint-malo", "périgueux", "perigueux", "bergerac", "libourne",
    "mont-de-marsan", "pau", "bayonne", "langon", "colmar", "alès", "ales",
    "montauban", "avignon", "troyes", "reims",
}

DEPARTEMENTS_CIBLES_NOMS = {DEPARTEMENTS_NOMS[c] for c in DEPARTEMENTS_CIBLES}

def _geo_norm(value):
    value = unicodedata.normalize("NFKD", str(value or "").lower())
    value = "".join(c for c in value if not unicodedata.combining(c))
    value = re.sub(r"[^a-z0-9\s-]", " ", value)
    return re.sub(r"\s+", " ", value).strip()

def _contient_mot(texte, mot):
    return bool(re.search(r"(?<![a-z0-9])" + re.escape(mot) + r"(?![a-z0-9])", texte))

def preanalyser_geographie(resultat):
    """Pré-analyse géographique stricte avant Gemini.

    Règle fondamentale : on cherche le LIEU DE L'ÉVÉNEMENT, pas toutes les
    villes citées dans le contenu. Une ville hors périmètre n'entraîne donc
    un rejet que lorsqu'elle est utilisée comme localisation explicite de
    l'événement (ex. « Campus Nantes », « à Nantes », « Nantes - ... »,
    code postal 44000). Une mention du type « entreprises de Nantes recrutent
    à Poitiers » ne rejette pas un événement à Poitiers.
    """
    champs_lieu = [
        resultat.get("lieu_page", ""),
        resultat.get("lieu", ""),
        resultat.get("adresse_page", ""),
    ]
    champs_titre = [resultat.get("titre_page", ""), resultat.get("titre", "")]
    champs_desc = [resultat.get("description_page", ""), resultat.get("description", "")]
    contenu = str(resultat.get("contenu_page", "") or "")[:16000]

    lieu = _geo_norm(" ".join(str(x or "") for x in champs_lieu))
    titre = _geo_norm(" ".join(str(x or "") for x in champs_titre))
    desc = _geo_norm(" ".join(str(x or "") for x in champs_desc))
    evenement = f"{titre} {desc}"

    zones, codes, villes, hors = [], [], [], []

    # 1) Code postal : priorité absolue quand présent dans un champ de lieu.
    for cp in re.findall(r"(?<!\d)(?:0[1-9]|[1-8]\d|9[0-5])\d{3}(?!\d)", lieu):
        code = cp[:2]
        codes.append(code)
        if code in DEPARTEMENTS_CIBLES:
            zones.append(f"{cp} ({DEPARTEMENTS_NOMS.get(code, code)})")
        else:
            hors.append(f"{cp} ({DEPARTEMENTS_NOMS.get(code, code)})")

    # 2) Département explicite dans un champ de lieu.
    for code, nom in DEPARTEMENTS_NOMS.items():
        nom_n = _geo_norm(nom)
        if nom_n and _contient_mot(lieu, nom_n):
            codes.append(code)
            if code in DEPARTEMENTS_CIBLES:
                zones.append(f"{nom} ({code})")
            else:
                hors.append(f"{nom} ({code})")

    # 3) Ville hors cible dans le champ lieu : preuve directe.
    for ville in VILLES_HORS_CIBLE:
        vn = _geo_norm(ville)
        if vn and _contient_mot(lieu, vn):
            villes.append(ville)
            hors.append(ville)

    # 4) Titre : les formulations de localisation explicite sont fortes.
    #    « Campus Nantes » est volontairement reconnu comme lieu.
    patterns_forts = [
        r"(?:\ba\s+|\bà\s+|\bcampus\s+){}(?:\s|[,|/\-]|$)",
        r"(?:^|[-–|/]\s*){}(?:\s*[,|/\-]|\s+\d{{5}}|$)",
        r"(?:^|\s){}\s*\(\s*\d{{2}}\s*\)",
    ]
    for ville in VILLES_HORS_CIBLE:
        vn = _geo_norm(ville)
        if not vn:
            continue
        for pat in patterns_forts:
            if re.search(pat.format(re.escape(vn)), titre, flags=re.I):
                villes.append(ville)
                hors.append(ville)
                break

    # 5) Description : seulement les formulations qui attribuent clairement
    #    une localisation à l'événement, jamais une simple mention contextuelle.
    for ville in VILLES_HORS_CIBLE:
        vn = _geo_norm(ville)
        if not vn:
            continue
        desc_patterns = [
            rf"\b(?:a|à)\s+{re.escape(vn)}(?:\s|,|\.|$)",
            rf"\b(?:lieu|adresse|site|campus)\s*[:\-]\s*[^.\n]*\b{re.escape(vn)}\b",
            rf"\b{re.escape(vn)}\s*\(\s*\d{{2}}\s*\)",
            rf"\b{re.escape(vn)}\s+\d{{5}}\b",
        ]
        if any(re.search(p, desc, flags=re.I) for p in desc_patterns):
            villes.append(ville)
            hors.append(ville)

    # 6) Contenu de page : uniquement des marqueurs de localisation forts.
    #    On évite de considérer toute occurrence de « Nantes » comme un lieu.
    for ville in VILLES_HORS_CIBLE:
        vn = _geo_norm(ville)
        if not vn:
            continue
        content_patterns = [
            rf"\b(?:lieu|adresse|site|campus)\s*[:\-]\s*[^\n]{{0,180}}\b{re.escape(vn)}\b",
            rf"\b(?:a|à)\s+{re.escape(vn)}\b[^\n]{{0,80}}\b(?:2026|2027|2028|portes ouvertes|forum|job dating|salon)\b",
        ]
        if any(re.search(p, contenu, flags=re.I) for p in content_patterns):
            villes.append(ville)
            hors.append(ville)

    codes = sorted(set(codes))
    zones = list(dict.fromkeys(zones))
    villes = list(dict.fromkeys(villes))
    hors = list(dict.fromkeys(hors))

    # Une preuve directe hors cible prime.
    if hors:
        return {
            "geo_preanalyse": "GEO_HORS_CIBLE",
            "geo_confiance": 100,
            "geo_zones_detectees": "; ".join(zones),
            "geo_departements_detectes": "; ".join(codes),
            "geo_villes_detectees": "; ".join(villes),
            "geo_hors_cible": "; ".join(hors),
            "geo_motif": "Lieu d'événement explicitement hors des 7 départements cibles.",
        }

    if codes and any(c in DEPARTEMENTS_CIBLES for c in codes):
        return {
            "geo_preanalyse": "GEO_CIBLE_CONFIRMEE",
            "geo_confiance": 95,
            "geo_zones_detectees": "; ".join(zones),
            "geo_departements_detectes": "; ".join(codes),
            "geo_villes_detectees": "; ".join(villes),
            "geo_hors_cible": "",
            "geo_motif": "Lieu explicitement rattaché à un département cible.",
        }

    dept_requete = str(resultat.get("departement", "")).strip()
    if dept_requete in DEPARTEMENTS_CIBLES:
        return {
            "geo_preanalyse": "GEO_AMBIGUE",
            "geo_confiance": 55,
            "geo_zones_detectees": "; ".join(zones),
            "geo_departements_detectes": "; ".join(codes),
            "geo_villes_detectees": "; ".join(villes),
            "geo_hors_cible": "",
            "geo_motif": f"Le département de recherche est cible ({dept_requete}) mais le lieu précis n'est pas confirmé.",
        }

    return {
        "geo_preanalyse": "GEO_INCONNUE",
        "geo_confiance": 0,
        "geo_zones_detectees": "; ".join(zones),
        "geo_departements_detectes": "; ".join(codes),
        "geo_villes_detectees": "; ".join(villes),
        "geo_hors_cible": "",
        "geo_motif": "Aucun lieu exploitable identifié.",
    }


def _appel_modele(client, modele, prompt):
    response = client.models.generate_content(model=modele, contents=prompt)
    texte = (response.text or "").strip()
    if texte.startswith("```"):
        texte = texte.replace("```json", "").replace("```", "").strip()
    return json.loads(texte)


def _normaliser_decision(data):
    decision = str(data.get("decision", "A_VERIFIER")).upper().strip()
    if decision not in {"PERTINENT", "A_VERIFIER", "NON_PERTINENT"}:
        decision = "A_VERIFIER"
    try:
        confiance = int(data.get("confiance", 0))
    except (TypeError, ValueError):
        confiance = 0
    return decision, max(0, min(100, confiance))


def _base_indisponible(message, modele, suffix=""):
    return {
        f"decision_ia{suffix}": "IA_INDISPONIBLE",
        f"confiance_ia{suffix}": 0,
        f"raison_ia{suffix}": message,
        f"modele_ia{suffix}": modele,
    }


def diagnostic_gemini():
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {"ok": False, "message": "GEMINI_API_KEY absente."}
    try:
        client = genai.Client(api_key=api_key)
        result = {}
        for label, model in (("stage31", MODEL_31), ("stage38", MODEL_38)):
            try:
                obj = client.models.get(model=model)
                result[label] = {"ok": True, "code": 200, "message": getattr(obj, "name", model)}
            except Exception as exc:
                result[label] = {"ok": False, "code": getattr(exc, "code", 0), "message": str(exc)[:500]}
        result["ok"] = result["stage31"]["ok"]
        return result
    except Exception as exc:
        return {"ok": False, "message": f"{type(exc).__name__}: {str(exc)[:500]}"}


def _contexte(resultat):
    return f"""
DONNÉES DE L'ÉVÉNEMENT
Titre : {resultat.get('titre','')}
Description : {resultat.get('description','')}
Source : {resultat.get('source','')}
URL : {resultat.get('url','')}
Département : {resultat.get('departement','')}
Année recherchée : {resultat.get('annee_recherchee','')}
Année détectée : {resultat.get('annee_detectee','')}
Date événement : {resultat.get('date_evenement_page','')}
Lieu : {resultat.get('lieu_page','')}
Organisateur : {resultat.get('organisateur_page','')}
Titre page : {resultat.get('titre_page','')}
Description page : {resultat.get('description_page','')}
Extrait page : {resultat.get('extrait_page','')}
Contenu page : {str(resultat.get('contenu_page',''))[:12000]}

DONNÉES PYTHON
Score : {resultat.get('score_cible_apec', resultat.get('score',0))}
Signaux Apec : {resultat.get('signaux_cible_apec','')}
Signaux supérieur : {resultat.get('signaux_superieur','')}
Signaux entreprises : {resultat.get('signaux_entreprise','')}
Signaux exclusion : {resultat.get('signaux_exclusion','')}
Priorité 2027 : {resultat.get('priorite_2027','')}
Pré-analyse géographique : {resultat.get('geo_preanalyse','')} | confiance {resultat.get('geo_confiance','')}
Zones détectées : {resultat.get('geo_zones_detectees','')}
Hors cible détecté : {resultat.get('geo_hors_cible','')}
"""


def _regles():
    return """
Tu analyses une veille événementielle destinée à l'Apec.
Zone : départements 16, 17, 19, 23, 79, 86, 87. Toute localisation d'événement explicitement hors de ces 7 départements est éliminatoire. Une ville citée comme origine du public, siège de l'organisateur ou partenaire ne doit pas être confondue avec le lieu de l'événement.
Priorité : 2027, puis 2026.

Cible : cadres, managers, experts, dirigeants, jeunes diplômés Bac+3/Bac+5+, masters,
IAE, universités, écoles d'ingénieurs, écoles de commerce, grandes écoles, alumni,
doctorants/docteurs, recrutement de cadres, mobilité, évolution, reconversion et
relations entreprises lorsqu'elles concernent ces publics.

Un simple mot « emploi », « étudiant », « formation » ou « recrutement » ne suffit pas.
Une décision PERTINENT doit reposer sur des éléments concrets. En l'absence de preuve
suffisante, préférer A_VERIFIER. Les publics lycée/collège, jobs étudiants, jobs d'été,
saisonniers, petits boulots, intérim généraliste sans public qualifié et événements
purement grand public sont à écarter.
Ne rien inventer.
"""


def _prompt_t1(resultat):
    return _regles() + """

TOUR 1 — ANALYSTE
Reconstruis l'événement comme le ferait un analyste humain. Identifie :
- ce qu'est réellement l'événement ;
- son public réel ;
- son niveau de qualification ;
- sa dimension entreprises/recrutement ;
- les preuves disponibles ;
- l'intérêt Apec ;
- les éléments manquants ;
- la qualité de la source ;
- le niveau de preuve de 0 à 4 ;
- la force de la preuve de cible Apec.

Retourne uniquement ce JSON :
{
  "decision":"PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance":0,
  "raison":"",
  "public_cible":"",
  "type_evenement":"",
  "niveau_qualification":"",
  "preuve_decisive":"",
  "preuves":"",
  "niveau_preuve":0,
  "preuve_cible":"FORTE|MOYENNE|FAIBLE|ABSENTE",
  "qualite_source":"OFFICIELLE|INSTITUTIONNELLE|PRESSE|AGREGATEUR|INCONNUE",
  "contradictions":"",
  "action_recommandee":""
}
""" + _contexte(resultat)


def _prompt_t2(resultat, t1):
    return _regles() + f"""

TOUR 2 — CONTRE-ANALYSTE
Tu dois chercher activement les erreurs possibles dans le TOUR 1. Ne cherche pas à
confirmer la première conclusion. Vérifie en particulier : public réellement visé,
confusion lycée/supérieur, événement générique, date, géographie, nature éditoriale,
absence de preuve, contradictions entre les éléments.

TOUR 1 :
{json.dumps(t1, ensure_ascii=False, indent=2)}

{_contexte(resultat)}

Retourne uniquement ce JSON :
{{
  "decision":"PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance":0,
  "raison":"",
  "erreur_potentielle":"",
  "contre_preuve":"",
  "preuve_confirmee":"",
  "contradiction":"",
  "recommandation_t3":""
}}
"""


def _prompt_t3(resultat, t1, t2):
    return _regles() + f"""

TOUR 3 — ARBITRE 3.1 / IL FAUT TRANCHER
Tu es le responsable de veille. Tu disposes des deux analyses précédentes.
Tu dois prendre une décision finale 3.1, même si les analyses divergent.
Ne fais pas une moyenne mécanique : pèse la qualité des preuves et les exclusions.
Une preuve explicite et fiable prime sur un simple signal lexical.

TOUR 1 :
{json.dumps(t1, ensure_ascii=False, indent=2)}

TOUR 2 :
{json.dumps(t2, ensure_ascii=False, indent=2)}

{_contexte(resultat)}

Retourne uniquement ce JSON :
{{
  "decision":"PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance":0,
  "raison":"",
  "preuve_decisive":"",
  "contre_argument_principal":"",
  "public_final":"",
  "type_evenement_final":"",
  "action_recommandee":"",
  "arbitrage":"CONFIRME_T1|CONFIRME_T2|TRANCHE_ENTRE_LES_DEUX"
}}
"""


def analyser_resultat_31(resultat):
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return _base_indisponible("Clé GEMINI_API_KEY absente.", MODEL_31, "_1")
    try:
        client = genai.Client(api_key=api_key)
        data = _appel_modele(client, MODEL_31, _prompt_t1(resultat))
        decision, confiance = _normaliser_decision(data)
        return {
            "modele_ia_1": MODEL_31,
            "decision_ia_1": decision,
            "confiance_ia_1": confiance,
            "raison_ia_1": data.get("raison", ""),
            "public_cible_31": data.get("public_cible", ""),
            "type_evenement_31": data.get("type_evenement", ""),
            "niveau_qualification_31": data.get("niveau_qualification", ""),
            "preuve_decisive_31": data.get("preuve_decisive", ""),
            "preuves_31": data.get("preuves", ""),
            "niveau_preuve_31": max(0, min(4, int(data.get("niveau_preuve", 0) or 0))),
            "preuve_cible_31": data.get("preuve_cible", ""),
            "qualite_source_31": data.get("qualite_source", ""),
            "contradictions_31": data.get("contradictions", ""),
            "action_recommandee_31": data.get("action_recommandee", ""),
        }
    except genai_errors.APIError as exc:
        return _base_indisponible(f"Gemini 3.1 HTTP {getattr(exc,'code','?')} : {str(exc)[:500]}", MODEL_31, "_1")
    except Exception as exc:
        return _base_indisponible(f"Erreur Gemini 3.1 T1 : {type(exc).__name__} : {str(exc)[:500]}", MODEL_31, "_1")


def analyser_resultat_31_tour2(resultat, t1):
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return _base_indisponible("Clé GEMINI_API_KEY absente.", MODEL_31, "_2")
    try:
        client = genai.Client(api_key=api_key)
        data = _appel_modele(client, MODEL_31, _prompt_t2(resultat, t1))
        decision, confiance = _normaliser_decision(data)
        return {
            "modele_ia_31_t2": MODEL_31,
            "decision_ia_31_t2": decision,
            "confiance_ia_31_t2": confiance,
            "raison_ia_31_t2": data.get("raison", ""),
            "erreur_potentielle_31_t2": data.get("erreur_potentielle", ""),
            "contre_preuve_31_t2": data.get("contre_preuve", ""),
            "preuve_confirmee_31_t2": data.get("preuve_confirmee", ""),
            "contradiction_31_t2": data.get("contradiction", ""),
            "recommandation_t3_31": data.get("recommandation_t3", ""),
        }
    except genai_errors.APIError as exc:
        return _base_indisponible(f"Gemini 3.1 HTTP {getattr(exc,'code','?')} : {str(exc)[:500]}", MODEL_31, "_2")
    except Exception as exc:
        return _base_indisponible(f"Erreur Gemini 3.1 T2 : {type(exc).__name__} : {str(exc)[:500]}", MODEL_31, "_2")


def analyser_resultat_31_tour3(resultat, t1, t2):
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return _base_indisponible("Clé GEMINI_API_KEY absente.", MODEL_31, "_3")
    try:
        client = genai.Client(api_key=api_key)
        data = _appel_modele(client, MODEL_31, _prompt_t3(resultat, t1, t2))
        decision, confiance = _normaliser_decision(data)
        return {
            "modele_ia_31_t3": MODEL_31,
            "decision_ia_31_t3": decision,
            "confiance_ia_31_t3": confiance,
            "raison_ia_31_t3": data.get("raison", ""),
            "preuve_decisive_31_t3": data.get("preuve_decisive", ""),
            "contre_argument_31_t3": data.get("contre_argument_principal", ""),
            "public_final_31_t3": data.get("public_final", ""),
            "type_evenement_final_31_t3": data.get("type_evenement_final", ""),
            "action_recommandee_31_t3": data.get("action_recommandee", ""),
            "arbitrage_31_t3": data.get("arbitrage", ""),
        }
    except genai_errors.APIError as exc:
        return _base_indisponible(f"Gemini 3.1 HTTP {getattr(exc,'code','?')} : {str(exc)[:500]}", MODEL_31, "_3")
    except Exception as exc:
        return _base_indisponible(f"Erreur Gemini 3.1 T3 : {type(exc).__name__} : {str(exc)[:500]}", MODEL_31, "_3")


def _prompt_38(resultat):
    return _regles() + f"""

GEMINI 3.8 — EXPERT SENIOR / ARBITRAGE FINAL
Tu interviens uniquement sur les candidats sélectionnés après trois analyses 3.1.
Tu dois examiner l'ensemble du dossier et déterminer si la décision 3.1 finale est
justifiée. Tu peux confirmer, rétrograder ou rejeter. Ne rien inventer.
La décision 3.1 T3 reste la décision de base : ton champ « arbitrage » est prioritaire
sur « decision_finale ». Utilise CONFIRMER pour conserver T3, RETROGRADER pour
faire descendre un PERTINENT vers A_VERIFIER, et REJETER pour passer à
NON_PERTINENT. Une simple différence de formulation ne suffit pas : il faut une
preuve ou une contradiction identifiable.

DOSSIER :
{_contexte(resultat)}

TOUR 1 : {json.dumps({k:v for k,v in resultat.items() if k.endswith('_31') or k.endswith('_31_t2')}, ensure_ascii=False)[:12000]}
TOUR 2 : {json.dumps({k:v for k,v in resultat.items() if '31_t2' in k}, ensure_ascii=False)[:10000]}
TOUR 3 : {json.dumps({k:v for k,v in resultat.items() if '31_t3' in k or k=='decision_ia'}, ensure_ascii=False)[:10000]}

Retourne uniquement ce JSON :
{{
  "arbitrage":"CONFIRMER|RETROGRADER|REJETER",
  "decision_finale":"PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance":0,
  "raison":"",
  "preuve_decisive":"",
  "analyse_humaine":"",
  "action_recommandee":"",
  "apport_38":""
}}
"""


def analyser_resultat_38(resultat, analyse31=None):
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {"modele_ia_38": MODEL_38, "arbitrage_38":"AUCUN", "decision_ia_38":"IA_INDISPONIBLE", "confiance_ia_38":0, "raison_ia_38":"Clé GEMINI_API_KEY absente."}
    try:
        client = genai.Client(api_key=api_key)
        data = _appel_modele(client, MODEL_38, _prompt_38(resultat))
        decision, confiance = _normaliser_decision({"decision": data.get("decision_finale"), "confiance": data.get("confiance",0)})
        return {
            "modele_ia_38": MODEL_38,
            "arbitrage_38": str(data.get("arbitrage","AUCUN")).upper(),
            "decision_ia_38": decision,
            "confiance_ia_38": confiance,
            "raison_ia_38": data.get("raison",""),
            "preuve_decisive_38": data.get("preuve_decisive",""),
            "analyse_humaine_38": data.get("analyse_humaine",""),
            "action_recommandee_38": data.get("action_recommandee",""),
            "apport_38": data.get("apport_38",""),
        }
    except genai_errors.APIError as exc:
        return {"modele_ia_38": MODEL_38, "arbitrage_38":"AUCUN", "decision_ia_38":"IA_INDISPONIBLE", "confiance_ia_38":0, "raison_ia_38":f"Gemini 3.8 HTTP {getattr(exc,'code','?')} : {str(exc)[:500]}"}
    except Exception as exc:
        return {"modele_ia_38": MODEL_38, "arbitrage_38":"AUCUN", "decision_ia_38":"IA_INDISPONIBLE", "confiance_ia_38":0, "raison_ia_38":f"Erreur Gemini 3.8 : {type(exc).__name__} : {str(exc)[:500]}"}


def analyser_resultat(resultat):
    return analyser_resultat_31(resultat)
