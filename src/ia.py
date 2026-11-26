import os
import json
from google import genai
from google.genai import errors as genai_errors

MODEL_31 = os.getenv("GEMINI_MODEL_31", "gemini-3.1-flash-lite")
MODEL_38 = os.getenv("GEMINI_MODEL_38", "gemini-3.8-flash")


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
"""


def _regles():
    return """
Tu analyses une veille événementielle destinée à l'Apec.
Zone : départements 16, 17, 19, 23, 79, 86, 87.
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
- les éléments manquants.

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
