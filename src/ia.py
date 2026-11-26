import os
import json
from google import genai
from google.genai import errors as genai_errors

MODEL_31 = os.getenv("GEMINI_MODEL_31", "gemini-3.1-flash-lite")
MODEL_38 = os.getenv("GEMINI_MODEL_38", "gemini-3.8-flash")
MODEL = MODEL_31
_DIAGNOSTIC_FAIT = False


def _diagnostic_modele(client, modele):
    try:
        obj = client.models.get(model=modele)
        return {"ok": True, "code": 200, "message": f"Modèle accessible : {getattr(obj, "name", modele)}"}
    except genai_errors.APIError as e:
        return {"ok": False, "code": getattr(e, "code", 0), "message": getattr(e, "message", str(e))[:500]}
    except Exception as e:
        return {"ok": False, "code": 0, "message": f"{type(e).__name__}: {str(e)[:500]}"}


def diagnostic_gemini():
    """Teste l'accessibilite des deux modeles sans lancer de generation metier."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {
            "ok": False,
            "primary": {"ok": False, "code": 0, "message": "GEMINI_API_KEY absente."},
            "fallback": {"ok": False, "code": 0, "message": "GEMINI_API_KEY absente."},
        }
    try:
        client = genai.Client(api_key=api_key)
        stage31 = _diagnostic_modele(client, MODEL_31)
        stage38 = _diagnostic_modele(client, MODEL_38)
        return {"ok": stage31.get("ok") or stage38.get("ok"), "stage31": stage31, "stage38": stage38}
    except Exception as e:
        err = {"ok": False, "code": 0, "message": f"{type(e).__name__}: {str(e)[:500]}"}
        return {"ok": False, "primary": err, "fallback": err}


def _appel_modele(client, modele, prompt):
    response = client.models.generate_content(model=modele, contents=prompt)
    texte = (response.text or "").strip()
    if texte.startswith("```"):
        texte = texte.replace("```json", "").replace("```", "").strip()
    return json.loads(texte)

def _construire_prompt_31(resultat):
    return f"""Tu es le filtre qualité d'une veille événementielle Apec.

OBJECTIF
Décider si la page décrit un VRAI événement exploitable pour une veille Apec, dans les départements 86,79,87,17,16,23,19 et les années 2026-2027, priorité 2027.

RÈGLE ABSOLUE DE CIBLE
Un simple mot-clé comme emploi, formation, étudiant, université, recrutement ou forum ne prouve PAS la cible Apec. Tu dois identifier QUI est réellement visé et citer la preuve disponible.

PERTINENT seulement si :
- cadres/managers/dirigeants/jeunes diplômés explicitement ciblés ; OU
- enseignement supérieur + entreprises/recruteurs + public supérieur clairement identifiable ; OU
- Bac+3/Bac+5/master/ingénieurs/école/grande école/IAE/alumni + entreprises/recrutement/carrière ; OU
- mobilité, évolution professionnelle ou reconversion destinée à des actifs qualifiés/cadres ; OU
- événement Apec/Apec Connect.

A_VERIFIER si l'événement est crédible mais que le public qualifié, la relation recrutement ou une autre preuve essentielle manque.

NON_PERTINENT si public lycée/collège, job étudiant/saisonnier, petits boulots, intérim généraliste, emplois sans qualification, article éditorial sans événement identifiable, hors zone, ou exclusion claire.

NE JAMAIS INVENTER. Si l'information n'est pas dans les données, écris INCONNU.

DONNÉES STRUCTURÉES
Score Python : {resultat.get('score',0)}
Statut Python : {resultat.get('statut','')}
Preuve cible Python : {resultat.get('preuve_cible','')}
Public cible Python : {resultat.get('public_cible','')}
Niveau preuve Python : {resultat.get('niveau_preuve',0)}
Type événement Python : {resultat.get('type_evenement_v25','')}
Motif exclusion Python : {resultat.get('motif_exclusion_v25','')}
Signaux cible : {resultat.get('signaux_cible_apec','')}
Signaux supérieur : {resultat.get('signaux_superieur','')}
Signaux entreprises : {resultat.get('signaux_entreprise','')}
Signaux exclusion : {resultat.get('signaux_exclusion','')}
Détails score : {resultat.get('details_score','')}

SOURCE
Titre : {resultat.get('titre','')}
Description : {resultat.get('description','')}
Source : {resultat.get('source','')}
URL : {resultat.get('url','')}
Département : {resultat.get('departement','')}
Année détectée : {resultat.get('annee_detectee','')}

PAGE ENRICHIE
Titre : {resultat.get('titre_page','')}
Description : {resultat.get('description_page','')}
Date événement : {resultat.get('date_evenement_page','')}
Lieu : {resultat.get('lieu_page','')}
Organisateur : {resultat.get('organisateur_page','')}

Réponds UNIQUEMENT avec ce JSON :
{{
  "decision": "PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance": 0,
  "raison": "une phrase factuelle indiquant le public et la preuve ou l'absence de preuve",
  "categorie": "emploi|recrutement|formation|enseignement_superieur|alumni|portes_ouvertes|reconversion|entreprises_rh|autre",
  "public_cible": "CADRES|JEUNES_DIPLOMES|BAC+5|BAC+3|ETUDIANTS_SUPERIEUR|ALUMNI|LYCEENS|JOB_ETUDIANT|GRAND_PUBLIC|INCONNU",
  "preuve_cible": "FORTE|MOYENNE|FAIBLE|ABSENTE",
  "type_evenement": "RECRUTEMENT|CAREER|FORUM_ENTREPRISES|FORMATION_SUPERIEURE|ALUMNI|PORTES_OUVERTES_SUPERIEUR|JOB_DATING|SALON_EMPLOI|SALON_ETUDIANT|ARTICLE_EDITORIAL|EVENEMENT_INDETERMINE",
  "preuve": "preuve textuelle courte ou INCONNUE",
  "date": "",
  "departement": ""
}}"""

def _normaliser_decision(data):
    decision = data.get("decision", "A_VERIFIER")
    if decision not in {"PERTINENT", "A_VERIFIER", "NON_PERTINENT"}:
        decision = "A_VERIFIER"
    try:
        confiance = int(data.get("confiance", 0))
    except (TypeError, ValueError):
        confiance = 0
    confiance = max(0, min(100, confiance))
    return decision, confiance


def _base_resultat_indisponible(message, modele):
    return {"decision_ia":"IA_INDISPONIBLE","confiance_ia":0,"raison_ia":message,"categorie_ia":"","date_ia":"","departement_ia":"","modele_ia":modele}


def analyser_resultat_31(resultat):
    """Analyse de base : Gemini 3.1 Flash Lite. Aucune limite ici."""
    api_key=os.getenv("GEMINI_API_KEY")
    if not api_key:
        return _base_resultat_indisponible("Clé GEMINI_API_KEY absente.", MODEL_31)
    try:
        client=genai.Client(api_key=api_key)
        data=_appel_modele(client, MODEL_31, _construire_prompt_31(resultat))
        decision, confiance=_normaliser_decision(data)
        return {"decision_ia":decision,"confiance_ia":confiance,"raison_ia":data.get("raison",""),"categorie_ia":data.get("categorie",""),"date_ia":data.get("date",""),"departement_ia":data.get("departement",""),"modele_ia":MODEL_31,"modele_ia_1":MODEL_31,"decision_ia_1":decision,"confiance_ia_1":confiance,"raison_ia_1":data.get("raison","")}
    except genai_errors.APIError as e:
        return _base_resultat_indisponible(f"Gemini 3.1 HTTP {getattr(e,'code','?')} : {str(getattr(e,'message',str(e)))[:500]}", MODEL_31)
    except Exception as e:
        return _base_resultat_indisponible(f"Erreur Gemini 3.1 : {type(e).__name__} : {str(e)[:500]}", MODEL_31)


def _prompt_38(resultat, analyse31):
    return f"""Tu es le contrôleur qualité final d'une veille Apec.

Tu ne dois PAS remplacer mécaniquement Gemini 3.1. Tu dois vérifier ses erreurs potentielles avec les preuves structurées disponibles.

RÈGLE : une promotion en PERTINENT est interdite si la preuve de cible Apec est faible/absente ou si l'événement n'est pas clairement identifiable. Une exclusion doit être justifiée par une preuve concrète.

GEMINI 3.1
Décision : {analyse31.get('decision_ia_1', analyse31.get('decision_ia',''))}
Confiance : {analyse31.get('confiance_ia_1', analyse31.get('confiance_ia',0))}
Raison : {analyse31.get('raison_ia_1', analyse31.get('raison_ia',''))}

PYTHON
Score : {resultat.get('score',0)}
Statut : {resultat.get('statut','')}
Preuve cible : {resultat.get('preuve_cible','')}
Public cible : {resultat.get('public_cible','')}
Niveau preuve : {resultat.get('niveau_preuve',0)}
Type événement : {resultat.get('type_evenement_v25','')}
Exclusion : {resultat.get('motif_exclusion_v25','')}
Signaux : {resultat.get('signaux_cible_apec','')} | {resultat.get('signaux_superieur','')} | {resultat.get('signaux_entreprise','')}

ÉVÉNEMENT
Titre : {resultat.get('titre','')}
Description : {resultat.get('description','')}
Source : {resultat.get('source','')}
URL : {resultat.get('url','')}
Titre page : {resultat.get('titre_page','')}
Description page : {resultat.get('description_page','')}
Date : {resultat.get('date_evenement_page','')}
Lieu : {resultat.get('lieu_page','')}
Organisateur : {resultat.get('organisateur_page','')}

ARBITRAGE
- CONFIRMER : la décision 3.1 est cohérente avec les preuves.
- RETROGRADER : la décision 3.1 est trop positive ; revenir à A_VERIFIER.
- REJETER : preuve forte que le candidat est hors cible ; utiliser seulement si la confiance est >= 80.
- AUCUN : pas assez d'éléments pour modifier 3.1.

Réponds UNIQUEMENT avec ce JSON :
{{
  "arbitrage": "CONFIRMER|RETROGRADER|REJETER|AUCUN",
  "confiance": 0,
  "raison": "constat factuel et court",
  "enrichissement": "élément nouveau ou vérification apportée",
  "priorite_38": "TRES_FORTE|FORTE|MOYENNE|FAIBLE"
}}"""

def analyser_resultat_38(resultat, analyse31):
    """Ré-enrichissement des meilleurs candidats par Gemini 3.8 Flash."""
    api_key=os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {"modele_ia_2":MODEL_38,"decision_ia_2":"IA_INDISPONIBLE","confiance_ia_2":0,"raison_ia_2":"Clé GEMINI_API_KEY absente.","enrichissement_38":""}
    try:
        client=genai.Client(api_key=api_key)
        data=_appel_modele(client, MODEL_38, _prompt_38(resultat, analyse31))
        arbitrage = str(data.get("arbitrage", "AUCUN")).upper()
        if arbitrage not in {"CONFIRMER", "RETROGRADER", "REJETER", "AUCUN"}:
            arbitrage = "AUCUN"
        try:
            confiance = int(data.get("confiance", 0))
        except (TypeError, ValueError):
            confiance = 0
        confiance = max(0, min(100, confiance))
        return {"modele_ia_2":MODEL_38,"decision_ia_2":"ARBITRAGE_OK","arbitrage_38":arbitrage,"confiance_ia_2":confiance,"raison_ia_2":data.get("raison",""),"enrichissement_38":data.get("enrichissement",""),"priorite_38":data.get("priorite_38","")}
    except genai_errors.APIError as e:
        return {"modele_ia_2":MODEL_38,"decision_ia_2":"IA_INDISPONIBLE","confiance_ia_2":0,"raison_ia_2":f"Gemini 3.8 HTTP {getattr(e,'code','?')} : {str(getattr(e,'message',str(e)))[:500]}","enrichissement_38":""}
    except Exception as e:
        return {"modele_ia_2":MODEL_38,"decision_ia_2":"IA_INDISPONIBLE","confiance_ia_2":0,"raison_ia_2":f"Erreur Gemini 3.8 : {type(e).__name__} : {str(e)[:500]}","enrichissement_38":""}


def analyser_resultat(resultat):
    return analyser_resultat_31(resultat)
