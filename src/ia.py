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
    champs = {
        "titre": resultat.get("titre", ""),
        "description": resultat.get("description", ""),
        "source": resultat.get("source", ""),
        "url": resultat.get("url", ""),
        "date": resultat.get("date_publication", ""),
        "annee_recherchee": resultat.get("annee_recherchee", ""),
        "annee_detectee": resultat.get("annee_detectee", ""),
        "departement": resultat.get("departement", ""),
        "evenement": resultat.get("evenement_detecte", ""),
        "geographie": resultat.get("geographie_verifiee", ""),
        "score": resultat.get("score_cible_apec", resultat.get("score", 0)),
        "signaux_cible": resultat.get("signaux_cible_apec", ""),
        "signaux_superieur": resultat.get("signaux_superieur", ""),
        "signaux_entreprise": resultat.get("signaux_entreprise", ""),
        "signaux_exclusion": resultat.get("signaux_exclusion", ""),
        "titre_page": resultat.get("titre_page", ""),
        "description_page": resultat.get("description_page", ""),
        "date_page": resultat.get("date_evenement_page", ""),
        "lieu_page": resultat.get("lieu_page", ""),
        "organisateur_page": resultat.get("organisateur_page", ""),
    }
    return f"""
Tu es l'analyste final d'une veille événementielle Apec.

CIBLE PRIORITAIRE
L'Apec s'intéresse principalement :
- aux cadres en poste ou en recherche ;
- aux managers, experts et cadres dirigeants ;
- aux jeunes diplômés Bac+3/Bac+5 et plus ;
- aux étudiants de l'enseignement supérieur proches du diplôme ;
- aux masters, écoles d'ingénieurs, écoles de commerce, grandes écoles,
  universités, IAE, doctorants, docteurs et alumni ;
- aux relations entreprises/recruteurs lorsqu'elles concernent ces publics.

ZONE : 86, 79, 87, 17, 16, 23, 19.
ANNÉES : 2025, 2026, 2027.

PRINCIPE ESSENTIEL
Ne demande jamais que les mots « cadre » ou « jeune diplômé » soient présents.
La pertinence peut être démontrée indirectement par le contexte.

PERTINENT si l'un de ces cas est suffisamment démontré :
1. cadres, managers, experts, dirigeants ou jeunes diplômés explicitement ciblés ;
2. université/IAE/école supérieure + entreprises/employeurs/recruteurs +
   étudiants/diplômés ;
3. école d'ingénieurs, école de commerce ou grande école + entreprises,
   recrutement, carrière ou insertion ;
4. master/Bac+5/doctorat/ingénieurs/alumni + entreprises, recrutement,
   carrière ou insertion ;
5. rencontre entreprises, forum entreprises ou job dating dans un
   établissement supérieur avec public qualifié ;
6. mobilité, évolution professionnelle ou reconversion destinée à des actifs
   qualifiés/cadres ;
7. événement Apec ou Apec Connect.

EXEMPLES
- « Forum entreprises – Université de Poitiers » peut être PERTINENT même sans
  le mot « cadre » si la relation université + entreprises + étudiants/diplômés
  est clairement établie.
- « Rencontre entreprises / étudiants – Université X » peut être PERTINENT.
- « Forum entreprises – école d'ingénieurs » est généralement PERTINENT.

A_VERIFIER lorsque le potentiel Apec est réel mais qu'une information
importante manque :
- forum emploi avec employeurs mais public non précisé ;
- événement universitaire avec entreprises mais niveau de diplôme non précisé ;
- événement carrière/recrutement dont le public qualifié est plausible mais
  pas suffisamment documenté.

NON_PERTINENT si :
- jobs étudiants, jobs d'été ou saisonniers ;
- petits boulots ;
- restauration, caisse, manutention, préparation de commandes ;
- intérim généraliste sans public qualifié ;
- emplois sans qualification ;
- collège, lycée, post-3e ;
- orientation générale pré-Bac+3 sans dimension Apec ;
- événement purement commercial ou grand public sans intérêt professionnel ;
- CCI comme événement ;
- événement non identifiable ;
- article éditorial qui ne permet pas d'identifier l'événement ;
- hors zone lorsque cela est démontré.

IMPORTANT
Un simple mot « emploi », « formation », « étudiant » ou « recrutement »
ne suffit pas.
Mais une combinaison cohérente de signaux doit être valorisée.

En cas de doute entre NON_PERTINENT et un intérêt Apec plausible,
choisis A_VERIFIER. Ne transforme pas une absence d'information en preuve
d'absence.

DONNÉES PYTHON
Score : {champs["score"]}
Cible directe : {champs["signaux_cible"]}
Enseignement supérieur : {champs["signaux_superieur"]}
Relation entreprises : {champs["signaux_entreprise"]}
Exclusions : {champs["signaux_exclusion"]}

SOURCE RSS
Titre : {champs["titre"]}
Description : {champs["description"]}
Source : {champs["source"]}
Date publication : {champs["date"]}
Année recherchée : {champs["annee_recherchee"]}
Année détectée : {champs["annee_detectee"]}
Département : {champs["departement"]}
Événement détecté : {champs["evenement"]}
Géographie vérifiée : {champs["geographie"]}
URL : {champs["url"]}

PAGE ENRICHIE
Titre : {champs["titre_page"]}
Description : {champs["description_page"]}
Date événement : {champs["date_page"]}
Lieu : {champs["lieu_page"]}
Organisateur : {champs["organisateur_page"]}

DÉCISION
PERTINENT = intérêt Apec clairement ou fortement démontré.
A_VERIFIER = potentiel réel mais preuve insuffisante.
NON_PERTINENT = hors cible ou exclusion claire.

Ne jamais inventer une date, un lieu, un public, un diplôme ou un organisateur.
La confiance est un entier de 0 à 100.

Réponds UNIQUEMENT avec ce JSON :
{{
  "decision": "PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance": 0,
  "raison": "explication courte et factuelle",
  "categorie": "emploi|recrutement|formation|enseignement_superieur|alumni|portes_ouvertes|reconversion|entreprises_rh|autre",
  "date": "",
  "departement": ""
}}
"""

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
    return f"""Tu es le second niveau d'analyse d'une veille événementielle Apec.

Tu dois RÉ-ENRICHIR l'analyse Gemini 3.1 ci-dessous, pas la remplacer mécaniquement. Vérifie les éléments disponibles et apporte uniquement des précisions justifiées. Ne rien inventer.

ANALYSE GEMINI 3.1
Décision : {analyse31.get('decision_ia_1', analyse31.get('decision_ia',''))}
Confiance : {analyse31.get('confiance_ia_1', analyse31.get('confiance_ia',0))}
Raison : {analyse31.get('raison_ia_1', analyse31.get('raison_ia',''))}

DONNÉES ÉVÉNEMENT
Titre : {resultat.get('titre','')}
Description : {resultat.get('description','')}
Source : {resultat.get('source','')}
URL : {resultat.get('url','')}
Département : {resultat.get('departement','')}
Année : {resultat.get('annee_detectee', resultat.get('annee_recherchee',''))}
Score Python : {resultat.get('score_cible_apec', resultat.get('score',0))}
Signaux cible : {resultat.get('signaux_cible_apec','')}
Signaux supérieur : {resultat.get('signaux_superieur','')}
Signaux entreprises : {resultat.get('signaux_entreprise','')}
Page : {resultat.get('titre_page','')} | {resultat.get('description_page','')}
Date événement : {resultat.get('date_evenement_page','')}
Lieu : {resultat.get('lieu_page','')}
Organisateur : {resultat.get('organisateur_page','')}

OBJECTIF
Renforce l'analyse sur : public visé, niveau de qualification, relation entreprises/recrutement, intérêt cadres/jeunes diplômés, importance Apec et priorité 2027.

Retourne UNIQUEMENT ce JSON :
{{
  "decision": "PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance": 0,
  "raison": "raison finale enrichie et factuelle",
  "enrichissement": "ce que le 3.8 apporte par rapport au 3.1",
  "priorite_38": "TRES_FORTE|FORTE|MOYENNE"
}}"""


def analyser_resultat_38(resultat, analyse31):
    """Ré-enrichissement des meilleurs candidats par Gemini 3.8 Flash."""
    api_key=os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {"modele_ia_2":MODEL_38,"decision_ia_2":"IA_INDISPONIBLE","confiance_ia_2":0,"raison_ia_2":"Clé GEMINI_API_KEY absente.","enrichissement_38":""}
    try:
        client=genai.Client(api_key=api_key)
        data=_appel_modele(client, MODEL_38, _prompt_38(resultat, analyse31))
        decision, confiance=_normaliser_decision(data)
        return {"modele_ia_2":MODEL_38,"decision_ia_2":decision,"confiance_ia_2":confiance,"raison_ia_2":data.get("raison",""),"enrichissement_38":data.get("enrichissement",""),"priorite_38":data.get("priorite_38","")}
    except genai_errors.APIError as e:
        return {"modele_ia_2":MODEL_38,"decision_ia_2":"IA_INDISPONIBLE","confiance_ia_2":0,"raison_ia_2":f"Gemini 3.8 HTTP {getattr(e,'code','?')} : {str(getattr(e,'message',str(e)))[:500]}","enrichissement_38":""}
    except Exception as e:
        return {"modele_ia_2":MODEL_38,"decision_ia_2":"IA_INDISPONIBLE","confiance_ia_2":0,"raison_ia_2":f"Erreur Gemini 3.8 : {type(e).__name__} : {str(e)[:500]}","enrichissement_38":""}


def analyser_resultat(resultat):
    return analyser_resultat_31(resultat)
