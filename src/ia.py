
import os
import json
from google import genai
from google.genai import errors as genai_errors

MODEL_PRIMARY = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
MODEL_FALLBACK = os.getenv("GEMINI_FALLBACK_MODEL", "gemini-3.1-flash-lite")
# Alias de compatibilite pour le reste du programme.
MODEL = MODEL_PRIMARY
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
        primary = _diagnostic_modele(client, MODEL_PRIMARY)
        fallback = _diagnostic_modele(client, MODEL_FALLBACK)
        return {"ok": primary.get("ok") or fallback.get("ok"), "primary": primary, "fallback": fallback}
    except Exception as e:
        err = {"ok": False, "code": 0, "message": f"{type(e).__name__}: {str(e)[:500]}"}
        return {"ok": False, "primary": err, "fallback": err}


def _doit_basculer_fallback(code):
    # Un 429 signifie generalement quota/rate-limit : changer de modele ne
    # resout pas forcement le probleme et risquerait de multiplier les appels.
    if code in (401, 429):
        return False
    return code in (0, 400, 403, 404, 408, 409, 500, 502, 503, 504)


def _appel_modele(client, modele, prompt):
    response = client.models.generate_content(model=modele, contents=prompt)
    texte = (response.text or "").strip()
    if texte.startswith("```"):
        texte = texte.replace("```json", "").replace("```", "").strip()
    return json.loads(texte)

def analyser_resultat(resultat):
    """Analyse finale Apec : cadres, jeunes diplômés et enseignement supérieur qualifié."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {
            "decision_ia": "IA_INDISPONIBLE",
            "confiance_ia": 0,
            "raison_ia": "Clé GEMINI_API_KEY absente.",
            "categorie_ia": "",
            "date_ia": "",
            "departement_ia": "",
            "modele_ia": MODEL_PRIMARY,
            "fallback_ia": "NON",
            "erreur_modele_primaire": "",
        }

    global _DIAGNOSTIC_FAIT

    if not _DIAGNOSTIC_FAIT:
        _DIAGNOSTIC_FAIT = True
        print("========================================", flush=True)
        print("TEST GEMINI API", flush=True)
        print("========================================", flush=True)
        print(f"Clé GEMINI_API_KEY : présente", flush=True)
        print(f"Modèle principal : {MODEL_PRIMARY}", flush=True)
        print(f"Modèle secours   : {MODEL_FALLBACK}", flush=True)
        diagnostic = diagnostic_gemini()
        pdiag = diagnostic.get("primary", {})
        fdiag = diagnostic.get("fallback", {})
        print(f"✓ {MODEL_PRIMARY} : {pdiag.get('message', 'OK')}" if pdiag.get("ok") else f"✗ {MODEL_PRIMARY} : code {pdiag.get('code', 0)} — {pdiag.get('message', '')}", flush=True)
        print(f"✓ {MODEL_FALLBACK} : {fdiag.get('message', 'OK')}" if fdiag.get("ok") else f"✗ {MODEL_FALLBACK} : code {fdiag.get('code', 0)} — {fdiag.get('message', '')}", flush=True)
        if not diagnostic.get("ok"):
            return {
                "decision_ia": "IA_INDISPONIBLE",
                "confiance_ia": 0,
                "raison_ia": "TEST GEMINI API — aucun des deux modèles n’est accessible.",
                "categorie_ia": "",
                "date_ia": "",
                "departement_ia": "",
                "modele_ia": "",
                "fallback_ia": "NON",
            }
        print("========================================", flush=True)

    try:
        client = genai.Client(api_key=api_key)

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

        prompt = f"""
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

        modele_utilise = MODEL_PRIMARY
        fallback_utilise = "NON"
        erreur_primaire = ""

        try:
            data = _appel_modele(client, MODEL_PRIMARY, prompt)
            print(f"→ {MODEL_PRIMARY} : OK", flush=True)
        except genai_errors.APIError as e_primary:
            code_primary = getattr(e_primary, "code", 0)
            msg_primary = getattr(e_primary, "message", str(e_primary))
            erreur_primaire = f"HTTP {code_primary} : {str(msg_primary)[:300]}"
            print(f"→ {MODEL_PRIMARY} : ERREUR {code_primary} — {str(msg_primary)[:180]}", flush=True)
            if not _doit_basculer_fallback(code_primary):
                raise
            print(f"→ bascule {MODEL_FALLBACK}", flush=True)
            data = _appel_modele(client, MODEL_FALLBACK, prompt)
            modele_utilise = MODEL_FALLBACK
            fallback_utilise = "OUI"
            print(f"→ {MODEL_FALLBACK} : OK", flush=True)
        except Exception as e_primary:
            erreur_primaire = f"{type(e_primary).__name__} : {str(e_primary)[:300]}"
            print(f"→ {MODEL_PRIMARY} : ERREUR — {str(e_primary)[:180]}", flush=True)
            print(f"→ bascule {MODEL_FALLBACK}", flush=True)
            data = _appel_modele(client, MODEL_FALLBACK, prompt)
            modele_utilise = MODEL_FALLBACK
            fallback_utilise = "OUI"
            print(f"→ {MODEL_FALLBACK} : OK", flush=True)

        decision = data.get("decision", "A_VERIFIER")
        if decision not in {"PERTINENT", "A_VERIFIER", "NON_PERTINENT"}:
            decision = "A_VERIFIER"

        try:
            confiance = int(data.get("confiance", 0))
        except (TypeError, ValueError):
            confiance = 0

        confiance = max(0, min(100, confiance))

        return {
            "decision_ia": decision,
            "confiance_ia": confiance,
            "raison_ia": data.get("raison", ""),
            "categorie_ia": data.get("categorie", ""),
            "date_ia": data.get("date", ""),
            "departement_ia": data.get("departement", ""),
            "modele_ia": modele_utilise,
            "fallback_ia": fallback_utilise,
            "erreur_modele_primaire": erreur_primaire,
        }

    except genai_errors.APIError as e:
        code = getattr(e, "code", "?")
        message = getattr(e, "message", str(e))
        return {
            "decision_ia": "IA_INDISPONIBLE",
            "confiance_ia": 0,
            "raison_ia": f"Gemini API HTTP {code} ({type(e).__name__}) : {str(message)[:500]}",
            "categorie_ia": "",
            "date_ia": "",
            "departement_ia": "",
            "modele_ia": MODEL_PRIMARY,
            "fallback_ia": "NON",
            "erreur_modele_primaire": "",
        }
    except Exception as e:
        return {
            "decision_ia": "IA_INDISPONIBLE",
            "confiance_ia": 0,
            "raison_ia": f"Erreur IA : {type(e).__name__} : {str(e)[:500]}",
            "categorie_ia": "",
            "date_ia": "",
            "departement_ia": "",
        }
