import os
import json
import time
from google import genai

MODEL = "gemini-3.1-flash-lite"

# Attente (secondes) avant de retenter après une erreur temporaire 503.
ATTENTES_503 = (20, 45)
# Attente avant de retenter après un quota par minute (429 « PerMinute »).
ATTENTE_QUOTA_MINUTE = 65


def _classer_erreur(exc):
    """Retourne (type, message) : QUOTA_JOUR, QUOTA_MINUTE, TEMPORAIRE ou AUTRE."""
    msg = f"{type(exc).__name__} : {exc}"
    bas = msg.lower()

    if "429" in bas or "resource_exhausted" in bas or "quota" in bas:
        if "perday" in bas or "per day" in bas:
            return "QUOTA_JOUR", msg
        if "perminute" in bas or "per minute" in bas:
            return "QUOTA_MINUTE", msg
        # Quota de nature inconnue : on le traite comme bloquant par prudence.
        return "QUOTA_JOUR", msg

    if any(t in bas for t in ("503", "unavailable", "500", "504", "timeout", "deadline")):
        return "TEMPORAIRE", msg

    return "AUTRE", msg


def _appeler_gemini(client, prompt):
    """Appel Gemini avec reprises limitées. Lève l'exception finale si échec."""
    tentative_503 = 0
    quota_minute_retente = False

    while True:
        try:
            return client.models.generate_content(model=MODEL, contents=prompt)
        except Exception as exc:
            type_err, _ = _classer_erreur(exc)

            if type_err == "TEMPORAIRE" and tentative_503 < len(ATTENTES_503):
                attente = ATTENTES_503[tentative_503]
                tentative_503 += 1
                print(f"   ⏳ Gemini temporairement indisponible, nouvelle tentative dans {attente} s")
                time.sleep(attente)
                continue

            if type_err == "QUOTA_MINUTE" and not quota_minute_retente:
                quota_minute_retente = True
                print(f"   ⏳ Quota par minute atteint, attente de {ATTENTE_QUOTA_MINUTE} s")
                time.sleep(ATTENTE_QUOTA_MINUTE)
                continue

            raise

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
            "erreur_ia_type": "CLE_ABSENTE",
        }

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
            "url_reelle": resultat.get("url_reelle", ""),
            "extrait_page": str(resultat.get("extrait_page", ""))[:2000],
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
ANNÉES : 2026 et 2027. L'année 2027 est PRIORITAIRE.

RÈGLE D'ANNÉE
- Édition 2027 (annoncée ou prévisionnelle) : forte valeur pour la veille.
  Si l'intérêt Apec est plausible, choisis PERTINENT ou A_VERIFIER, jamais
  NON_PERTINENT par simple manque de détails (dates ou programme pas encore publiés).
- Édition 2026 : à traiter normalement.
- Édition 2025 ou antérieure : NON_PERTINENT (hors période), sauf si le texte
  annonce clairement une prochaine édition en 2026 ou 2027.

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
- hors zone lorsque cela est démontré ;
- édition 2025 ou antérieure sans prochaine édition 2026/2027 annoncée.

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

PAGE ENRICHIE (contenu réel de l'article ou de la page de l'événement)
URL réelle : {champs["url_reelle"]}
Titre : {champs["titre_page"]}
Description : {champs["description_page"]}
Date événement : {champs["date_page"]}
Lieu : {champs["lieu_page"]}
Organisateur : {champs["organisateur_page"]}
Extrait du texte : {champs["extrait_page"]}

UTILISATION DE LA PAGE
- Appuie ta décision sur le texte de la page quand il est présent : public,
  niveau d'études, entreprises, date, lieu, organisateur.
- Si le texte de la page contredit le titre RSS, fais confiance à la page.
- Si la page est vide ou absente, ne conclus pas à NON_PERTINENT pour ce seul
  motif : choisis A_VERIFIER si le potentiel Apec reste plausible.

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

        response = _appeler_gemini(client, prompt)

        texte = (response.text or "").strip()
        if texte.startswith("```"):
            texte = texte.replace("```json", "").replace("```", "").strip()

        data = json.loads(texte)

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
            "erreur_ia_type": "",
        }

    except Exception as e:
        type_err, msg = _classer_erreur(e)
        return {
            "decision_ia": "IA_INDISPONIBLE",
            "confiance_ia": 0,
            "raison_ia": f"Erreur IA ({type_err}) : {msg[:300]}",
            "categorie_ia": "",
            "date_ia": "",
            "departement_ia": "",
            "erreur_ia_type": type_err,
        }
