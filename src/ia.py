import os
import json
from google import genai

MODEL = "gemini-3.1-flash-lite"


def analyser_resultat(resultat):
    """Analyse un résultat RSS selon la cible Apec cadres / jeunes diplômés."""

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {
            "decision_ia": "IA_INDISPONIBLE",
            "confiance_ia": 0,
            "raison_ia": "Clé GEMINI_API_KEY absente.",
            "categorie_ia": "",
            "date_ia": "",
            "departement_ia": "",
        }

    try:
        client = genai.Client(api_key=api_key)

        titre = resultat.get("titre", "")
        description = resultat.get("description", "")
        titre_page = resultat.get("titre_page", "")
        description_page = resultat.get("description_page", "")
        date_evenement_page = resultat.get("date_evenement_page", "")
        lieu_page = resultat.get("lieu_page", "")
        organisateur_page = resultat.get("organisateur_page", "")
        source = resultat.get("source", "")
        url = resultat.get("url", "")
        date = resultat.get("date_publication", "")
        annee_recherchee = resultat.get("annee_recherchee", "")
        annee_detectee = resultat.get("annee_detectee", "")
        departement = resultat.get("departement", "")
        evenement_detecte = resultat.get("evenement_detecte", "")
        geographie_verifiee = resultat.get("geographie_verifiee", "")
        score_apec = resultat.get("score_cible_apec", resultat.get("score", 0))
        signaux_cible = resultat.get("signaux_cible_apec", "")
        signaux_superieur = resultat.get("signaux_superieur", "")
        signaux_entreprise = resultat.get("signaux_entreprise", "")
        signaux_exclusion = resultat.get("signaux_exclusion", "")

        prompt = f"""
Tu es l'analyste final d'une veille événementielle Apec.

OBJECTIF UNIQUE : ne conserver que les événements réellement utiles à l'Apec,
principalement pour :
- les cadres en poste ou en recherche ;
- les managers et profils experts ;
- les cadres dirigeants ;
- les jeunes diplômés Bac+3/Bac+5 et plus ;
- les étudiants de l'enseignement supérieur proches du diplôme ;
- les masters, écoles d'ingénieurs, écoles de commerce, grandes écoles,
  universités, IAE, doctorants, docteurs et alumni ;
- les relations entreprises / recrutement lorsque le public visé est qualifié.

ZONE : 86, 79, 87, 17, 16, 23, 19.
ANNÉES : 2025, 2026, 2027.

RÈGLES DE TRI :
1. Un événement doit être réellement identifiable.
2. La simple présence des mots « emploi », « formation », « étudiant » ou
   « recrutement » ne suffit pas.
3. Un forum universitaire peut être PERTINENT même sans écrire « cadre » ou
   « jeune diplômé » si le contexte montre un public supérieur qualifié et une
   vraie relation avec les entreprises / recruteurs.
4. Un événement réunissant université + employeurs + étudiants/diplômés est
   particulièrement intéressant.
5. Les événements explicitement destinés aux cadres ou jeunes diplômés sont
   prioritaires.
6. Les rencontres Apec, Apec Connect, cadres, diplômés, alumni, recrutement
   qualifié, écoles d'ingénieurs / commerce et relations entreprises sont des
   signaux forts.
7. Les événements de reconversion, mobilité ou évolution professionnelle sont
   pertinents lorsqu'ils concernent des actifs qualifiés ou des cadres.
8. Les événements uniquement collège, lycée, post-3e, orientation générale
   pré-Bac+3, jobs étudiants, jobs saisonniers, petits boulots, recrutement
   sans qualification, intérim générique ou métiers très éloignés de la cible
   Apec sont NON_PERTINENTS, sauf preuve claire d'un public cadre / diplômé.
   En revanche, un job dating ou forum de recrutement ne doit PAS être rejeté
   simplement parce que le mot « cadre » ou le niveau de diplôme n'apparaît pas :
   si la page montre des employeurs/recruteurs + université/enseignement
   supérieur/école/masters/ingénieurs/diplômés, il faut au minimum A_VERIFIER.
9. En cas de public potentiellement qualifié mais de preuve incomplète, choisir
   A_VERIFIER plutôt que NON_PERTINENT. NON_PERTINENT est réservé aux cas où
   l'absence de pertinence Apec est suffisamment démontrée.
10. Ne jamais inventer une date, un lieu, un public ou un niveau de diplôme.
11. Un article d'actualité qui parle d'un événement sans permettre d'identifier
    réellement l'événement ne doit pas être retenu comme PERTINENT.
12. Les CCI sont exclues comme événement selon les règles du projet.
13. En cas de doute réel, utiliser A_VERIFIER plutôt que forcer PERTINENT.

SIGNALS PYTHON DÉJÀ CALCULÉS :
- score Apec : {score_apec}
- cible directe : {signaux_cible}
- enseignement supérieur : {signaux_superieur}
- relation entreprises : {signaux_entreprise}
- exclusions détectées : {signaux_exclusion}

DONNÉES RSS :
Titre RSS : {titre}
Description RSS : {description}
Titre de la page : {titre_page}
Description de la page : {description_page}
Date de l'événement sur la page : {date_evenement_page}
Lieu détecté sur la page : {lieu_page}
Organisateur détecté sur la page : {organisateur_page}
Source : {source}
Date publication : {date}
Année recherchée : {annee_recherchee}
Année détectée : {annee_detectee}
Département : {departement}
Événement détecté : {evenement_detecte}
Géographie vérifiée : {geographie_verifiee}
URL : {url}

DÉCISION :
- PERTINENT = intérêt Apec clairement démontré.
- A_VERIFIER = potentiel réel mais preuve insuffisante ou public/date à confirmer.
- NON_PERTINENT = hors cible Apec ou signal insuffisant.

Réponds UNIQUEMENT avec un JSON valide :
{{
  "decision": "PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance": 0,
  "raison": "explication courte et factuelle",
  "categorie": "emploi|recrutement|formation|enseignement_superieur|alumni|portes_ouvertes|reconversion|entreprises_rh|autre",
  "date": "",
  "departement": ""
}}

Le champ confiance est un entier entre 0 et 100.
Le champ date contient uniquement une date d'événement explicitement identifiable.
Ne mets aucun texte avant ou après le JSON.
"""

        response = client.models.generate_content(
            model=MODEL,
            contents=prompt,
        )

        texte = response.text.strip()
        if texte.startswith("```"):
            texte = texte.replace("```json", "")
            texte = texte.replace("```", "")
            texte = texte.strip()

        data = json.loads(texte)
        return {
            "decision_ia": data.get("decision", "A_VERIFIER"),
            "confiance_ia": int(data.get("confiance", 0)),
            "raison_ia": data.get("raison", ""),
            "categorie_ia": data.get("categorie", ""),
            "date_ia": data.get("date", ""),
            "departement_ia": data.get("departement", ""),
        }

    except Exception as e:
        return {
            "decision_ia": "IA_INDISPONIBLE",
            "confiance_ia": 0,
            "raison_ia": f"Erreur IA : {str(e)[:300]}",
            "categorie_ia": "",
            "date_ia": "",
            "departement_ia": "",
        }
