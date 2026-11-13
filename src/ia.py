import os
import json
from google import genai


MODEL = "gemini-2.5-flash"


def analyser_resultat(resultat):
    """
    Analyse un résultat RSS avec Gemini.
    Retourne un dictionnaire exploitable par main.py.
    """

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
        source = resultat.get("source", "")
        url = resultat.get("url", "")
        date = resultat.get("date", "")
        departement = resultat.get("departement", "")

        prompt = f"""
Tu es un analyste spécialisé dans la veille événementielle pour l'Apec.

Tu dois déterminer si l'événement suivant est pertinent pour une veille
portant sur :

- salons emploi
- forums emploi
- recrutement
- job dating
- formation professionnelle
- événements liés aux cadres
- enseignement supérieur
- universités
- écoles d'ingénieurs
- écoles de commerce
- événements alumni
- portes ouvertes de l'enseignement supérieur

Zone géographique cible :
86, 79, 87, 17, 16, 23, 19

Années recherchées :
2025, 2026, 2027

Règles importantes :
- Ne jamais inventer une information.
- Un événement doit réellement être identifiable.
- Un événement uniquement collège/lycée n'est pas pertinent.
- Les événements CCI sont à exclure.
- Un événement mixte peut être pertinent s'il comporte une vraie composante
  enseignement supérieur, emploi, recrutement ou formation.
- Si les informations sont insuffisantes pour décider, utiliser A_VERIFIER.
- Si l'événement est clairement hors périmètre, utiliser NON_PERTINENT.
- Si l'événement correspond clairement au périmètre Apec, utiliser PERTINENT.

DONNÉES :

Titre :
{titre}

Description :
{description}

Source :
{source}

Date :
{date}

Département détecté :
{departement}

URL :
{url}

Réponds UNIQUEMENT avec un JSON valide ayant exactement cette structure :

{{
  "decision": "PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance": 0,
  "raison": "explication courte",
  "categorie": "emploi|recrutement|formation|enseignement_superieur|alumni|portes_ouvertes|autre",
  "date": "",
  "departement": ""
}}

Le champ confiance doit être un entier entre 0 et 100.
Ne complète pas une information inconnue.
"""

        response = client.models.generate_content(
            model=MODEL,
            contents=prompt,
        )

        texte = response.text.strip()

        # Nettoyage éventuel des balises Markdown
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
