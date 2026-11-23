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

        # Champs issus de l'enrichissement de la page lorsqu'ils existent.
        titre_page = resultat.get("titre_page", "")
        description_page = resultat.get("description_page", "")
        date_evenement_page = resultat.get("date_evenement_page", "")
        lieu_page = resultat.get("lieu_page", "")
        organisateur_page = resultat.get("organisateur_page", "")

        prompt = f"""
Tu es l'analyste final d'une veille événementielle Apec.

OBJECTIF
Identifier les événements réellement utiles à l'Apec, avec une priorité sur :
- les cadres en poste ou en recherche ;
- les managers, experts et cadres dirigeants ;
- les jeunes diplômés Bac+3/Bac+5 et plus ;
- les étudiants de l'enseignement supérieur proches du diplôme ;
- les masters, écoles d'ingénieurs, écoles de commerce, grandes écoles,
  universités, IAE, doctorants, docteurs et alumni ;
- les relations entreprises/recruteurs lorsqu'elles concernent un public
  qualifié ou de l'enseignement supérieur.

ZONE : départements 86, 79, 87, 17, 16, 23, 19.
ANNÉES RECHERCHÉES : 2025, 2026, 2027.

IMPORTANT : NE PAS EXIGER QUE LES MOTS « CADRE » OU « JEUNE DIPLÔMÉ »
APPARAISSENT LITTÉRALEMENT.

Un événement peut être PERTINENT lorsque le contexte démontre indirectement
qu'il touche une population correspondant à la cible Apec.

CRITÈRES PERTINENT
Un événement peut être PERTINENT si au moins une des situations suivantes
est clairement démontrée :

1. Il est explicitement destiné aux cadres, jeunes diplômés, diplômés,
   managers, experts ou profils qualifiés.

2. Il est organisé avec ou par une université, une école d'ingénieurs,
   une école de commerce, une grande école, un IAE ou une structure
   d'enseignement supérieur ET comporte une vraie relation avec des
   entreprises, employeurs, recruteurs, recrutement, carrières ou insertion
   professionnelle.

3. Il s'agit d'un forum entreprises / rencontre entreprises / job dating /
   recrutement dans un établissement d'enseignement supérieur, avec un
   public étudiant, diplômé, master, ingénieur ou autre public qualifié.

4. Il concerne des masters, Bac+5, doctorants, docteurs, ingénieurs,
   alumni ou jeunes talents et comporte une dimension entreprises,
   recrutement, carrière ou insertion.

5. Il s'agit d'un événement de mobilité, évolution professionnelle ou
   reconversion et le public visé est composé d'actifs qualifiés, cadres
   ou profils professionnels correspondant à l'Apec.

6. Il est explicitement lié à Apec, Apec Connect ou à une action Apec.

7. Un forum général de l'emploi peut être PERTINENT si les informations
   disponibles montrent un public qualifié, des employeurs ciblant des
   diplômés/cadres, ou une forte composante enseignement supérieur.

EXEMPLE IMPORTANT
« Forum entreprises – Université de Poitiers » peut être PERTINENT même si
la page ne contient pas le mot « cadre », lorsque la relation université +
entreprises + étudiants/diplômés est clairement établie.

« Rencontre entreprises / étudiants – Université X » peut également être
PERTINENT si le contexte indique un public d'enseignement supérieur.

À NE PAS FAIRE
- Ne pas rejeter un événement uniquement parce que « cadre » n'est pas écrit.
- Ne pas rejeter un événement uniquement parce que le public exact n'est
  pas détaillé si un contexte enseignement supérieur + entreprises est clair.
- Ne pas considérer automatiquement tout événement universitaire comme
  pertinent : il faut une dimension professionnelle, entreprise, recrutement,
  carrière ou insertion, sauf événement explicitement destiné aux diplômés,
  alumni ou cadres.
- Ne pas considérer automatiquement tout « forum emploi » comme pertinent.

NON_PERTINENT
Rejeter clairement :
- jobs étudiants ;
- jobs d'été ;
- jobs saisonniers ;
- petits boulots ;
- recrutement principalement restauration ;
- caisse, manutention, préparation de commandes ou emplois similaires ;
- intérim généraliste sans public qualifié ;
- emplois sans qualification ;
- événements uniquement collège ;
- événements uniquement lycée ;
- orientation post-3e ;
- orientation générale pré-Bac+3 sans dimension Apec ;
- événements commerciaux ou grand public sans lien professionnel pertinent ;
- événement qui n'est pas réellement identifiable ;
- simple article d'actualité mentionnant un événement sans permettre
  d'identifier l'événement ;
- événements hors zone géographique lorsque l'information le démontre ;
- CCI comme événement, conformément aux règles du projet.

A_VERIFIER
Utiliser A_VERIFIER lorsqu'un intérêt Apec réel est possible mais qu'une
information importante manque pour conclure.

Exemples :
- forum emploi avec entreprises clairement présentes mais public insuffisamment
  documenté ;
- événement universitaire avec employeurs mais niveau de diplôme non précisé ;
- rencontre professionnelle pouvant concerner des cadres mais public ambigu ;
- date, lieu ou public à confirmer.

IMPORTANT SUR LES FAUX NÉGATIFS
En cas de combinaison forte :
enseignement supérieur + entreprises/recruteurs + étudiants/diplômés,
préférer PERTINENT ou A_VERIFIER plutôt que NON_PERTINENT.

En cas de combinaison :
cadres/diplômés/alumni/école supérieure + recrutement/carrière/entreprises,
préférer PERTINENT.

Ne jamais inventer une information absente de la source.

DONNÉES DE SCORING PYTHON
Score Apec : {score_apec}
Cible directe : {signaux_cible}
Enseignement supérieur : {signaux_superieur}
Relation entreprises : {signaux_entreprise}
Exclusions détectées : {signaux_exclusion}

DONNÉES RSS
Titre : {titre}
Description : {description}
Source : {source}
Date publication : {date}
Année recherchée : {annee_recherchee}
Année détectée : {annee_detectee}
Département : {departement}
Événement détecté : {evenement_detecte}
Géographie vérifiée : {geographie_verifiee}
URL : {url}

DONNÉES PAGE ENRICHIE
Titre page : {titre_page}
Description page : {description_page}
Date événement page : {date_evenement_page}
Lieu page : {lieu_page}
Organisateur page : {organisateur_page}

DÉCISION
- PERTINENT = intérêt Apec suffisamment démontré.
- A_VERIFIER = potentiel Apec réel mais preuve insuffisante ou information
  importante à confirmer.
- NON_PERTINENT = hors cible ou exclusion claire.

La confiance est un entier de 0 à 100.
Le champ date contient uniquement une date d'événement explicitement
identifiable. Ne jamais inventer une date.
Le département doit être celui explicitement identifiable dans les données.
Si absent, laisser vide.

Réponds UNIQUEMENT avec un JSON valide :
{{
  "decision": "PERTINENT|A_VERIFIER|NON_PERTINENT",
  "confiance": 0,
  "raison": "explication courte et factuelle",
  "categorie": "emploi|recrutement|formation|enseignement_superieur|alumni|portes_ouvertes|reconversion|entreprises_rh|autre",
  "date": "",
  "departement": ""
}}

Aucun texte avant ou après le JSON.
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
