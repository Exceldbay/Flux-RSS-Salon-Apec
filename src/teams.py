import json
import os
from datetime import datetime, timedelta
from pathlib import Path

import requests


TEAMS_WEBHOOK_ENV = "TEAMS_PRIVATE_WEBHOOK_URL"
TEAMS_STATE_FILE = "teams_last_sent.txt"
TEAMS_MIN_INTERVAL_HOURS = 20


def _clean(value, limit=240):
    text = " ".join(str(value or "").split())
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _url(resultat):
    return resultat.get("url_finale") or resultat.get("url_reelle") or resultat.get("url") or ""


def _date(resultat):
    return resultat.get("date_evenement_page") or resultat.get("date_ia") or "date non précisée"


def diagnostic_teams(state_file: Path):
    """Retourne l'état du système Teams sans exposer le secret."""
    webhook = os.getenv(TEAMS_WEBHOOK_ENV, "").strip()
    return {
        "secret_present": bool(webhook),
        "force_send": os.getenv("TEAMS_FORCE_SEND", "false"),
        "etat_recent": _etat_recent(state_file),
        "state_file": str(state_file),
    }


def construire_resume_teams(delta, stats, meta, taches=None):
    """Construit une carte Adaptive Card compacte, orientée delta utile."""
    nouveaux_pertinents = delta.get("nouveaux_pertinents", [])
    promus = delta.get("promus_pertinents", [])
    nouveaux_verif = delta.get("nouveaux_a_verifier_2027", [])

    items = []
    vus = set()
    for groupe, label in (
        (nouveaux_pertinents, "NOUVEAU PERTINENT"),
        (promus, "PROMU"),
        (nouveaux_verif, "2027 À VÉRIFIER"),
    ):
        for r in groupe:
            cle = _url(r) or r.get("titre", "")
            if cle in vus:
                continue
            vus.add(cle)
            items.append((label, r))
            if len(items) >= 10:
                break
        if len(items) >= 10:
            break

    body = [
        {
            "type": "TextBlock",
            "text": "🔎 Veille APEC — synthèse quotidienne",
            "weight": "Bolder",
            "size": "Large",
            "wrap": True,
        },
        {
            "type": "TextBlock",
            "text": meta.get("date", datetime.now().strftime("%d/%m/%Y")),
            "isSubtle": True,
            "spacing": "None",
        },
        {
            "type": "FactSet",
            "facts": [
                {"title": "Nouveaux pertinents", "value": str(len(nouveaux_pertinents))},
                {"title": "Promus pertinents", "value": str(len(promus))},
                {"title": "Nouveaux 2027 à vérifier", "value": str(len(nouveaux_verif))},
                {"title": "Évolutions utiles", "value": str(delta.get("evolutions_utiles", 0))},
                {"title": "Pertinents actuels", "value": str(meta.get("pertinents", 0))},
                {"title": "À vérifier actuels", "value": str(meta.get("a_verifier", 0))},
                {"title": "Tâches terminées", "value": f"{meta.get('taches_terminees', 0)}/{meta.get('taches_total', 0)}"},
                {"title": "Erreurs tâches", "value": str(meta.get("taches_erreurs", 0))},
            ],
        },
    ]

    if items:
        body.append({
            "type": "TextBlock",
            "text": "À regarder en priorité",
            "weight": "Bolder",
            "spacing": "Medium",
        })
        for label, r in items:
            title = _clean(r.get("titre_page") or r.get("titre", ""), 180)
            dept = r.get("departement", "?")
            score = r.get("score", "")
            priorite = r.get("priorite_2027", "")
            date_evt = _date(r)
            source = _clean(r.get("source", ""), 80)
            detail = f"**{label}** · dép. {dept} · score {score}"
            if priorite and priorite != "FAIBLE":
                detail += f" · 2027 {priorite}"
            detail += f" · {date_evt}"
            if source:
                detail += f" · {source}"
            body.append({
                "type": "TextBlock",
                "text": title,
                "weight": "Bolder",
                "wrap": True,
                "spacing": "Small",
            })
            body.append({
                "type": "TextBlock",
                "text": detail,
                "wrap": True,
                "spacing": "None",
            })
            link = _url(r)
            if link:
                body.append({
                    "type": "ActionSet",
                    "actions": [{"type": "Action.OpenUrl", "title": "Ouvrir la source", "url": link}],
                    "spacing": "None",
                })
    else:
        body.append({
            "type": "TextBlock",
            "text": "Aucun nouveau signal suffisamment utile depuis la dernière veille.",
            "wrap": True,
            "spacing": "Medium",
        })

    # Suivi détaillé des 12 tâches : affiché uniquement si fourni.
    if taches:
        body.append({
            "type": "TextBlock",
            "text": "Suivi détaillé des tâches",
            "weight": "Bolder",
            "spacing": "Medium",
        })
        for t in taches:
            statut = str(t.get("statut", "EN ATTENTE"))
            icone = {
                "TERMINEE": "✅",
                "EN COURS": "🔄",
                "ERREUR": "❌",
                "NON EXECUTEE": "⏭️",
                "EN ATTENTE": "⏳",
            }.get(statut, "•")
            duree = t.get("duree_s", 0) or 0
            try:
                duree_txt = f" · {float(duree):.1f}s" if float(duree) > 0 else ""
            except Exception:
                duree_txt = ""
            detail = _clean(t.get("details", ""), 120)
            texte = f"{icone} {t.get('id', '')} — {t.get('tache', '')}{duree_txt}"
            if detail:
                texte += f" — {detail}"
            body.append({
                "type": "TextBlock",
                "text": texte,
                "wrap": True,
                "spacing": "None",
                "size": "Small",
            })

    top_ops = stats[:5]
    if top_ops:
        body.append({
            "type": "TextBlock",
            "text": "Opérateurs les plus utiles",
            "weight": "Bolder",
            "spacing": "Medium",
        })
        for op in top_ops:
            body.append({
                "type": "TextBlock",
                "text": (
                    f"{_clean(op.get('operateur', 'non défini'), 70)} — "
                    f"{op.get('uniques', 0)} uniques · "
                    f"{op.get('pertinents', 0)} pertinents · "
                    f"{op.get('a_verifier', 0)} à vérifier"
                ),
                "wrap": True,
                "spacing": "None",
            })

    erreurs_taches = int(meta.get("taches_erreurs", 0) or 0)
    statut_taches = "✅ Toutes les tâches sont terminées" if erreurs_taches == 0 else f"⚠️ {erreurs_taches} tâche(s) en erreur"
    body.append({
        "type": "TextBlock",
        "text": f"Suivi d'exécution : {statut_taches} · {meta.get('taches_terminees', 0)}/{meta.get('taches_total', 0)} terminées",
        "wrap": True,
        "spacing": "Medium",
        "weight": "Bolder",
    })

    body.append({
        "type": "TextBlock",
        "text": (
            f"IA : {meta.get('ia_analysees', 0)} appels · "
            f"{meta.get('ia_memoire', 0)} mémoire · "
            f"{meta.get('ia_non_analyse', 0)} non analysés · "
            f"quota bloqué : {meta.get('quota_bloque', 'NON')} · "
            f"modèle : {meta.get('gemini_model', 'non défini')} · "
            f"seuil : {meta.get('gemini_score_min', '—')} · "
            f"max : {meta.get('gemini_max_analyses', '—')}"
        ),
        "isSubtle": True,
        "wrap": True,
        "spacing": "Medium",
    })

    card = {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "contentUrl": None,
            "content": {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.2",
                "body": body,
            },
        }],
    }
    return card


def _etat_recent(path: Path):
    if not path.exists():
        return False
    try:
        dt = datetime.fromisoformat(path.read_text(encoding="utf-8").strip())
        return datetime.now() - dt < timedelta(hours=TEAMS_MIN_INTERVAL_HOURS)
    except Exception:
        return False


def envoyer_teams(card, state_file: Path):
    """Envoie la synthèse Teams via un webhook Workflows.

    Le workflow peut être déclenché manuellement avec TEAMS_FORCE_SEND=true
    pour tester l'envoi sans attendre l'expiration du délai de 20 h.
    """
    webhook = os.getenv(TEAMS_WEBHOOK_ENV, "").strip()
    force_send = os.getenv("TEAMS_FORCE_SEND", "false").strip().lower() in {
        "1", "true", "yes", "oui"
    }

    if not webhook:
        return {
            "envoye": False,
            "raison": "secret TEAMS_PRIVATE_WEBHOOK_URL absent",
        }

    recent = _etat_recent(state_file)
    if recent and not force_send:
        return {
            "envoye": False,
            "raison": "notification déjà envoyée dans les 20 dernières heures",
        }

    try:
        response = requests.post(
            webhook,
            json=card,
            headers={"Content-Type": "application/json"},
            timeout=20,
        )

        if response.status_code >= 400:
            detail = _clean(response.text, 500)
            if response.status_code in (404, 410):
                detail += (
                    " | Vérifier que l'URL provient bien d'un Workflow Teams "
                    "'When a Teams webhook request is received' et non d'un ancien "
                    "Office 365 Connector."
                )
            return {
                "envoye": False,
                "raison": f"Teams HTTP {response.status_code}: {detail}",
            }

        state_file.parent.mkdir(parents=True, exist_ok=True)
        state_file.write_text(
            datetime.now().isoformat(timespec="seconds"),
            encoding="utf-8",
        )
        return {
            "envoye": True,
            "raison": "message Teams envoyé via le webhook Workflows",
        }

    except Exception as e:
        return {
            "envoye": False,
            "raison": f"erreur Teams {type(e).__name__}: {e}",
        }


def sauvegarder_payload(card, output_file: Path):
    output_file.parent.mkdir(parents=True, exist_ok=True)
    output_file.write_text(json.dumps(card, ensure_ascii=False, indent=2), encoding="utf-8")
