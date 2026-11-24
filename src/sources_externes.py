"""
Collecte multi-sources pour la veille Apec.

Complète Google News avec :
  1. des SITES OFFICIELS (universités, écoles, villes, parcs d'expositions) :
     exploration de la page d'accueil/agenda, puis lecture des pages
     d'événements trouvées ;
  2. OPENAGENDA, via le jeu de données public OpenDataSoft (sans clé) ;
  3. BING NEWS (flux RSS, liens déjà décodés) ;
  4. des FLUX RSS/ATOM directs listés dans config/sources_collecte.yml.

Principes :
  - une source en erreur n'arrête jamais les autres ;
  - chaque résultat reprend exactement le format des résultats Google News,
    avec deux champs en plus : `origine` et `niveau_source` ;
  - les sources officielles sont placées en tête (le dédoublonnage conserve
    la première occurrence : la source primaire l'emporte sur la presse).
"""

import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from email.utils import format_datetime
from urllib.parse import parse_qs, quote_plus, urljoin, urlparse

import feedparser
from bs4 import BeautifulSoup

# ------------------------------------------------------------
# Vocabulaire
# ------------------------------------------------------------

MOTS_EVENEMENT = (
    "forum", "salon", "job dating", "jobdating", "portes ouvertes",
    "porte ouverte", "open day", "rencontre", "recrutement", "alumni",
    "carrière", "carriere", "career", "emploi", "entreprises", "job fair",
    "speed", "journée des", "journee des", "conférence", "conference",
    "table ronde", "insertion", "networking", "semaine de l",
)

MOTS_AGENDA = (
    "agenda", "evenement", "événement", "actualit", "calendrier",
    "manifestation", "salon", "forum", "rendez-vous", "sortir",
)

# Mots-clés d'intérêt pour filtrer les événements OpenAgenda (très généraliste).
MOTS_INTERET_OPENAGENDA = (
    "emploi", "recrutement", "job", "forum", "salon", "alumni", "étudiant",
    "etudiant", "université", "universite", "école", "ecole", "formation",
    "carrière", "carriere", "entreprise", "portes ouvertes", "orientation",
    "insertion", "diplôm", "diplom", "master", "ingénieur", "ingenieur",
    "career", "networking", "reconversion", "cadre",
)

REQUETES_BING = (
    "forum emploi", "forum entreprises", "salon étudiant", "job dating",
    "portes ouvertes université", "career day", "rencontres entreprises",
)

EXTENSIONS_IGNOREES = (
    ".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".zip", ".doc", ".docx",
    ".xls", ".xlsx", ".ppt", ".pptx", ".mp4", ".mp3", ".css", ".js", ".ico",
    ".pdf",
)

ODS_URL = (
    "https://public.opendatasoft.com/api/explore/v2.1/catalog/datasets/"
    "evenements-publics-openagenda/records"
)

# ------------------------------------------------------------
# Utilitaires
# ------------------------------------------------------------


def _nettoyer(texte):
    return " ".join(str(texte or "").split())


def _bas(texte):
    return _nettoyer(texte).lower()


def _vers_rfc(valeur):
    """Date ISO (ou assimilée) -> format RFC 822, '' si illisible."""
    texte = str(valeur or "").strip()
    if not texte:
        return ""
    try:
        dt = datetime.fromisoformat(texte.replace("Z", "+00:00"))
        return format_datetime(dt.replace(tzinfo=None))
    except Exception:
        return ""


def _domaine_racine(url):
    hote = urlparse(url).netloc.lower().split(":")[0]
    morceaux = hote.split(".")
    return ".".join(morceaux[-2:]) if len(morceaux) >= 2 else hote


def _meme_site(url_a, url_b):
    return _domaine_racine(url_a) == _domaine_racine(url_b)


def _lien_utilisable(href):
    if not href:
        return False
    h = href.strip().lower()
    if h.startswith(("#", "javascript:", "mailto:", "tel:")):
        return False
    chemin = urlparse(h).path
    return not chemin.endswith(EXTENSIONS_IGNOREES)


def _meta(soup, *noms):
    for nom in noms:
        balise = soup.find("meta", attrs={"property": nom}) or soup.find(
            "meta", attrs={"name": nom}
        )
        if balise and balise.get("content"):
            return _nettoyer(balise["content"])
    return ""


def _jsonld(soup):
    """Objets JSON-LD de la page (liste, jamais d'exception)."""
    objets = []
    for balise in soup.find_all("script", attrs={"type": "application/ld+json"}):
        try:
            data = json.loads(balise.string or balise.get_text() or "")
        except Exception:
            continue
        pile = data if isinstance(data, list) else [data]
        while pile:
            o = pile.pop()
            if isinstance(o, dict):
                objets.append(o)
                for v in o.values():
                    if isinstance(v, (list, dict)):
                        pile.extend(v if isinstance(v, list) else [v])
    return objets


def _texte_article(soup, limite=700):
    copie = BeautifulSoup(str(soup), "lxml")
    for b in copie(["script", "style", "nav", "footer", "header", "aside",
                    "form", "noscript", "iframe"]):
        b.decompose()
    zone = copie.find("article") or copie.find("main") or copie.body or copie
    morceaux = [
        p.get_text(" ", strip=True)
        for p in zone.find_all(["h1", "h2", "h3", "p", "li"])
    ]
    return _nettoyer(" ".join(m for m in morceaux if len(m) > 25))[:limite]


def _telecharger(session, url, timeout):
    reponse = session.get(url, timeout=timeout, allow_redirects=True)
    reponse.raise_for_status()
    type_contenu = reponse.headers.get("Content-Type", "").lower()
    if type_contenu and "html" not in type_contenu and "xml" not in type_contenu \
            and "json" not in type_contenu:
        raise ValueError(f"contenu non textuel ({type_contenu[:30]})")
    return reponse


def _resultat(titre, url, source, description, date_publication, departement,
              nom_departement, origine, niveau, categorie_numero=2,
              geographie_source=False):
    return {
        "titre": _nettoyer(titre),
        "url": url,
        "source": source,
        "description": _nettoyer(description),
        "date_publication": date_publication,
        "departement": departement,
        "nom_departement": nom_departement,
        "annee_recherchee": "",
        "categorie_numero": categorie_numero,
        "categorie": (
            "Salon étudiant / enseignement supérieur"
            if categorie_numero == 1
            else "Emploi / recrutement / formation"
        ),
        "mot_cle": f"source:{origine}",
        "origine": origine,
        "niveau_source": niveau,
        "geographie_source": "OUI" if geographie_source else "NON",
    }


def _mot_present(mot, texte):
    return re.search(rf"(?<![a-zà-ÿ]){re.escape(mot)}(?![a-zà-ÿ])", texte) is not None


def _departement_depuis_texte(texte, zone):
    """Code département du premier nom/ville de la zone trouvé en mot entier."""
    t = _bas(texte)
    for mot, code in zone.items():
        if _mot_present(mot, t):
            return code
    return ""


def _categorie(texte):
    t = _bas(texte)
    superieur = ("universit", "étudiant", "etudiant", "école", "ecole", "master",
                 "alumni", "portes ouvertes", "faculté", "faculte", "iae")
    return 1 if any(m in t for m in superieur) else 2


# ------------------------------------------------------------
# 1. Sites officiels (pages HTML)
# ------------------------------------------------------------


def _liens_page(html, url_page):
    """[(texte, url_absolue)] des liens d'une page."""
    soup = BeautifulSoup(html, "lxml")
    liens, vus = [], set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if not _lien_utilisable(href):
            continue
        absolu = urljoin(url_page, href).split("#")[0]
        if absolu in vus:
            continue
        vus.add(absolu)
        liens.append((_nettoyer(a.get_text(" ", strip=True)), absolu))
    return liens


def analyser_page_evenement(session, url, timeout):
    """Lit une page d'événement : titre, description, dates."""
    reponse = _telecharger(session, url, timeout)
    soup = BeautifulSoup(reponse.content, "lxml")

    titre = (
        _meta(soup, "og:title")
        or (soup.title.get_text(" ", strip=True) if soup.title else "")
        or (soup.h1.get_text(" ", strip=True) if soup.h1 else "")
    )
    description = _meta(soup, "og:description", "description")
    corps = _texte_article(soup)

    date_pub = _meta(soup, "article:published_time", "og:updated_time",
                     "datePublished", "date")
    date_evt = ""
    for o in _jsonld(soup):
        date_evt = date_evt or str(o.get("startDate") or "")
        date_pub = date_pub or str(o.get("datePublished") or "")
        if not description and o.get("description"):
            description = _nettoyer(o["description"])

    morceaux = [description, corps]
    if date_evt:
        morceaux.append(f"Date de l'événement : {date_evt[:10]}.")

    return {
        "titre": _nettoyer(titre),
        "description": _nettoyer(" ".join(m for m in morceaux if m))[:1500],
        "date_publication": _vers_rfc(date_pub),
        "url_finale": reponse.url,
    }


def explorer_source_html(source, session_factory, timeout, zone,
                         max_agenda=5, max_evenements=15, delai=0.4):
    """Explore un site : accueil -> pages agenda -> pages d'événements."""
    nom = source["nom"]
    depart = str(source.get("departement", "") or "")
    nom_dep = source.get("nom_departement", "")
    session = session_factory()
    origine = source.get("origine", "site_officiel")
    niveau = int(source.get("niveau", 2))

    pages_depart = [source["url"]] + list(source.get("pages", []) or [])
    candidats_evt, pages_agenda, vus = [], [], set()

    for url_depart in pages_depart:
        try:
            html = _telecharger(session, url_depart, timeout).content
        except Exception as e:
            print(f"   ⚠️ {nom} : page inaccessible ({type(e).__name__})")
            continue

        for texte, lien in _liens_page(html, url_depart):
            if not _meme_site(lien, source["url"]) or lien in vus:
                continue
            vus.add(lien)
            repere = _bas(texte + " " + urlparse(lien).path)

            if len(texte) >= 12 and any(m in _bas(texte) for m in MOTS_EVENEMENT):
                candidats_evt.append((texte, lien))
            elif any(m in repere for m in MOTS_AGENDA) and len(pages_agenda) < max_agenda:
                pages_agenda.append(lien)

    for page in pages_agenda:
        time.sleep(delai)
        try:
            html = _telecharger(session, page, timeout).content
        except Exception:
            continue
        for texte, lien in _liens_page(html, page):
            if not _meme_site(lien, source["url"]) or lien in vus:
                continue
            vus.add(lien)
            if len(texte) >= 12 and any(m in _bas(texte) for m in MOTS_EVENEMENT):
                candidats_evt.append((texte, lien))

    resultats, lues = [], 0
    for texte, lien in candidats_evt[:max_evenements]:
        time.sleep(delai)
        try:
            page = analyser_page_evenement(session, lien, timeout)
        except Exception:
            continue
        lues += 1

        titre = page["titre"] or texte
        contenu = f"{titre} {page['description']}"

        # Source nationale (sans département) : on exige un indice de zone,
        # et on en déduit le département de l'événement.
        dep_item, nom_dep_item = depart, nom_dep
        if not depart:
            dep_item = _departement_depuis_texte(contenu, zone)
            if not dep_item:
                continue
            nom_dep_item = source.get("_noms_dep", {}).get(dep_item, "")

        resultats.append(_resultat(
            titre, page["url_finale"] or lien, nom, page["description"],
            page["date_publication"], dep_item, nom_dep_item, origine, niveau,
            categorie_numero=_categorie(contenu),
            geographie_source=True,
        ))

    print(f"   {nom} : {len(candidats_evt)} lien(s) repéré(s), "
          f"{lues} page(s) lue(s), {len(resultats)} retenu(s)")
    return resultats


# ------------------------------------------------------------
# 2. OpenAgenda (OpenDataSoft)
# ------------------------------------------------------------


def _champ(fiche, *noms):
    for nom in noms:
        valeur = fiche.get(nom)
        if isinstance(valeur, list):
            valeur = ", ".join(str(v) for v in valeur)
        if valeur:
            return _nettoyer(valeur)
    return ""


def collecter_openagenda(departements, session_factory, timeout,
                         max_par_departement=300, aujourdhui=None):
    session = session_factory()
    aujourdhui = aujourdhui or datetime.now().strftime("%Y-%m-%d")
    resultats = []

    for code, nom_dep in departements.items():
        total_dep, retenus = 0, 0
        for avec_date in (True, False):
            clause = f'location_department="{nom_dep}"'
            if avec_date:
                clause += f" and lastdate_end >= date'{aujourdhui}'"
            fiches, erreur = [], ""
            try:
                for offset in range(0, max_par_departement, 100):
                    rep = session.get(ODS_URL, params={
                        "where": clause, "limit": 100, "offset": offset,
                        "order_by": "firstdate_begin",
                    }, timeout=timeout)
                    rep.raise_for_status()
                    data = rep.json()
                    lot = data.get("results") or [
                        r.get("fields", r) for r in data.get("records", [])
                    ]
                    fiches.extend(lot)
                    if len(lot) < 100:
                        break
            except Exception as e:
                erreur = f"{type(e).__name__}: {str(e)[:80]}"
            if fiches or not erreur:
                break
        if erreur and not fiches:
            print(f"   ⚠️ OpenAgenda {nom_dep} : {erreur}")
            continue

        for fiche in fiches:
            total_dep += 1
            titre = _champ(fiche, "title_fr", "title", "title_en")
            descr = _champ(fiche, "description_fr", "longdescription_fr",
                           "description", "description_en")
            mots = _champ(fiche, "keywords_fr", "keywords")
            url = _champ(fiche, "canonicalurl", "url")
            if not titre or not url:
                continue

            contenu = f"{titre} {descr} {mots}"
            if not any(m in _bas(contenu) for m in MOTS_INTERET_OPENAGENDA):
                continue

            debut = _champ(fiche, "firstdate_begin", "firstdate")
            fin = _champ(fiche, "lastdate_end", "lastdate")
            lieu = ", ".join(x for x in (
                _champ(fiche, "location_name"),
                _champ(fiche, "location_city"),
            ) if x)
            details = f"Date : {debut[:10]} au {fin[:10]}. Lieu : {lieu}. {mots}"

            resultats.append(_resultat(
                titre, url, "OpenAgenda", f"{descr[:800]} {details}",
                "", code, nom_dep, "openagenda", 3,
                categorie_numero=_categorie(contenu),
                geographie_source=True,
            ))
            retenus += 1

        print(f"   OpenAgenda {nom_dep} : {total_dep} événement(s), {retenus} retenu(s)")

    return resultats


# ------------------------------------------------------------
# 3. Bing News (RSS)
# ------------------------------------------------------------


def _url_reelle_bing(lien):
    """Les liens Bing News embarquent l'URL de l'article dans ?url=."""
    try:
        valeur = parse_qs(urlparse(lien).query).get("url", [""])[0]
        return valeur or lien
    except Exception:
        return lien


def _flux_vers_resultats(contenu, source_defaut, depart, nom_dep, origine,
                         niveau, decoder_bing=False):
    flux = feedparser.parse(contenu)
    sortie = []
    for article in flux.entries:
        titre = article.get("title", "").strip()
        lien = article.get("link", "").strip()
        if not titre or not lien:
            continue
        if decoder_bing:
            lien = _url_reelle_bing(lien)
        source = source_defaut
        if hasattr(article, "source"):
            source = article.source.get("title", "") or source
        source = article.get("news_source") or source

        description = article.get("summary", "") or article.get("description", "")
        description = BeautifulSoup(description, "lxml").get_text(" ", strip=True)

        sortie.append(_resultat(
            titre, lien, source, description,
            article.get("published", "").strip(),
            depart, nom_dep, origine, niveau,
            categorie_numero=_categorie(f"{titre} {description}"),
        ))
    return sortie


def collecter_bing(departements, annees, session_factory, timeout):
    session = session_factory()
    resultats = []
    for code, nom_dep in departements.items():
        total = 0
        for annee in annees:
            for mot in REQUETES_BING:
                texte = f'"{mot}" "{nom_dep}" {annee}'
                url = (
                    "https://www.bing.com/news/search?"
                    f"q={quote_plus(texte)}&format=rss&setlang=fr&cc=FR"
                )
                try:
                    rep = session.get(url, timeout=timeout)
                    rep.raise_for_status()
                    lot = _flux_vers_resultats(
                        rep.content, "Bing News", code, nom_dep, "bing_news", 4,
                        decoder_bing=True,
                    )
                    for r in lot:
                        r["annee_recherchee"] = annee
                    resultats.extend(lot)
                    total += len(lot)
                except Exception:
                    continue
                time.sleep(0.3)
        print(f"   Bing News {nom_dep} : {total} résultat(s)")
    return resultats


# ------------------------------------------------------------
# 4. Flux RSS/Atom directs
# ------------------------------------------------------------


def collecter_rss_directs(flux_liste, session_factory, timeout):
    session = session_factory()
    resultats = []
    for flux in flux_liste or []:
        nom = flux.get("nom", flux.get("url", "flux"))
        try:
            rep = session.get(flux["url"], timeout=timeout)
            rep.raise_for_status()
            lot = _flux_vers_resultats(
                rep.content, nom, str(flux.get("departement", "") or ""),
                flux.get("nom_departement", ""), "rss_direct",
                int(flux.get("niveau", 3)),
            )
            if flux.get("departement"):
                for r in lot:
                    r["geographie_source"] = "OUI"
            resultats.extend(lot)
            print(f"   Flux {nom} : {len(lot)} élément(s)")
        except Exception as e:
            print(f"   ⚠️ Flux {nom} : {type(e).__name__}")
    return resultats


# ------------------------------------------------------------
# Orchestration
# ------------------------------------------------------------


def sources_html_depuis_config(sources_yml, extra):
    """Transforme sources.yml (+ sources_collecte.yml) en liste de sites."""
    liste = []
    for section in ("enseignement_superieur", "evenementiel",
                    "plateformes_evenements"):
        for s in (sources_yml or {}).get("sources", {}).get(section, []) or []:
            liste.append({
                "nom": s["nom"], "url": s["url"],
                "departement": s.get("departement", ""),
                "niveau": 2 if section != "plateformes_evenements" else 3,
                "origine": "site_officiel" if section != "plateformes_evenements"
                           else "plateforme",
            })
    for s in (extra or {}).get("pages_html", []) or []:
        liste.append(s)
    return liste


def collecter_sources_externes(sources_yml, extra, departements, annees,
                               session_factory, timeout=10, max_workers=6,
                               zone=None):
    """
    Retourne la liste des résultats de toutes les sources externes, sites
    officiels en tête. Jamais d'exception : au pire, une liste vide.
    """
    extra = extra or {}
    activer = extra.get("activer", {}) or {}
    zone = {str(k).lower(): v for k, v in (zone or {}).items()}
    noms_dep = dict(departements)

    print("\n========================================")
    print("SOURCES EXTERNES")
    print("========================================")

    officiels, openagenda, bing, rss = [], [], [], []

    if activer.get("sites_officiels", True):
        sites = sources_html_depuis_config(sources_yml, extra)
        for s in sites:
            s["nom_departement"] = noms_dep.get(str(s.get("departement", "")), "")
            s["_noms_dep"] = noms_dep
        print(f"Sites à explorer : {len(sites)}")
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = {
                ex.submit(explorer_source_html, s, session_factory, timeout,
                          zone): s for s in sites
            }
            for f in as_completed(futures):
                try:
                    officiels.extend(f.result())
                except Exception as e:
                    print(f"   ⚠️ {futures[f]['nom']} : {type(e).__name__}")

    if activer.get("openagenda", True):
        print("OpenAgenda…")
        try:
            openagenda = collecter_openagenda(noms_dep, session_factory, timeout)
        except Exception as e:
            print(f"   ⚠️ OpenAgenda : {type(e).__name__}")

    if activer.get("bing_news", True):
        print("Bing News…")
        try:
            bing = collecter_bing(noms_dep, annees, session_factory, timeout)
        except Exception as e:
            print(f"   ⚠️ Bing News : {type(e).__name__}")

    if extra.get("flux_rss"):
        print("Flux RSS directs…")
        try:
            rss = collecter_rss_directs(extra["flux_rss"], session_factory, timeout)
        except Exception as e:
            print(f"   ⚠️ Flux RSS : {type(e).__name__}")

    total = officiels + openagenda + rss + bing
    print(f"\nSources externes : {len(officiels)} site(s) officiel(s), "
          f"{len(openagenda)} OpenAgenda, {len(rss)} flux direct(s), "
          f"{len(bing)} Bing News = {len(total)} résultat(s)", flush=True)
    return total
