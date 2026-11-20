import csv
import re
import unicodedata
from copy import copy
from datetime import datetime, date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"

INPUT_CSV = DATA / "resultats_ia.csv"
WORKBOOK = DATA / "Apec_Salons_Evenements.xlsx"

YEAR_SHEETS = {
    "2026": "Événements 2026",
    "2027": "Événements 2027",
}

DEPT_SHEETS = {
    "16": "16-Charente",
    "17": "17-Charente-Maritime",
    "19": "19-Corrèze",
    "23": "23-Creuse",
    "79": "79-Deux-Sèvres",
    "86": "86-Vienne",
    "87": "87-Haute-Vienne",
}

CHRONOLOGY_SHEET = "À venir (chronologie)"
NON_PERTINENT_SHEET = "Non pertinent"
NON_VERIFIABLE_SHEET = "Non vérifiable"

EVENT_HEADERS = {
    "N°",
    "Dép.",
    "Catégorie",
    "Événement",
    "Date début",
    "Date fin (ou 2e date distincte)",
    "Lieu",
    "Détails et récurrence",
    "Statut au 12/11/2026",
    "Prochaine édition connue",
    "Apec présent",
    "Source (lien)",
}

CHRONO_HEADERS = {
    "Date début",
    "Date fin (ou 2e date distincte)",
    "Dép.",
    "Catégorie",
    "Événement",
    "Lieu",
    "Précisions sur la date et horaires",
    "Source (lien)",
}


# ============================================================
# OUTILS
# ============================================================

def clean(value):
    if value is None:
        return ""
    return str(value).strip()


def norm(value):
    value = clean(value).lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(
        char for char in value
        if not unicodedata.combining(char)
    )
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def parse_date(value):
    if isinstance(value, datetime):
        return value

    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())

    value = clean(value)
    if not value:
        return None

    formats = (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d/%m/%Y %H:%M:%S",
    )

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass

    match = re.search(r"\b(20\d{2})\b", value)
    if match:
        return datetime(int(match.group(1)), 1, 1)

    return None


def year_from(value):
    parsed = parse_date(value)
    return str(parsed.year) if parsed else ""


def normalize_url(value):
    value = clean(value)
    value = value.rstrip("/")
    return norm(value)


def copy_row_style(ws, source_row, target_row):
    for col in range(1, ws.max_column + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)

        if source.has_style:
            target._style = copy(source._style)

        if source.number_format:
            target.number_format = source.number_format

        if source.alignment:
            target.alignment = copy(source.alignment)

        if source.protection:
            target.protection = copy(source.protection)

    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = (
            ws.row_dimensions[source_row].height
        )


def get_headers(ws):
    return {
        clean(ws.cell(1, col).value): col
        for col in range(1, ws.max_column + 1)
        if clean(ws.cell(1, col).value)
    }


def require_sheet(wb, sheet_name):
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Feuille obligatoire absente du classeur : {sheet_name}"
        )
    return wb[sheet_name]


def ensure_event_headers(ws):
    """Répare les anciens onglets événement dont le statut Apec manque."""
    headers = get_headers(ws)
    status_header = "Statut au 12/11/2026"

    if status_header in headers:
        return headers

    # On insère le statut avant "Prochaine édition connue" pour conserver
    # l'ordre historique des colonnes et la mise en forme du classeur.
    anchor = (
        headers.get("Prochaine édition connue")
        or headers.get("Apec présent")
        or (ws.max_column + 1)
    )

    ws.insert_cols(anchor, 1)
    ws.cell(1, anchor).value = status_header

    # Recopie du style de l'en-tête voisin.
    source_col = anchor + 1 if anchor < ws.max_column else max(anchor - 1, 1)
    source = ws.cell(1, source_col)
    target = ws.cell(1, anchor)
    if source.has_style:
        target._style = copy(source._style)
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)

    return get_headers(ws)


def require_headers(ws, required_headers):
    headers = get_headers(ws)
    missing = [
        header for header in required_headers
        if header not in headers
    ]
    if missing:
        raise ValueError(
            f"Colonnes manquantes dans '{ws.title}' : "
            + ", ".join(missing)
        )
    return headers


def next_number(ws):
    headers = get_headers(ws)
    number_col = headers.get("N°")

    if not number_col:
        return None

    numbers = []
    for row in range(2, ws.max_row + 1):
        value = ws.cell(row, number_col).value
        if isinstance(value, int):
            numbers.append(value)
        elif isinstance(value, float) and value.is_integer():
            numbers.append(int(value))

    return max(numbers, default=0) + 1


# ============================================================
# LECTURE CSV
# ============================================================

def load_csv():
    if not INPUT_CSV.exists():
        raise FileNotFoundError(
            f"CSV introuvable : {INPUT_CSV}"
        )

    with INPUT_CSV.open(
        "r",
        encoding="utf-8-sig",
        newline=""
    ) as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames:
            raise ValueError("Le CSV ne contient aucune colonne.")

        rows = list(reader)

    return rows


# ============================================================
# DONNÉES ÉVÉNEMENT
# ============================================================

def choose_decision(row):
    decision = clean(row.get("decision_ia"))

    if decision:
        return decision

    statut_final = clean(row.get("statut_final"))

    if statut_final == "Pertinent":
        return "PERTINENT"

    if statut_final == "Non pertinent":
        return "NON_PERTINENT"

    return "A_VERIFIER"


# ============================================================
# DOUBLE FILTRE DE PERTINENCE
# 1) filtre anti-bruit avant insertion
# 2) contrôle final après qualification IA
# ============================================================

ALLOWED_CATEGORIES = {
    "emploi",
    "recrutement",
    "formation",
    "enseignement_superieur",
    "alumni",
    "portes_ouvertes",
}

# CIBLE APEC : cadres + jeunes diplômés / étudiants du supérieur.
# On distingue les signaux de l'événement des signaux qui prouvent
# qu'il concerne réellement le public Apec.
EVENT_SIGNAL_PATTERNS = (
    r"\bforum\b",
    r"\bsalon\b",
    r"\bjob[ -]?dating\b",
    r"\brecrutement\b",
    r"\bemploi\b",
    r"\bformation\b",
    r"\bcareer\b",
    r"\bcarri[èe]re\b",
    r"\bportes? ouvertes?\b",
    r"\balumni\b",
    r"\brencontre(s)?\b",
    r"\bjourn[ée]e\b",
)

# Signaux forts de cible Apec.
APEC_TARGET_PATTERNS = (
    # Cadres / marché cadre
    r"\bcadres?\b",
    r"\bcadre(s)? dirigeant(s)?\b",
    r"\bemploi(s)? cadres?\b",
    r"\brecrutement(s)? cadres?\b",
    r"\bmanagement\b",
    r"\bmanager(s)?\b",
    # Jeunes diplômés / niveau supérieur
    r"\bjeunes? diplôm[ée]s?\b",
    r"\bdiplôm[ée]s?\b",
    r"\bbac\s*\+\s*[3-8]\b",
    r"\bbac\+\s*[3-8]\b",
    r"\bmaster(s)?\b",
    r"\bdoctorat\b",
    r"\bdoctorants?\b",
    r"\bjeunes? talents?\b",
    # Enseignement supérieur à potentiel Apec
    r"\benseignement supérieur\b",
    r"\buniversit[ée]\b",
    r"\buniversit[ée]s\b",
    r"\bIAE\b",
    r"\balumni\b",
    r"\b[ée]cole d[' ]ing[ée]nieurs?\b",
    r"\b[ée]cole de commerce\b",
    r"\b[ée]coles? de commerce\b",
    r"\bgrande[s]? [ée]cole[s]?\b",
    r"\b[ée]tudiants?\b.*\b(master|bac\s*\+|ing[ée]nieur|universit[ée]|IAE)\b",
    # Mobilité / évolution professionnelle
    r"\b[ée]volution professionnelle\b",
    r"\breconversion professionnelle\b",
    r"\btransition professionnelle\b",
    r"\bmobilit[ée] professionnelle\b",
    r"\bprojet professionnel\b",
    r"\bcarri[èe]re[s]?\b",
    # Recrutement de profils supérieurs
    r"\bprofils? cadres?\b",
    r"\bprofils? [ée]tudiants?\b",
    r"\bprofils? dipl[ôo]m[ée]s?\b",
    r"\bing[ée]nieur(s)?\b",
    r"\bfonctions? support\b",
    r"\bfonctions? [ée]xpertise\b",
    r"\bressources humaines\b",
    r"\bRH\b",
    # Apec explicite / dispositifs très proches
    r"\bAPEC\b",
    r"\bapec connect\b",
    r"\bcadre(s)? / jeunes? diplôm[ée]s?\b",
)

# Contexte supérieur suffisamment qualifié : on ne retient pas un simple
# salon étudiant post-bac généraliste uniquement parce qu'il contient
# "étudiant".
SUPERIOR_CONTEXT_PATTERNS = (
    r"\benseignement supérieur\b",
    r"\buniversit[ée]\b",
    r"\bIAE\b",
    r"\bmaster(s)?\b",
    r"\bbac\s*\+\s*[3-8]\b",
    r"\bbac\+\s*[3-8]\b",
    r"\b[ée]cole d[' ]ing[ée]nieurs?\b",
    r"\b[ée]cole de commerce\b",
    r"\bgrande[s]? [ée]cole[s]?\b",
    r"\balumni\b",
    r"\bdoctorat\b",
    r"\bdoctorants?\b",
)

# Bruit éditorial / articles qui parlent d'un événement sans être
# l'événement lui-même.
NOISE_PATTERNS = (
    r"\bgoogle news\b",
    r"\bactualit[ée]s?\b",
    r"\ben images\b",
    r"\bportrait\b",
    r"\binterview\b",
    r"\btribune\b",
    r"\b[ée]dito\b",
    r"\bchevilles? ouvri[èe]res?\b",
    r"\bla famille\b",
    r"\bjeux intercommunaux\b",
    r"\bpalmar[èe]s\b",
    r"\brecensement\b",
    r"\bvie locale\b",
    r"\bretour sur\b",
)

PRESS_SUFFIXES = (
    " - charente libre",
    " - centre presse",
    " - la nouvelle république",
    " - sud ouest",
    " - le populaire du centre",
)

# Hors cible Apec, même si le contenu contient "emploi", "forum" ou
# "recrutement". On garde volontairement les événements supérieurs et
# cadres lorsqu'ils sont explicitement présents.
APEC_EXCLUSION_PATTERNS = (
    r"\bjob(s)? étudiant(s)?\b",
    r"\bemploi(s)? étudiant(s)?\b",
    r"\bjobs? saisonnier(s)?\b",
    r"\bemploi(s)? saisonnier(s)?\b",
    r"\brecrutement(s)? saisonnier(s)?\b",
    r"\bpetits? boulots?\b",
    r"\bjobs? d[' ]?été\b",
    r"\bjob(s)? vacances\b",
    r"\bint[ée]rim\b",
    r"\bint[ée]rimaires?\b",
    r"\brestauration\b.*\bjob[ -]?dating\b",
    r"\bjob[ -]?dating\b.*\brestauration\b",
    r"\bvendanges?\b",
    r"\bplongeur(s)?\b",
    r"\bserveur(s)?\b",
    r"\bcaissier(s)?\b",
    r"\bcaissi[èe]re(s)?\b",
    r"\bmanutentionnaire(s)?\b",
    r"\bouvrier(s)?\b",
    r"\bagent(s)? d[' ]entretien\b",
    r"\bemplois? sans qualification\b",
)


def _text_for_filter(row, data=None):
    values = (
        row.get("titre_page"), row.get("titre"), row.get("description_page"),
        row.get("description"), row.get("categorie_ia"), row.get("categorie"),
        row.get("organisateur_page"), row.get("lieu_page"),
    )
    if data:
        values += (data.get("title"), data.get("details"))
    return norm(" ".join(clean(v) for v in values if clean(v)))


def _has_pattern(text, patterns):
    return any(re.search(pattern, text, re.IGNORECASE) for pattern in patterns)


def _apec_target_level(text):
    """Retourne un niveau de pertinence Apec déterministe :
    2 = cible Apec explicite, 1 = cible supérieure probable, 0 = hors cible.
    """
    if _has_pattern(text, APEC_TARGET_PATTERNS):
        return 2
    if _has_pattern(text, SUPERIOR_CONTEXT_PATTERNS):
        return 1
    return 0


def prefilter_candidate(row):
    """Premier tri : bruit + événement + premier signal de cible Apec."""
    title = norm(row.get("titre_page") or row.get("titre"))
    category = norm(row.get("categorie_ia") or row.get("categorie"))
    text = _text_for_filter(row)

    if not title:
        return False, "Titre absent."

    if _has_pattern(title, NOISE_PATTERNS):
        return False, "Contenu éditorial / bruit."

    if any(title.endswith(suffix) for suffix in PRESS_SUFFIXES):
        return False, "Article de presse, pas un intitulé événement."

    if _has_pattern(text, APEC_EXCLUSION_PATTERNS) and not _has_pattern(text, APEC_TARGET_PATTERNS):
        return False, "Événement hors cible Apec."

    if category and category not in ALLOWED_CATEGORIES:
        if category not in {
            "emploi & formation", "alumni / réseau", "salon étudiant / orientation", "portes ouvertes"
        }:
            return False, "Catégorie hors périmètre Apec."

    if not _has_pattern(text, EVENT_SIGNAL_PATTERNS):
        return False, "Aucun événement identifiable."

    if _apec_target_level(text) == 0:
        return False, "Aucun signal de cible cadres / jeunes diplômés / supérieur qualifié."

    return True, ""


def final_event_filter(data, row):
    """Second tri indépendant : valide la vraie pertinence Apec."""
    title = norm(data.get("title"))
    details = norm(data.get("details"))
    combined = _text_for_filter(row, data)
    category = norm(row.get("categorie_ia") or row.get("categorie"))

    if not title:
        return False, "Titre absent."

    if _has_pattern(title, NOISE_PATTERNS):
        return False, "Titre éditorial rejeté."

    if any(title.endswith(suffix) for suffix in PRESS_SUFFIXES):
        return False, "Article de presse rejeté."

    if _has_pattern(combined, APEC_EXCLUSION_PATTERNS) and not _has_pattern(combined, APEC_TARGET_PATTERNS):
        return False, "Hors cible Apec."

    if category and category not in ALLOWED_CATEGORIES and category not in {
        "emploi & formation", "alumni / réseau", "salon étudiant / orientation", "portes ouvertes"
    }:
        return False, "Catégorie hors périmètre."

    if not _has_pattern(combined, EVENT_SIGNAL_PATTERNS):
        return False, "Ce n'est pas un événement identifiable."

    target_level = _apec_target_level(combined)
    if target_level == 0:
        return False, "Cible Apec absente : pas de signal cadres / jeunes diplômés / supérieur qualifié."

    # Un événement supérieur généraliste doit au minimum comporter un
    # contexte qualifié (université, master, ingénieur, alumni, etc.).
    # Cela évite de transformer tout salon étudiant post-bac en événement Apec.
    if target_level == 1 and not _has_pattern(combined, SUPERIOR_CONTEXT_PATTERNS):
        return False, "Enseignement supérieur insuffisamment qualifié."

    return True, ""


def choose_title(row):
    return (
        clean(row.get("titre_page"))
        or clean(row.get("titre"))
    )


def choose_source(row):
    return (
        clean(row.get("url_finale"))
        or clean(row.get("url"))
    )


def choose_date(row):
    # Priorité absolue à la date validée par l'IA.
    for key in (
        "date_ia",
        "date_evenement_page",
    ):
        parsed = parse_date(row.get(key))
        if parsed:
            return parsed

    # Fallback : date détectée par le programme.
    year = clean(row.get("annee_detectee"))
    match = re.search(r"\b(20\d{2})\b", year)
    if match:
        return datetime(int(match.group(1)), 1, 1)

    # Ne jamais utiliser la date de publication comme date
    # d'événement : elle sert uniquement à l'audit.
    return None


def choose_year(row, event_date):
    if event_date:
        return str(event_date.year)

    for key in ("date_ia", "date_evenement_page"):
        year = year_from(row.get(key))
        if year:
            return year

    for key in ("annee_detectee", "annee_recherchee"):
        value = clean(row.get(key))
        match = re.search(r"\b(20\d{2})\b", value)
        if match:
            return match.group(1)

    return ""


def choose_department(row):
    for key in ("departement_ia", "departement"):
        value = clean(row.get(key))
        if value in DEPT_SHEETS:
            return value
    return ""


def choose_category(row):
    value = clean(row.get("categorie_ia"))

    mapping = {
        "emploi": "Emploi & formation",
        "recrutement": "Emploi & formation",
        "formation": "Emploi & formation",
        "enseignement_superieur": "Salon étudiant / orientation",
        "alumni": "Alumni / réseau",
        "portes_ouvertes": "Portes ouvertes",
    }

    if value in mapping:
        return mapping[value]

    if value:
        return value

    fallback = clean(row.get("categorie"))
    return fallback or "À vérifier"


def choose_location(row):
    return clean(row.get("lieu_page"))


def choose_details(row):
    return (
        clean(row.get("raison_ia"))
        or clean(row.get("description_page"))
        or clean(row.get("description"))
        or clean(row.get("raison"))
    )


def choose_status(event_date):
    if not event_date:
        return "À vérifier"

    today = datetime.now().date()

    if event_date.date() >= today:
        return "À venir"

    return "Passé"


def choose_apec_presence(row):
    title = norm(row.get("titre_page") or row.get("titre"))
    source = norm(row.get("url_finale") or row.get("url"))
    organisation = norm(row.get("organisateur_page"))

    if (
        "apec" in title
        or "event.apec.fr" in source
        or "apec" in organisation
    ):
        return "Oui"

    return "Non confirmé"


def build_event(row):
    event_date = choose_date(row)
    department = choose_department(row)

    return {
        "title": choose_title(row),
        "source": choose_source(row),
        "date": event_date,
        "year": choose_year(row, event_date),
        "department": department,
        "category": choose_category(row),
        "location": choose_location(row),
        "details": choose_details(row),
        "status": choose_status(event_date),
        "apec": choose_apec_presence(row),
        "decision": choose_decision(row),
    }


# ============================================================
# IDENTIFICATION / DÉDOUBLONNAGE
# ============================================================

def event_key(data):
    source = normalize_url(data.get("source"))

    if source:
        return "url|" + source

    return "event|{}|{}|{}".format(
        norm(data.get("title")),
        clean(data.get("department")),
        clean(data.get("year")),
    )


def sheet_row_key(ws, row_number):
    headers = get_headers(ws)

    url_col = headers.get("Source (lien)")
    title_col = headers.get("Événement")
    dep_col = headers.get("Dép.")
    date_col = headers.get("Date début")

    source = (
        ws.cell(row_number, url_col).value
        if url_col
        else ""
    )

    title = (
        ws.cell(row_number, title_col).value
        if title_col
        else ""
    )

    department = (
        ws.cell(row_number, dep_col).value
        if dep_col
        else ""
    )

    event_date = (
        ws.cell(row_number, date_col).value
        if date_col
        else ""
    )

    source = normalize_url(source)

    if source:
        return "url|" + source

    return "event|{}|{}|{}".format(
        norm(title),
        clean(department),
        year_from(event_date),
    )


def build_sheet_index(ws):
    index = {}

    for row in range(2, ws.max_row + 1):
        key = sheet_row_key(ws, row)

        if key != "event|||":
            index[key] = row

    return index


# ============================================================
# ÉCRITURE DES FEUILLES PRINCIPALES
# ============================================================

def write_event_row(ws, row_number, data, update_existing=False):
    headers = get_headers(ws)

    values = {
        "Dép.": int(data["department"])
        if data["department"].isdigit()
        else data["department"],

        "Catégorie": data["category"],
        "Événement": data["title"],
        "Date début": data["date"],
        "Source (lien)": data["source"],
    }

    # Pour une ligne existante, on ne remplace pas les champs
    # éditoriaux manuels si la nouvelle information est vide.
    optional_values = {
        "Lieu": data["location"],
        "Détails et récurrence": data["details"],
        "Statut au 12/11/2026": data["status"],
        "Apec présent": data["apec"],
    }

    for name, value in values.items():
        col = headers.get(name)
        if col:
            ws.cell(row_number, col).value = value

    for name, value in optional_values.items():
        col = headers.get(name)
        if not col:
            continue

        if update_existing and not clean(value):
            continue

        ws.cell(row_number, col).value = value

    if not update_existing:
        number_col = headers.get("N°")
        if number_col:
            ws.cell(
                row_number,
                number_col
            ).value = next_number(ws)


def update_or_append_event(ws, data, index):
    key = event_key(data)

    if key in index:
        row_number = index[key]

        write_event_row(
            ws,
            row_number,
            data,
            update_existing=True,
        )

        return "updated"

    row_number = ws.max_row + 1

    if ws.max_row >= 2:
        copy_row_style(ws, ws.max_row, row_number)

    write_event_row(
        ws,
        row_number,
        data,
        update_existing=False,
    )

    index[key] = row_number

    return "added"


# ============================================================
# CHRONOLOGIE
# ============================================================

def chronology_key(ws, row_number):
    headers = get_headers(ws)

    date_col = headers["Date début"]
    dep_col = headers["Dép."]
    title_col = headers["Événement"]

    return (
        norm(ws.cell(row_number, title_col).value),
        clean(ws.cell(row_number, dep_col).value),
        year_from(ws.cell(row_number, date_col).value),
    )


def build_chronology_index(ws):
    index = set()

    for row in range(2, ws.max_row + 1):
        index.add(chronology_key(ws, row))

    return index


def add_to_chronology(ws, data, index):
    if not data["date"]:
        return False

    # La chronologie contient les événements à venir.
    if data["date"].date() < datetime.now().date():
        return False

    key = (
        norm(data["title"]),
        clean(data["department"]),
        data["year"],
    )

    if key in index:
        return False

    row_number = ws.max_row + 1

    if ws.max_row >= 2:
        copy_row_style(ws, ws.max_row, row_number)

    headers = get_headers(ws)

    values = {
        "Date début": data["date"],
        "Dép.": int(data["department"])
        if data["department"].isdigit()
        else data["department"],
        "Catégorie": data["category"],
        "Événement": data["title"],
        "Lieu": data["location"],
        "Source (lien)": data["source"],
    }

    for name, value in values.items():
        col = headers.get(name)
        if col:
            ws.cell(row_number, col).value = value

    index.add(key)
    return True


# ============================================================
# FEUILLES 2 COLONNES
# ============================================================

def append_audit_row(ws, title, reason):
    title = clean(title)
    reason = clean(reason)

    if not title:
        return False

    existing = set()

    for row in range(2, ws.max_row + 1):
        existing.add(
            (
                norm(ws.cell(row, 1).value),
                norm(ws.cell(row, 2).value),
            )
        )

    key = (norm(title), norm(reason))

    if key in existing:
        return False

    row_number = ws.max_row + 1

    if ws.max_row >= 2:
        copy_row_style(ws, ws.max_row, row_number)

    ws.cell(row_number, 1).value = title
    ws.cell(row_number, 2).value = reason

    return True


# ============================================================
# VÉRIFICATION DU CLASSEUR
# ============================================================

def validate_workbook(wb):
    required_sheets = [
        "Synthèse",
        CHRONOLOGY_SHEET,
        *YEAR_SHEETS.values(),
        *DEPT_SHEETS.values(),
        "En ligne",
        "Contacts",
        "Pistes (Phase 8)",
        NON_PERTINENT_SHEET,
        NON_VERIFIABLE_SHEET,
    ]

    for sheet_name in required_sheets:
        require_sheet(wb, sheet_name)

    require_headers(
        wb[CHRONOLOGY_SHEET],
        CHRONO_HEADERS,
    )

    for sheet_name in YEAR_SHEETS.values():
        ensure_event_headers(wb[sheet_name])
        require_headers(wb[sheet_name], EVENT_HEADERS)

    for sheet_name in DEPT_SHEETS.values():
        ensure_event_headers(wb[sheet_name])
        require_headers(wb[sheet_name], EVENT_HEADERS)

    require_headers(
        wb[NON_PERTINENT_SHEET],
        {"Source ou élément", "Raison"},
    )

    require_headers(
        wb[NON_VERIFIABLE_SHEET],
        {"Élément", "Raison"},
    )


# ============================================================
# UNIFORMISATION DE LA MISE EN FORME
# Basée sur la charte du classeur original
# ============================================================

EVENT_SHEET_NAMES = [*YEAR_SHEETS.values(), *DEPT_SHEETS.values(), "En ligne"]


def _copy_style_from(ws, source_row, source_col, target_cell):
    source = ws.cell(source_row, source_col)
    if source.has_style:
        target_cell._style = copy(source._style)
    if source.number_format:
        target_cell.number_format = source.number_format
    target_cell.alignment = copy(source.alignment)
    target_cell.protection = copy(source.protection)


def _set_column_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _uniformize_event_sheet(ws, template_ws):
    # Le classeur original sert de modèle : mêmes largeurs, en-têtes,
    # filtres et comportement de lecture sur toutes les feuilles événement.
    # Base sur le tableur original, puis élargissement contrôlé pour éviter
    # les intitulés tronqués et les colonnes trop serrées.
    widths = {
        "A": 7, "B": 10, "C": 28, "D": 65, "E": 15, "F": 22,
        "G": 38, "H": 35, "I": 80, "J": 22, "K": 25, "L": 16, "M": 65,
    }
    _set_column_widths(ws, widths)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:M{max(ws.max_row, 1)}"

    # En-tête : copie exacte du style original.
    for col in range(1, 14):
        _copy_style_from(template_ws, 1, col, ws.cell(1, col))
    ws.row_dimensions[1].height = template_ws.row_dimensions[1].height

    # Style de référence des colonnes texte/source du classeur original.
    text_style_row = 2
    for row in range(2, ws.max_row + 1):
        # N° et département
        _copy_style_from(template_ws, text_style_row, 1, ws.cell(row, 1))
        _copy_style_from(template_ws, text_style_row, 2, ws.cell(row, 2))

        # Colonnes texte
        for col in (3, 4, 7, 8, 9, 10, 12):
            _copy_style_from(template_ws, text_style_row, 3, ws.cell(row, col))

        # Dates : une date complète reprend le format date du modèle ;
        # une cellule vide garde le style neutre du modèle.
        for col in (5, 6, 11):
            cell = ws.cell(row, col)
            if cell.value is None or cell.value == "":
                _copy_style_from(template_ws, 4, 6, cell)
            else:
                _copy_style_from(template_ws, 4, 5, cell)

        # Source : style spécifique des liens du classeur original.
        _copy_style_from(template_ws, text_style_row, 13, ws.cell(row, 13))

        # Format de date standardisé dans tout le classeur.
        for col in (5, 6, 11):
            cell = ws.cell(row, col)
            if cell.value not in (None, ""):
                cell.number_format = "dd/mm/yyyy"

        # Lecture homogène des cellules longues.
        for col in range(1, 14):
            cell = ws.cell(row, col)
            alignment = copy(cell.alignment)
            alignment.wrap_text = True
            alignment.vertical = "top"
            cell.alignment = alignment

        ws.row_dimensions[row].height = None


def _uniformize_chronology(ws, template_ws):
    widths = {
        "A": 15, "B": 22, "C": 10, "D": 28,
        "E": 65, "F": 35, "G": 42, "H": 65,
    }
    _set_column_widths(ws, widths)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(ws.max_row, 1)}"

    for col in range(1, 9):
        _copy_style_from(template_ws, 1, col, ws.cell(1, col))
    ws.row_dimensions[1].height = template_ws.row_dimensions[1].height

    for row in range(2, ws.max_row + 1):
        for col in (1, 2):
            cell = ws.cell(row, col)
            if cell.value is None or cell.value == "":
                _copy_style_from(template_ws, 3, col, cell)
            else:
                _copy_style_from(template_ws, 2, 1, cell)
        _copy_style_from(template_ws, 2, 3, ws.cell(row, 3))
        for col in (4, 5, 6, 7):
            _copy_style_from(template_ws, 2, 4, ws.cell(row, col))
        _copy_style_from(template_ws, 2, 8, ws.cell(row, 8))
        for col in (1, 2):
            if ws.cell(row, col).value not in (None, ""):
                ws.cell(row, col).number_format = "dd/mm/yyyy"
        for col in range(1, 9):
            cell = ws.cell(row, col)
            alignment = copy(cell.alignment)
            alignment.wrap_text = True
            alignment.vertical = "top"
            cell.alignment = alignment
        ws.row_dimensions[row].height = None


def _uniformize_simple_sheet(ws, template_ws):
    # Pour les feuilles à 2/5 colonnes, on harmonise la structure sans
    # écraser les éventuels surlignages métier déjà présents dans l'original.
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        _copy_style_from(template_ws, 1, min(col, template_ws.max_column), ws.cell(1, col))
    ws.row_dimensions[1].height = template_ws.row_dimensions[1].height

    if max_col <= 5:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max(ws.max_row, 1)}"

    for row in range(2, ws.max_row + 1):
        for col in range(1, max_col + 1):
            # Style texte standard du modèle, sauf si la cellule originale
            # possède déjà un style métier distinct : celui-ci est conservé.
            if ws.cell(row, col).style_id == 0:
                _copy_style_from(template_ws, 2, min(col, template_ws.max_column), ws.cell(row, col))
        ws.row_dimensions[row].height = None


def uniformize_workbook(wb):
    """Applique la charte visuelle du classeur original à chaque exécution."""
    template_event = wb["Événements 2026"]
    template_chrono = wb[CHRONOLOGY_SHEET]

    for sheet_name in EVENT_SHEET_NAMES:
        if sheet_name in wb.sheetnames:
            _uniformize_event_sheet(wb[sheet_name], template_event)

    if CHRONOLOGY_SHEET in wb.sheetnames:
        _uniformize_chronology(wb[CHRONOLOGY_SHEET], template_chrono)

    if "Contacts" in wb.sheetnames:
        _uniformize_simple_sheet(wb["Contacts"], wb["Contacts"])

    for sheet_name in (NON_PERTINENT_SHEET, NON_VERIFIABLE_SHEET):
        if sheet_name in wb.sheetnames:
            _uniformize_simple_sheet(wb[sheet_name], wb[sheet_name])

    # Synthèse et Pistes restent basées sur leur mise en page d'origine :
    # on ne réécrit pas leurs styles métier, mais on homogénéise les hauteurs.
    for sheet_name in ("Synthèse", "Pistes (Phase 8)"):
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in range(1, ws.max_row + 1):
                if ws.row_dimensions[row].height is not None:
                    ws.row_dimensions[row].height = None



# ============================================================
# TRAITEMENT PRINCIPAL
# ============================================================

def main():
    print("========================================")
    print("SYNCHRONISATION CSV -> EXCEL")
    print("========================================")

    if not WORKBOOK.exists():
        raise FileNotFoundError(
            f"Classeur introuvable : {WORKBOOK}"
        )

    rows = load_csv()

    print(f"CSV : {INPUT_CSV}")
    print(f"Classeur : {WORKBOOK}")
    print(f"Lignes CSV : {len(rows)}")

    wb = load_workbook(WORKBOOK)
    validate_workbook(wb)

    indexes = {}

    for sheet_name in YEAR_SHEETS.values():
        indexes[sheet_name] = build_sheet_index(
            wb[sheet_name]
        )

    for sheet_name in DEPT_SHEETS.values():
        indexes[sheet_name] = build_sheet_index(
            wb[sheet_name]
        )

    chronology_index = build_chronology_index(
        wb[CHRONOLOGY_SHEET]
    )

    stats = {
        "ajoutes": 0,
        "mis_a_jour": 0,
        "chronologie_ajoutee": 0,
        "non_pertinent_ajoute": 0,
        "non_verifiable_ajoute": 0,
        "ignores": 0,
    }

    for raw in rows:
        # TRI 1 : filtre déterministe anti-bruit avant tout ajout au classeur.
        pre_ok, pre_reason = prefilter_candidate(raw)
        if not pre_ok:
            stats["ignores"] += 1
            continue

        data = build_event(raw)
        decision = data["decision"]

        if decision == "NON_PERTINENT":
            stats["ignores"] += 1
            continue

        if decision == "PERTINENT":
            # TRI 2 : contrôle final indépendant de la décision IA.
            final_ok, final_reason = final_event_filter(data, raw)
            if not final_ok:
                stats["ignores"] += 1
                continue
            added = append_audit_row(
                wb[NON_PERTINENT_SHEET],
                data["title"],
                clean(
                    raw.get("raison_ia")
                    or raw.get("raison")
                    or "Classé non pertinent par l'IA."
                ),
            )

            if added:
                stats["non_pertinent_ajoute"] += 1

            continue

        if decision in (
            "A_VERIFIER",
            "IA_INDISPONIBLE",
            "IA_NON_ANALYSE",
            "",
        ):
            added = append_audit_row(
                wb[NON_VERIFIABLE_SHEET],
                data["title"],
                clean(
                    raw.get("raison_ia")
                    or raw.get("raison")
                    or "Validation nécessaire."
                ),
            )

            if added:
                stats["non_verifiable_ajoute"] += 1

            continue

        if decision != "PERTINENT":
            stats["ignores"] += 1
            continue

        if not data["title"]:
            stats["ignores"] += 1
            continue

        if not data["department"]:
            # Un événement pertinent mais sans département
            # vérifié ne doit pas être placé arbitrairement.
            append_audit_row(
                wb[NON_VERIFIABLE_SHEET],
                data["title"],
                "Département non vérifié.",
            )
            stats["non_verifiable_ajoute"] += 1
            continue

        if data["year"] not in YEAR_SHEETS:
            append_audit_row(
                wb[NON_VERIFIABLE_SHEET],
                data["title"],
                "Année d'événement non gérée par le classeur.",
            )
            stats["non_verifiable_ajoute"] += 1
            continue

        year_sheet = YEAR_SHEETS[data["year"]]
        dept_sheet = DEPT_SHEETS[data["department"]]

        result = update_or_append_event(
            wb[year_sheet],
            data,
            indexes[year_sheet],
        )

        if result == "added":
            stats["ajoutes"] += 1
        else:
            stats["mis_a_jour"] += 1

        # Même événement dans la feuille départementale.
        result = update_or_append_event(
            wb[dept_sheet],
            data,
            indexes[dept_sheet],
        )

        if result == "added":
            stats["ajoutes"] += 1
        else:
            stats["mis_a_jour"] += 1

        if add_to_chronology(
            wb[CHRONOLOGY_SHEET],
            data,
            chronology_index,
        ):
            stats["chronologie_ajoutee"] += 1

    # Uniformisation finale : le classeur original sert de référence de mise en forme.
    uniformize_workbook(wb)

    wb.save(WORKBOOK)

    print()
    print("========================================")
    print("SYNCHRONISATION TERMINÉE")
    print("========================================")

    for key, value in stats.items():
        print(f"{key}: {value}")

    print(f"Classeur mis à jour : {WORKBOOK}")


if __name__ == "__main__":
    main()
