import csv
import re
import unicodedata
from copy import copy
from datetime import datetime, date
from pathlib import Path

from openpyxl import load_workbook


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
        require_headers(wb[sheet_name], EVENT_HEADERS)

    for sheet_name in DEPT_SHEETS.values():
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
    widths = {
        key: template_ws.column_dimensions[key].width
        for key in "ABCDEFGHIJKLM"
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

        ws.row_dimensions[row].height = None


def _uniformize_chronology(ws, template_ws):
    widths = {key: template_ws.column_dimensions[key].width for key in "ABCDEFGH"}
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
        ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(max_col)}{max(ws.max_row, 1)}"

    for row in range(2, ws.max_row + 1):
        for col in range(1, max_col + 1):
            # Style texte standard du modèle, sauf si la cellule originale
            # possède déjà un style métier distinct : celui-ci est conservé.
            if ws.cell(row, col).style_id == 0:
                _copy_style_from(template_ws, 2, min(col, template_ws.max_column), ws.cell(row, col))
        ws.row_dimensions[row].height = None


def uniformize_workbook(wb):
    """Applique la charte visuelle du classeur original à chaque exécution."""
    import openpyxl.utils

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
        data = build_event(raw)
        decision = data["decision"]

        if decision == "NON_PERTINENT":
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
