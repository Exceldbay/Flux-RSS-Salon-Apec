from pathlib import Path
import yaml
import pandas as pd
from datetime import datetime


ROOT = Path(__file__).resolve().parent.parent

CONFIG = ROOT / "config"
DATA = ROOT / "data"
OUTPUT = ROOT / "output"


def load_yaml(filename):
    path = CONFIG / filename

    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def main():

    DATA.mkdir(exist_ok=True)
    OUTPUT.mkdir(exist_ok=True)

    departements = load_yaml("departements.yml")
    categories = load_yaml("categories.yml")
    exclusions = load_yaml("exclusions.yml")
    operators = load_yaml("operators.yml")

    print("====================================")
    print("VEILLE APEC - DEMARRAGE")
    print("====================================")

    print("\nDépartements :")

    for code, nom in departements["departements"].items():
        print(f"- {code} : {nom}")

    print("\nAnnées :")

    for annee in departements["annees"]:
        print(f"- {annee}")

    print("\nCatégories :")

    for numero, categorie in categories["categories"].items():
        print(
            f"- {numero} : "
            f"{categorie['nom']} "
            f"({categorie['poids_cible']} %)"
        )

    print("\nNombre d'exclusions :")

    total_exclusions = sum(
        len(values)
        for values in exclusions["exclusions"].values()
    )

    print(total_exclusions)

    print("\nNombre d'opérateurs intitle :")
    print(len(operators["intitle"]))

    maintenant = datetime.now().isoformat()

    test = pd.DataFrame([
        {
            "date_execution": maintenant,
            "statut": "TEST",
            "message": "Pipeline GitHub opérationnel"
        }
    ])

    fichier = DATA / "test_pipeline.csv"

    test.to_csv(
        fichier,
        index=False,
        encoding="utf-8-sig"
    )

    print("\nFichier créé :")
    print(fichier)

    print("\n====================================")
    print("VEILLE APEC - FIN")
    print("====================================")


if __name__ == "__main__":
    main()
