from pathlib import Path

# --- Chemins du projet ---
PROJECT_ROOT = Path(__file__).parent.parent
DATA_DIR = PROJECT_ROOT / "data"
RESULT_DIR = PROJECT_ROOT / "result"

# # --- Fichiers locaux ---
# DATA_DEPENSES = DATA_DIR / "data_ameli_depenses.csv"
# DATA_EFFECTIFS = DATA_DIR / "data_ameli_effectifs.csv"

# # --- Fichiers de sortie ---
# OUTPUT_DEP = RESULT_DIR / "Dashboard_Depenses.xlsx"
# OUTPUT_EFF = RESULT_DIR / "Dashboard_Effectifs.xlsx"
OUTPUT_DASHBOARD = RESULT_DIR / "Dashboard_Santé.xlsx"

# --- URLs MinIO ---
URL_EFFECTIFS = "https://minio.lab.sspcloud.fr/aissataa/PROJET_MEYTE/effectifs.csv"
URL_REGIONS = (
    "https://minio.lab.sspcloud.fr/aissataa/PROJET_MEYTE/LibelleRegionDept.csv"
)
URL_DEPENSES = "https://minio.lab.sspcloud.fr/aissataa/PROJET_MEYTE/depenses.csv"
URLS = {
    "effectifs": URL_EFFECTIFS,
    "Régions": URL_REGIONS,
    "depenses": URL_DEPENSES,
}
# --- Paramètres de nettoyage ---
YEAR_MIN = 2022
EXCLUDE_DEPT = "999"
EXCLUDE_REG = "99"
MIN_MONTANT = 0
CHUNKSIZE = 100_000
# --- Valeurs par défaut des filtres
DEFAULT_ANNEE = 2023
DEFAULT_PATHO_NIV1 = "Tous"
DEFAULT_POSTE = "Tous"
# --- Noms de sheets ---
SHEET_CLEANED_DEPENSES = "cleanedData_Depenses"
SHEET_CLEANED_EFFECTIFS = "cleanedData_Effectifs"
SHEET_FILTRES = "Filtres"
SHEET_POSTES = "Postes"
SHEET_GRAPHIQUES = "Graphiques"
SHEET_REGION = "Region"
SHEET_DEPARTEMENT = "Departement"

# --- Couleurs principales ---
COULEURS = {
    "principal": "FF004A99",
    "secondaire": "FF008080",
    "accent": "FF4CAF50",
    "gris_clair": "FFF5F5F5",
    "blanc": "FFFFFFFF",
    "noir": "FF000000",
}

# --- Noms de colonnes ---
COLONNES_DEPENSES = {
    "annee": "annee",
    "patho_niv1": "patho_niv1",
    "patho_niv2": "patho_niv2",
    "patho_niv3": "patho_niv3",
    "poste": "poste de dépense",
    "sous_poste": "sous poste",
    "montant": "montant",
    "effectif": "Effectif",
}

COLONNES_EFFECTIFS = {
    "annee": "annee",
    "patho_niv1": "patho_niv1",
    "region": "Région",
    "department": "Département",
    "age": "Classe d'age",
    "sexe": "Sexe",
    "effectif": "Effectif",
}
PIE_COLORS = [
    COULEURS["principal"],
    COULEURS["secondaire"],
    COULEURS["accent"],
    "FF0070AD",
    "FF6372C4",
    "FFA9D480",
]

# --- Styles de texte ---
FONT_BOLD_TITLE = {"bold": True, "size": 13, "color": "333333"}
FONT_HEADER = {"bold": True, "color": "FFFFFF", "size": 9}
FONT_DATA = {"size": 8, "color": "333333"}

# --- Fill (couleur de fond) ---
FILL_HEADER = {"start_color": COULEURS["principal"], "fill_type": "solid"}
FILL_ALT = {"start_color": COULEURS["gris_clair"], "fill_type": "solid"}
FILL_BLANC = {"start_color": COULEURS["blanc"], "fill_type": "solid"}

# --- Limites d'affichage ---
MAX_ROWS_AFFICHAGE = 300
TOP_N = 20
