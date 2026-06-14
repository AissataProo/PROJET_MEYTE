import pandas as pd
from loguru import logger
from config import URLS


def load_data() -> pd.DataFrame:
    """Charge le CSV depuis l'URL publique MinIO définie dans la config."""
    logger.info(f"Chargement des données depuis {URLS}")
    return pd.read_csv(URLS)


def get_cleaned_effectifs() -> pd.DataFrame:
    """Charge, fusionne et nettoie les données d'effectifs."""
    logger.info("Chargement et nettoyage des effectifs...")
    
    usecols = [
        "annee",
        "dept",
        "sexe",
        "patho_niv1",
        "patho_niv2",
        "prev",
        "Npop",
        "Ntop",
        "libelle_classe_age",
    ]
    
    chunks = []
    for chunk in pd.read_csv(
        URLS["effectifs"],
        sep=";",
        chunksize=100_000,
        low_memory=False,
        usecols=usecols,
    ):
        filtered = chunk[
            (chunk["annee"].isin([2021, 2022, 2023]))
            & (chunk["dept"] != "999")
            & (chunk["prev"] > 0)
            & (chunk["patho_niv1"] != "0")
            & (chunk["patho_niv2"] != "0")
        ]
        chunks.append(filtered)
    
    df = pd.concat(chunks, ignore_index=True)
    
    def convert_sexe(val):
        if val == 1:
            return "HOMME"
        elif val == 2:
            return "FEMME"
        else:
            return "ENSEMBLE"
    
    df["sexe"] = df["sexe"].apply(convert_sexe)
    logger.info(f"Sexe converti: {df['sexe'].value_counts().to_dict()}")
    
    df_regions = pd.read_csv(
        URLS["Régions"],
        sep=";",
        encoding="latin1",
        usecols=["departement", "libelle_departement", "libelle_region"],
    )
    
    df_regions["departement"] = df_regions["departement"].astype(str).str.zfill(2)
    df["dept"] = df["dept"].astype(str).str.zfill(2)
    df = pd.merge(df, df_regions, left_on="dept", right_on="departement", how="inner")
    
    df["patho_niv2"] = (
        df["patho_niv2"]
        .astype(str)
        .str.replace(r"\s*\(.*\)", "", regex=True)
        .str.strip()
    )
    
    hors_hexa = [
        "Tout département",
        "Haute-Corse",
        "Martinique",
        "La Réunion",
        "Guyane",
        "Mayotte",
        "Corse-du-Sud",
        "FRANCE",
    ]
    
    df = df[(~df["libelle_departement"].isin(hors_hexa)) & (df["libelle_region"] != "FRANCE")]
    
    df = (
        df.groupby(
            ["annee", "sexe", "dept", "patho_niv1", "patho_niv2", "libelle_classe_age", "departement", "libelle_departement", "libelle_region"],
            as_index=False,
        )[["prev", "Npop", "Ntop"]]
        .sum()
    )
    
    return (
        df.rename(
            columns={
                "libelle_departement": "Département",
                "libelle_region": "Région",
                "prev": "Prévalence",
                "Npop": "Population de référence",
                "Ntop": "Effectif",
                "libelle_classe_age": "Classe d'age",
                "sexe": "Sexe",
            }
        )
        .drop(
            columns=[
                "dept",
                "departement",
            ],
            errors="ignore",
        )
    )

def get_cleaned_depenses() -> pd.DataFrame:
    """Charge et nettoie les données de dépenses."""
    logger.info("Chargement et nettoyage des dépenses...")

    chunks = []
    for chunk in pd.read_csv(
        URLS["depenses"], sep=";", chunksize=100_000, low_memory=False
    ):
        filtered = chunk[(chunk["annee"] >= 2022) & (chunk["montant"] > 0)]
        chunks.append(filtered)

    df = pd.concat(chunks, ignore_index=True)

    df = df.rename(
        columns={
            "Ntop": "Effectif",
            "dep_niv_1": "poste de dépense",
            "dep_niv_2": "sous poste",
        }
    )

    cols_to_drop = ["tri", "Niveau prioritaire", "type_somme", "top"]
    df = df.drop(columns=cols_to_drop, errors="ignore")

    lignes_a_exclure = ["Total consommants tous régimes", "Soins courants", "nan"]
    df = df[~df["sous poste"].isin(lignes_a_exclure)]

    df = df.drop_duplicates()
    return df


def compute_list_lengths(
    df_depenses: pd.DataFrame, df_effectifs: pd.DataFrame
) -> dict[str, int]:
    """Calcule le nombre de valeurs uniques par colonne pour les formules Excel."""
    len_dict = {}

    pathos = (
        df_depenses.loc[
            ~df_depenses["patho_niv1"].isin(
                ["Total", "Total consommants tous régimes"]
            ),
            "patho_niv1",
        ]
        .dropna()
        .unique()
    )

    len_dict["len_annee"] = len(df_depenses["annee"].dropna().unique()) + 1
    len_dict["len_patho_niv1"] = len(pathos) + 1
    len_dict["len_poste de dépense"] = (
        len(df_depenses["poste de dépense"].dropna().unique()) + 1
    )
    len_dict["len_sous poste"] = len(df_depenses["sous poste"].dropna().unique()) + 1

    classes_age = (
        df_effectifs.loc[
            ~df_effectifs["Classe d'age"]
            .astype(str)
            .str.lower()
            .isin(["tous âges", "tous ages", "total"]),
            "Classe d'age",
        ]
        .dropna()
        .unique()
    )

    pathos_eff = (
        df_effectifs.loc[
            ~df_effectifs["patho_niv1"]
            .astype(str)
            .str.lower()
            .isin(["total", "total consommants tous régimes"]),
            "patho_niv1",
        ]
        .dropna()
        .unique()
    )

    len_dict["len_departement_eff"] = (
        len(df_effectifs["Département"].dropna().unique()) + 1
    )
    len_dict["len_classes_age"] = len(classes_age) + 1
    len_dict["len_region"] = len(df_effectifs["Région"].dropna().unique()) + 1
    len_dict["len_patho_effectif"] = len(pathos_eff) + 1

    return len_dict
