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

    dtype_map = {
        "annee": "Int16",
        "dept": "string",
        "sexe": "Int8",
        "patho_niv1": "string",
        "patho_niv2": "string",
        "prev": "float32",
        "Npop": "Int32",
        "Ntop": "Int32",
        "libelle_classe_age": "string",
    }

    chunks = []
    for chunk in pd.read_csv(
        URLS["effectifs"],
        sep=";",
        low_memory=False,
        usecols=usecols,
        dtype=dtype_map,
        chunksize=100_000,
    ):
        chunk = chunk[
            (chunk["annee"].isin([2021, 2022, 2023]))
            & (chunk["dept"] != "999")
            & (chunk["prev"] > 0)
            & (chunk["patho_niv1"] != "0")
            & (chunk["patho_niv2"] != "0")
        ].copy()
        chunks.append(chunk)

    df = pd.concat(chunks, ignore_index=True)

    sexe_map = {1: "HOMME", 2: "FEMME"}
    df["sexe"] = df["sexe"].map(sexe_map).fillna("ENSEMBLE").astype("category")
    logger.info(f"Sexe converti: {df['sexe'].value_counts().to_dict()}")

    df_regions = pd.read_csv(
        URLS["Régions"],
        sep=";",
        encoding="latin1",
        usecols=["departement", "libelle_departement", "libelle_region"],
        dtype={
            "departement": "string",
            "libelle_departement": "string",
            "libelle_region": "string",
        },
    )

    df_regions["departement"] = df_regions["departement"].astype("string").str.zfill(2)
    df["dept"] = df["dept"].astype("string").str.zfill(2)

    df = df.merge(df_regions, left_on="dept", right_on="departement", how="inner")

    df["patho_niv2"] = (
        df["patho_niv2"]
        .astype("string")
        .str.replace(r"\s*\(.*\)", "", regex=True)
        .str.strip()
    )

    hors_hexa = {
        "Tout département",
        "Haute-Corse",
        "Martinique",
        "La Réunion",
        "Guyane",
        "Mayotte",
        "Corse-du-Sud",
        "FRANCE",
    }

    df = df[
        (~df["libelle_departement"].isin(hors_hexa))
        & (df["libelle_region"] != "FRANCE")
    ].copy()

    df = df.groupby(
        [
            "annee",
            "sexe",
            "dept",
            "patho_niv1",
            "patho_niv2",
            "libelle_classe_age",
            "departement",
            "libelle_departement",
            "libelle_region",
        ],
        as_index=False,
        observed=True,
    )[["prev", "Npop", "Ntop"]].sum()

    df = df.rename(
        columns={
            "libelle_departement": "Département",
            "libelle_region": "Région",
            "prev": "Prévalence",
            "Npop": "Population de référence",
            "Ntop": "Effectif",
            "libelle_classe_age": "Classe d'age",
            "sexe": "Sexe",
        }
    ).drop(columns=["dept", "departement"], errors="ignore")

    for col in [
        "patho_niv1",
        "patho_niv2",
        "Département",
        "Région",
        "Classe d'age",
        "Sexe",
    ]:
        if col in df.columns:
            df[col] = df[col].astype("category")

    logger.info(f"Effectifs nettoyés: {len(df):,} lignes")
    return df


def get_cleaned_depenses() -> pd.DataFrame:
    """Charge et nettoie les données de dépenses."""
    logger.info("Chargement et nettoyage des dépenses...")

    usecols = [
        "annee",
        "patho_niv1",
        "patho_niv2",
        "dep_niv_1",
        "dep_niv_2",
        "montant",
        "Ntop",
    ]

    dtype_map = {
        "annee": "Int16",
        "patho_niv1": "string",
        "patho_niv2": "string",
        "dep_niv_1": "string",
        "dep_niv_2": "string",
        "montant": "float32",
        "Ntop": "Int32",
    }

    chunks = []
    for chunk in pd.read_csv(
        URLS["depenses"],
        sep=";",
        low_memory=False,
        usecols=usecols,
        dtype=dtype_map,
        chunksize=100_000,
    ):
        chunk = chunk[(chunk["annee"] >= 2022) & (chunk["montant"] > 0)].copy()
        chunks.append(chunk)

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
    if "sous poste" in df.columns:
        df = df[~df["sous poste"].isin(lignes_a_exclure)]

    df = df.drop_duplicates()

    for col in ["patho_niv1", "patho_niv2", "poste de dépense", "sous poste"]:
        if col in df.columns:
            df[col] = df[col].astype("category")

    logger.info(f"Dépenses nettoyées: {len(df):,} lignes")
    return df


def compute_list_lengths(
    df_depenses: pd.DataFrame, df_effectifs: pd.DataFrame
) -> dict[str, int]:
    len_dict = {}

    len_dict["len_annee"] = df_depenses["annee"].nunique(dropna=True) + 1
    len_dict["len_patho_niv1"] = (
        df_depenses.loc[
            ~df_depenses["patho_niv1"].isin(
                ["Total", "Total consommants tous régimes"]
            ),
            "patho_niv1",
        ].nunique(dropna=True)
        + 1
    )
    len_dict["len_poste de dépense"] = (
        df_depenses["poste de dépense"].nunique(dropna=True) + 1
    )
    len_dict["len_sous poste"] = df_depenses["sous poste"].nunique(dropna=True) + 1

    len_dict["len_departement_eff"] = (
        df_effectifs["Département"].nunique(dropna=True) + 1
    )
    len_dict["len_classes_age"] = (
        df_effectifs.loc[
            ~df_effectifs["Classe d'age"]
            .astype(str)
            .str.lower()
            .isin(["tous âges", "tous ages", "total"]),
            "Classe d'age",
        ].nunique(dropna=True)
        + 1
    )
    len_dict["len_region"] = df_effectifs["Région"].nunique(dropna=True) + 1
    len_dict["len_patho_effectif"] = (
        df_effectifs.loc[
            ~df_effectifs["patho_niv1"]
            .astype(str)
            .str.lower()
            .isin(["total", "total consommants tous régimes"]),
            "patho_niv1",
        ].nunique(dropna=True)
        + 1
    )

    return len_dict
