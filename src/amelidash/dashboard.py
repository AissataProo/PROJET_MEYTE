import pandas as pd
from openpyxl import Workbook
from sheets.cleanedData_Depenses import OngletCleanedDataDepenses
from sheets.cleanedData_Effectifs import OngletCleanedDataEffectifs
from sheets.Analyse import OngletAnalysesUnifiees
from sheets.postes import OngletPostes
from sheets.depenses import OngletDepenses
from sheets.departement import OngletDepartement
from sheets.filtres import OngletFiltres
from sheets.filterseff import OngletFiltreseff


def build_workbook(
    df_dep: pd.DataFrame, df_eff: pd.DataFrame, len_dict: dict[str, int]
) -> Workbook:
    """Assemble le classeur Excel complet en créant successivement :

    - les onglets de données nettoyées (dépenses et effectifs) ;
    - les onglets de filtres associés ;
    - l’onglet d’analyses unifiées (Top 10, graphiques, filtres dynamiques) ;
    - l’onglet Postes (synthèse des dépenses par poste et pathologie) ;
    - l’onglet Dépenses (analyses détaillées) ;
    - l’onglet Département (analyses territoriales).

    Retourne un classeur entièrement structuré et prêt pour le dashboard Excel."""

    wb = Workbook()

    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    OngletCleanedDataDepenses(wb, df_dep).create()
    OngletCleanedDataEffectifs(wb, df_eff).create()
    OngletFiltres(wb, df_dep).create_depenses()
    OngletFiltreseff(wb, df_eff).create_effectifs()
    OngletAnalysesUnifiees(wb, df_dep).create()
    OngletPostes(wb, df_dep).create()
    OngletDepenses(wb, df_dep).create()
    OngletDepartement(wb, df_eff).create()

    return wb
