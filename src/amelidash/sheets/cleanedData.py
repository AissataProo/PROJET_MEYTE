import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from config import SHEET_CLEANED_EFFECTIFS, SHEET_CLEANED_DEPENSES


def build_cleaned_data_effectifs(wb: Workbook, df: pd.DataFrame) -> None:
    """Onglet cleaned_data : écrit le DataFrame nettoyé (effectifs ou dépenses) dans
    le classeur, en créant l’onglet dédié et en y insérant toutes les lignes du DataFrame
    via `dataframe_to_rows`, sans index et avec en-têtes."""

    ws = wb.create_sheet(SHEET_CLEANED_EFFECTIFS)
    ws.sheet_view.showGridLines = False
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)


def build_cleaned_data_depenses(wb: Workbook, df: pd.DataFrame) -> None:
    """Écrit le DataFrame nettoyé des dépenses dans l'onglet cleanedData_Depenses.

    Args:
        wb: Classeur cible.
        df: DataFrame nettoyé des dépenses.
    """
    ws = wb.create_sheet(SHEET_CLEANED_DEPENSES)
    ws.sheet_view.showGridLines = False
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
