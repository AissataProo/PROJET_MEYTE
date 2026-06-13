"""Filtres interactifs : listes déroulantes avec validation de données Excel."""

from openpyxl.styles import Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from config import COULEURS


def add_filter(
    sheet: Worksheet,
    cell: str,
    title: str,
    values: list[str | int],
    default_value: str | int,
    color_value: str,
    color_label: str,
) -> None:
    """Ajoute un filtre avec liste déroulante à la feuille."""
    col_letters = "".join(ch for ch in cell if ch.isalpha())
    row_num = int("".join(ch for ch in cell if ch.isdigit()))
    col_num = ord(col_letters.upper()) - ord("A") + 1

    cell_title = sheet.cell(row=row_num, column=col_num)
    cell_title.value = title
    cell_title.alignment = Alignment(horizontal="center", vertical="center")
    cell_title.fill = PatternFill(start_color=color_label, end_color=color_label, fill_type="solid")

    cell_value = sheet.cell(row=row_num, column=col_num + 1)
    cell_value.value = default_value
    cell_value.alignment = Alignment(horizontal="center", vertical="center")
    cell_value.fill = PatternFill(start_color=color_value, end_color=color_value, fill_type="solid")

    formula1 = f'"{",".join(str(v) for v in values)}"'
    dv = DataValidation(type="list", formula1=formula1)
    sheet.add_data_validation(dv)
    dv.add(cell_value.coordinate)


def _add_filter_title(
    sheet: Worksheet,
    title: str,
    start_row: int,
    start_column: int,
    end_row: int,
    end_column: int,
) -> None:
    """Insère et met en forme le libellé d'un filtre avec fusion de cellules."""
    cell = sheet.cell(row=start_row, column=start_column)
    cell.value = title
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = PatternFill(
        start_color=COULEURS["gris_clair"],
        end_color=COULEURS["gris_clair"],
        fill_type="solid",
    )
    sheet.merge_cells(
        start_row=start_row,
        start_column=start_column,
        end_row=end_row,
        end_column=end_column,
    )


def _add_filter_value(
    sheet: Worksheet,
    value: str | int,
    start_row: int,
    start_column: int,
    end_row: int,
    end_column: int,
) -> None:
    """Insère et met en forme la valeur par défaut d'un filtre avec fusion de cellules."""
    cell = sheet.cell(row=start_row, column=start_column)
    cell.value = value
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = PatternFill(
        start_color=COULEURS["gris_clair"],
        end_color=COULEURS["gris_clair"],
        fill_type="solid",
    )
    sheet.merge_cells(
        start_row=start_row,
        start_column=start_column,
        end_row=end_row,
        end_column=end_column,
    )


def _add_data_validation(
    sheet: Worksheet,
    start_row: int,
    start_column: int,
    formula: str,
) -> None:
    """Ajoute une validation de données (liste déroulante) à une cellule."""
    dv = DataValidation(type="list", formula1=formula)
    sheet.add_data_validation(dv)
    dv.add(f"{get_column_letter(start_column)}{start_row}")