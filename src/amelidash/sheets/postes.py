"""Onglet Postes : tableau de bord dépenses par poste (filtres, tableau, graphique)."""

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from config import COULEURS, SHEET_FILTRES
from components.filters import add_filter


def OngletPostes(wb: Workbook, df: None, len_dict: dict[str, int]) -> None:
    """Construit le tableau de bord dépenses par poste (Postes).

    Assemble dans l'ordre les filtres, le tableau des postes, puis le graphique.

    Args:
        wb: Classeur cible.
        df: DataFrame des dépenses (optionnel, utilisé pour extraire la liste des postes).
        len_dict: Tailles des listes uniques pour les listes de validation.
    """
    ws = wb.create_sheet("Postes")
    ws.sheet_view.showGridLines = False

    _add_filters(ws, len_dict)
    _add_tables(ws, df)
    _add_charts(ws)


def _add_filters(ws: Worksheet, len_dict: dict[str, int]) -> None:
    """Ajoute les filtres déroulants Année et Poste.

    Args:
        ws: Feuille Postes.
        len_dict: Tailles des listes uniques pour borner les plages de validation.
    """
    # Filtre Année en cellule E2
    add_filter(
        ws,
        "Année",
        2025,
        2,
        4,
        2,
        5,
        f"='{SHEET_FILTRES}'!$B$2:$B${len_dict['len_annee']}",
    )

    # Filtre Poste en cellule E3
    add_filter(
        ws,
        "Poste",
        "Tous",
        3,
        4,
        3,
        5,
        f"='{SHEET_FILTRES}'!$A$2:$A${len_dict['len_poste de dépense']}",
    )


def _add_tables(ws: Worksheet, df: None) -> None:
    """Ajoute le titre, les en-têtes et les formules SUMIFS du tableau des postes.

    Args:
        ws: Feuille Postes.
        df: DataFrame des dépenses (pour extraire la liste des postes).
    """
    # Titre principal
    ws["A1"] = "Dépenses par Poste"
    ws["A1"].font = Font(bold=True, size=13, color=COULEURS["blanc"])
    ws["A1"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")

    # Extraire la liste des postes (si df est fourni)
    if df is not None:
        postes = sorted(df["poste de dépense"].dropna().unique())
    else:
        # Optionnel : lire directement de la feuille cleanedData_Depenses
        ws_dep = ws.parent["cleanedData_Depenses"]
        postes = sorted(
            set(
                str(r[4]).strip()
                for r in ws_dep.iter_rows(min_row=2, values_only=True)
                if r[4] and "total" not in str(r[4]).lower()
            )
        )

    # En-têtes de colonnes
    ws["A5"] = "Poste"
    ws["B5"] = "Montant (€)"
    for col in ["A", "B"]:
        cell = ws[f"{col}5"]
        cell.font = Font(bold=True, color=COULEURS["blanc"], size=9)
        cell.fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")

    # Formules SUMIFS dynamiques (liées aux filtres $E$2 et $E$3)
    source_sheet_name = "cleanedData_Depenses"

    # On suppose:
    #   - colonne C = Année
    #   - colonne E = Poste de dépense
    #   - colonne G = Montant
    # Adaptez si tes colonnes sont différentes.

    for idx, poste in enumerate(postes, start=6):
        ws.cell(idx, 1).value = poste
        ws.cell(idx, 2).value = (
            f"=SUMIFS('{source_sheet_name}'!G:G, "
            f"'{source_sheet_name}'!C:C, $E$2, "
            f"'{source_sheet_name}'!E:E, $E$3)"
        )
        ws.cell(idx, 2).number_format = "#,##0"


def _add_charts(ws: Worksheet) -> None:
    """Ajoute le graphique à barres des dépenses par poste.

    Args:
        ws: Feuille Postes.
    """
    chart = BarChart()
    chart.type = "col"
    chart.title = "Dépenses par Poste"
    chart.height = 12
    chart.width = 20

    # On détermine le nombre de lignes à partir des en-têtes + postes
    # Pour simplique, on suppose que les postes vont de A6 à B<last_row>
    # Tu peux ajuster last_row si tu veux une valeur fixe.
    last_row = ws.max_row

    data_ref = Reference(ws, min_col=2, min_row=5, max_row=last_row)
    cat_ref = Reference(ws, min_col=1, min_row=6, max_row=last_row)

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cat_ref)
    ws.add_chart(chart, "G5")

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 16