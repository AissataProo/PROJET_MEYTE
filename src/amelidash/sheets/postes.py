"""Onglet Postes : Tableaux de synthèse + détails par patho_niv2 + camembert."""

from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from config import SHEET_CLEANED_DEPENSES


class OngletPostes:
    def __init__(self, wb, df):
        self.wb = wb
        self.df = df
        self.sheet_cleaned_dep = SHEET_CLEANED_DEPENSES
        self.create()

    def create(self):
        ws = self.wb.create_sheet("Postes")
        ws.sheet_view.showGridLines = False

        ws_dep = self.wb[self.sheet_cleaned_dep]
        sheet_dep = self.sheet_cleaned_dep

        headers = {
            str(cell.value).strip(): idx
            for idx, cell in enumerate(ws_dep[1], start=1)
            if cell.value
        }
        L_annee = get_column_letter(headers.get("annee", 1))
        L_patho2 = get_column_letter(headers.get("patho_niv2", 3))
        L_poste = get_column_letter(headers.get("poste de dépense", 4))
        L_montant = get_column_letter(headers.get("montant", 6))

        annees_list = sorted(self.df["annee"].dropna().unique(), reverse=True)
        annee_def = int(annees_list[0]) if annees_list else 2022

        category_mapping = {
            "Soins de ville": "Soins de ville",
            "Hospitalisations": "Hospitalisations (tous secteurs)",
            "Prestations en espèces": "Prestations en espèces",
        }

        pathos_list = sorted(
            {str(v).strip() for v in self.df["patho_niv2"].dropna().unique() if v}
        )

        COLOR_ROUGE = "C00000"
        COLOR_BLEU = "4472C4"
        bord_d = Border(*([Side(style="thin", color="CCCCCC")] * 4))

        # --- Filtre Année ---
        ws["A2"] = "Année :"
        ws["A2"].font = Font(bold=True)
        ws["B2"] = annee_def
        ws["B2"].font = Font(bold=True)
        dv_a = DataValidation(
            type="list", formula1=f'"{",".join(str(int(a)) for a in annees_list)}"'
        )
        ws.add_data_validation(dv_a)
        dv_a.add("B2")

        # Total des dépenses de l'année sélectionnée
        ws["A4"] = "Total des dépenses"
        ws["B4"] = (
            f"=SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
            f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2)"
        )
        ws["B4"].number_format = "#,##0 €"
        ws["A4"].font = Font(bold=True)
        ws["B4"].font = Font(bold=True)

        # --- Tableau 1 : Détail par patho_niv2 en haut ---
        start_row = 5
        ws[f"A{start_row}"] = "Détails des dépenses par Pathologie (Niv 2)"
        ws.merge_cells(f"A{start_row}:C{start_row}")
        ws.cell(start_row, 1).font = Font(bold=True, color="FFFFFF")
        ws.cell(start_row, 1).fill = PatternFill("solid", fgColor=COLOR_ROUGE)

        headers_patho = ["Pathologie", "Montant (€)", "%"]
        for i, h in enumerate(headers_patho, 1):
            ws.cell(start_row + 1, i, h).fill = PatternFill("solid", fgColor=COLOR_BLEU)
            ws.cell(start_row + 1, i).font = Font(color="FFFFFF")
            ws.cell(start_row + 1, i).alignment = Alignment(horizontal="center")

        ws["D1"] = (
            f"=SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
            f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2)"
        )

        for idx, patho in enumerate(pathos_list):
            r = start_row + 2 + idx
            ws.cell(r, 1, patho).border = bord_d
            ws.cell(r, 2).value = (
                f"=SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
                f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2,"
                f"'{sheet_dep}'!${L_patho2}:${L_patho2},A{r})"
            )
            ws.cell(r, 2).number_format = "#,##0 €"
            ws.cell(r, 3).value = f"=IF($D$1=0,0,B{r}/$D$1)"
            ws.cell(r, 3).number_format = "0.0%"

        # Largeurs du tableau pathologie
        ws.column_dimensions["A"].width = 52
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 10

        # --- Mini tableau postes caché ---
        poste_title_row = 5
        poste_col = 8  # H

        ws.cell(poste_title_row, poste_col, "Répartition par Poste")
        ws.merge_cells(
            start_row=poste_title_row,
            start_column=poste_col,
            end_row=poste_title_row,
            end_column=poste_col + 1
        )
        ws.cell(poste_title_row, poste_col).font = Font(bold=True, color="FFFFFF")
        ws.cell(poste_title_row, poste_col).fill = PatternFill("solid", fgColor=COLOR_BLEU)

        ws.cell(poste_title_row + 1, poste_col, "Poste")
        ws.cell(poste_title_row + 1, poste_col + 1, "Montant")

        for i, (label, source_val) in enumerate(category_mapping.items()):
            row = poste_title_row + 2 + i
            ws.cell(row, poste_col, label)
            ws.cell(row, poste_col + 1).value = (
                f"=SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
                f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2,"
                f"'{sheet_dep}'!${L_poste}:${L_poste},\"{source_val}\")"
            )
            ws.cell(row, poste_col + 1).number_format = "#,##0 €"

        # Cacher le tableau des postes
        ws.column_dimensions["H"].hidden = True
        ws.column_dimensions["I"].hidden = True

        # --- Camembert à droite ---
        pie = PieChart()
        pie.title = "Répartition des dépenses"

        data = Reference(
            ws,
            min_col=poste_col + 1,
            min_row=poste_title_row + 2,
            max_row=poste_title_row + 4
        )
        cats = Reference(
            ws,
            min_col=poste_col,
            min_row=poste_title_row + 2,
            max_row=poste_title_row + 4
        )

        pie.add_data(data, titles_from_data=False)
        pie.set_categories(cats)
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.dataLabels.showVal = False
        pie.legend = None

        ws.add_chart(pie, "J5")

        # Largeur de la zone du graphique
        ws.column_dimensions["J"].width = 3

        # Hauteurs utiles
        ws.row_dimensions[2].height = 22
        ws.row_dimensions[start_row].height = 22