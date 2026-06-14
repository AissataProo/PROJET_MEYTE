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

        # --- Total des dépenses ---
        ws["A4"] = "Total des dépenses"
        ws["A4"].font = Font(bold=True)
        ws["B4"] = (
            f"=SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
            f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2)"
        )
        ws["B4"].font = Font(bold=True)
        ws["B4"].number_format = "#,##0 €"

        # --- Détail par patho_niv2 ---
        start_row = 5
        ws[f"A{start_row}"] = "Détails des dépenses par Pathologie (Niv 2)"
        ws.merge_cells(f"A{start_row}:C{start_row}")
        ws.cell(start_row, 1).font = Font(bold=True, color="FFFFFF")
        ws.cell(start_row, 1).fill = PatternFill("solid", fgColor=COLOR_ROUGE)
        ws.cell(start_row, 1).alignment = Alignment(horizontal="left")

        headers_patho = ["Pathologie", "Montant (€)", "%"]
        for i, h in enumerate(headers_patho, 1):
            ws.cell(start_row + 1, i, h).fill = PatternFill("solid", fgColor=COLOR_BLEU)
            ws.cell(start_row + 1, i).font = Font(color="FFFFFF")
            ws.cell(start_row + 1, i).alignment = Alignment(horizontal="center")

        for idx, patho in enumerate(pathos_list):
            r = start_row + 2 + idx
            ws.cell(r, 1, patho).border = bord_d
            ws.cell(r, 2).value = (
                f"=SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
                f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2,"
                f"'{sheet_dep}'!${L_patho2}:${L_patho2},A{r})"
            )
            ws.cell(r, 2).number_format = "#,##0 €"
            ws.cell(r, 3).value = f"=SI($B$4=0;0;B{r}/$B$4)"
            ws.cell(r, 3).number_format = "0.0%"

        # --- Répartition des postes ---
        poste_table_row = max(12, start_row + len(pathos_list) + 4)
        ws[f"A{poste_table_row}"] = "Répartition des postes"
        ws.merge_cells(f"A{poste_table_row}:C{poste_table_row}")
        ws.cell(poste_table_row, 1).font = Font(bold=True, color="FFFFFF")
        ws.cell(poste_table_row, 1).fill = PatternFill("solid", fgColor=COLOR_ROUGE)
        ws.cell(poste_table_row, 1).alignment = Alignment(horizontal="left")

        ws.cell(poste_table_row + 1, 1, "Poste")
        ws.cell(poste_table_row + 1, 2, "Montant (€)")
        ws.cell(poste_table_row + 1, 3, "%")
        for c in range(1, 4):
            ws.cell(poste_table_row + 1, c).fill = PatternFill(
                "solid", fgColor=COLOR_BLEU
            )
            ws.cell(poste_table_row + 1, c).font = Font(color="FFFFFF")
            ws.cell(poste_table_row + 1, c).alignment = Alignment(horizontal="center")

        for i, (label, source_val) in enumerate(category_mapping.items()):
            r = poste_table_row + 2 + i
            ws.cell(r, 1, label)
            ws.cell(r, 2).value = (
                f"=SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
                f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2,"
                f"'{sheet_dep}'!${L_poste}:${L_poste},\"{source_val}\")"
            )
            ws.cell(r, 2).number_format = "#,##0 €"
            ws.cell(r, 3).value = f"=SI($B$4=0;0;B{r}/$B$4)"
            ws.cell(r, 3).number_format = "0.0%"

        # --- Camembert configuré E15 ---
        pie = PieChart()
        pie.title = "Répartition des dépenses"
        pie.height = 12.5
        pie.width = 16.0

        data = Reference(
            ws, min_col=2, min_row=poste_table_row + 1, max_row=poste_table_row + 4
        )
        cats = Reference(
            ws, min_col=1, min_row=poste_table_row + 2, max_row=poste_table_row + 4
        )

        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.dataLabels.showVal = False
        pie.legend.position = "b"

        ws.add_chart(pie, "E15")

        # --- Largeurs ---
        ws.column_dimensions["A"].width = 52
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 12
