from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from config import COULEURS, SHEET_CLEANED_DEPENSES


class OngletGraphiques:
    def __init__(self, wb, df):
        self.wb = wb
        self.df = df

    def create(self):
        ws = self.wb.create_sheet("Graphiques", 2)
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A6"
        ws.sheet_view.zoomScale = 90

        ws_dep = self.wb[SHEET_CLEANED_DEPENSES]

        headers = {}
        for idx, cell in enumerate(ws_dep[1], start=1):
            if cell.value:
                headers[str(cell.value).strip()] = idx

        col_annee = headers.get("annee", 1)
        col_poste = headers.get("poste de dépense", 5)
        col_patho2 = headers.get("patho_niv2", 3)
        col_montant = headers.get("montant", 7)
        col_effectif = headers.get("effectif", 8)

        col_annee_letter = get_column_letter(col_annee)
        col_poste_letter = get_column_letter(col_poste)
        col_patho2_letter = get_column_letter(col_patho2)
        col_montant_letter = get_column_letter(col_montant)
        col_effectif_letter = get_column_letter(col_effectif)

        last_row_dep = min(ws_dep.max_row, 5000)

        annees = sorted(
            {
                int(r[col_annee - 1])
                for r in ws_dep.iter_rows(min_row=2, max_row=last_row_dep, values_only=True)
                if r[col_annee - 1]
            },
            reverse=True,
        )
        postes = sorted(
            {
                str(r[col_poste - 1]).strip()
                for r in ws_dep.iter_rows(min_row=2, max_row=last_row_dep, values_only=True)
                if r[col_poste - 1] and "total" not in str(r[col_poste - 1]).lower()
            }
        )

        annee_defaut = annees[0] if annees else 2023
        poste_defaut = postes[0] if postes else ""

        ws["A1"] = "Tableau de bord des dépenses"
        ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells("A1:H1")
        ws.row_dimensions[1].height = 28

        ws["A3"] = "Année"
        ws["A3"].font = Font(bold=True, color="FFFFFF")
        ws["A3"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
        ws["B3"] = annee_defaut

        ws["D3"] = "Poste"
        ws["D3"].font = Font(bold=True, color="FFFFFF")
        ws["D3"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
        ws["E3"] = poste_defaut

        dv_annee = DataValidation(type="list", formula1=f'"{",".join(str(a) for a in annees)}"', allow_blank=False)
        dv_poste = DataValidation(type="list", formula1=f'"{",".join(postes)}"', allow_blank=False)
        ws.add_data_validation(dv_annee)
        ws.add_data_validation(dv_poste)
        dv_annee.add("B3")
        dv_poste.add("E3")

        ws["A5"] = "Pathologie niv2"
        ws["B5"] = "Montant"
        ws["C5"] = "Effectif"
        ws["D5"] = "Coût/patient"

        for cell in ws["5:5"]:
            if cell.column <= 4:
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")

        patho2_list = sorted(
            {
                str(r[col_patho2 - 1]).strip()
                for r in ws_dep.iter_rows(min_row=2, max_row=last_row_dep, values_only=True)
                if r[col_patho2 - 1]
                and "total" not in str(r[col_patho2 - 1]).lower()
                and str(r[col_poste - 1]).strip() == poste_defaut
            }
        )

        start_row = 6
        for idx, patho in enumerate(patho2_list):
            r = start_row + idx
            shade = "FFF7F7F7" if idx % 2 == 0 else "FFFFFFFF"
            for c in range(1, 5):
                ws.cell(r, c).fill = PatternFill(start_color=shade, fill_type="solid")
                ws.cell(r, c).border = Border(
                    left=Side(style="thin", color="D9D9D9"),
                    right=Side(style="thin", color="D9D9D9"),
                    top=Side(style="thin", color="D9D9D9"),
                    bottom=Side(style="thin", color="D9D9D9"),
                )

            ws.cell(r, 1).value = patho
            ws.cell(r, 2).value = (
                f"=IFERROR(SUMIFS('{ws_dep.title}'!${col_montant_letter}:${col_montant_letter},"
                f"'{ws_dep.title}'!${col_annee_letter}:${col_annee_letter},$B$3,"
                f"'{ws_dep.title}'!${col_poste_letter}:${col_poste_letter},$E$3,"
                f"'{ws_dep.title}'!${col_patho2_letter}:${col_patho2_letter},$A{r}),0)"
            )
            ws.cell(r, 3).value = (
                f"=IFERROR(SUMIFS('{ws_dep.title}'!${col_effectif_letter}:${col_effectif_letter},"
                f"'{ws_dep.title}'!${col_annee_letter}:${col_annee_letter},$B$3,"
                f"'{ws_dep.title}'!${col_poste_letter}:${col_poste_letter},$E$3,"
                f"'{ws_dep.title}'!${col_patho2_letter}:${col_patho2_letter},$A{r}),0)"
            )
            ws.cell(r, 4).value = f"=IFERROR(B{r}/C{r},0)"
            ws.cell(r, 2).number_format = '#,##0 €'
            ws.cell(r, 3).number_format = '#,##0'
            ws.cell(r, 4).number_format = '#,##0.00 €'
            ws.cell(r, 1).alignment = Alignment(horizontal="left")
            ws.cell(r, 2).alignment = Alignment(horizontal="right")
            ws.cell(r, 3).alignment = Alignment(horizontal="right")
            ws.cell(r, 4).alignment = Alignment(horizontal="right")

        end_row = start_row + len(patho2_list) - 1

        chart = BarChart()
        chart.type = "bar"
        chart.style = 10
        chart.title = f"Dépenses par pathologie niveau 2 — {poste_defaut}"
        chart.y_axis.title = "Pathologie"
        chart.x_axis.title = "Montant (€)"
        chart.height = max(12, len(patho2_list) * 0.28)
        chart.width = 20
        chart.legend = None

        data = Reference(ws, min_col=2, min_row=5, max_row=end_row)
        cats = Reference(ws, min_col=1, min_row=6, max_row=end_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.varyColors = True
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.position = "outEnd"

        ws.add_chart(chart, "F5")

        ws.column_dimensions["A"].width = 45
        ws.column_dimensions["B"].width = 16
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 16
        ws.column_dimensions["E"].width = 28