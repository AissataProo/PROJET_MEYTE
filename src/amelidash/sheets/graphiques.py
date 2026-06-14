from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from config import COULEURS, SHEET_CLEANED_DEPENSES, SHEET_CLEANED_EFFECTIFS


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
        headers = {
            str(cell.value).strip(): idx
            for idx, cell in enumerate(ws_dep[1], start=1)
            if cell.value
        }

        col_annee = headers.get("annee", 1)
        col_patho2 = headers.get("patho_niv2", 3)
        col_poste = headers.get("poste de dépense", 5)
        col_montant = headers.get("montant", 7)
        col_effectif = headers.get("Effectif", 8)

        c_annee = get_column_letter(col_annee)
        c_patho2 = get_column_letter(col_patho2)
        c_poste = get_column_letter(col_poste)
        c_montant = get_column_letter(col_montant)
        c_effectif = get_column_letter(col_effectif)

        last_row = min(ws_dep.max_row, 5000)

        annees = sorted(
            {
                int(r[col_annee - 1])
                for r in ws_dep.iter_rows(min_row=2, max_row=last_row, values_only=True)
                if r[col_annee - 1]
            },
            reverse=True,
        )
        postes = sorted(
            {
                str(r[col_poste - 1]).strip()
                for r in ws_dep.iter_rows(min_row=2, max_row=last_row, values_only=True)
                if r[col_poste - 1] and "total" not in str(r[col_poste - 1]).lower()
            }
        )

        annee_defaut = annees[0] if annees else ""
        poste_defaut = postes[0] if postes else ""

        ws["A1"] = "Tableau de bord des dépenses"
        ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
        ws["A1"].fill = PatternFill(
            start_color=COULEURS["principal"], fill_type="solid"
        )
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells("A1:H1")
        ws.row_dimensions[1].height = 28

        ws["A3"] = "Année"
        ws["A3"].font = Font(bold=True, color="FFFFFF")
        ws["A3"].fill = PatternFill(
            start_color=COULEURS["principal"], fill_type="solid"
        )
        ws["B3"] = annee_defaut

        ws["D3"] = "Poste"
        ws["D3"].font = Font(bold=True, color="FFFFFF")
        ws["D3"].fill = PatternFill(
            start_color=COULEURS["principal"], fill_type="solid"
        )
        ws["E3"] = poste_defaut

        dv_annee = DataValidation(
            type="list",
            formula1=f'"{",".join(str(a) for a in annees)}"',
            allow_blank=False,
        )
        dv_poste = DataValidation(
            type="list", formula1=f'"{",".join(postes)}"', allow_blank=False
        )
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
                cell.fill = PatternFill(
                    start_color=COULEURS["principal"], fill_type="solid"
                )
                cell.alignment = Alignment(horizontal="center", vertical="center")

        patho2_list = sorted(
            {
                str(r[col_patho2 - 1]).strip()
                for r in ws_dep.iter_rows(min_row=2, max_row=last_row, values_only=True)
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
                f"=IFERROR(SUMIFS('{ws_dep.title}'!${c_montant}:${c_montant},"
                f"'{ws_dep.title}'!${c_annee}:${c_annee},$B$3,"
                f"'{ws_dep.title}'!${c_poste}:${c_poste},$E$3,"
                f"'{ws_dep.title}'!${c_patho2}:${c_patho2},$A{r}),0)"
            )
            ws.cell(r, 3).value = (
                f"=IFERROR(SUMIFS('{ws_dep.title}'!${c_effectif}:${c_effectif},"
                f"'{ws_dep.title}'!${c_annee}:${c_annee},$B$3,"
                f"'{ws_dep.title}'!${c_poste}:${c_poste},$E$3,"
                f"'{ws_dep.title}'!${c_patho2}:${c_patho2},$A{r}),0)"
            )
            ws.cell(r, 4).value = f"=IFERROR(B{r}/C{r},0)"
            ws.cell(r, 2).number_format = "#,##0 €"
            ws.cell(r, 3).number_format = "#,##0"
            ws.cell(r, 4).number_format = "#,##0.00 €"

        end_row = start_row + len(patho2_list) - 1 if patho2_list else start_row

        if patho2_list:
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


class OngletGraphiquesEffectifs:
    def __init__(self, wb, df):
        self.wb = wb
        self.df = df

    def create(self):
        ws = self.wb.create_sheet("Graphiques Effectifs", 3)
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A6"
        ws.sheet_view.zoomScale = 90

        ws_eff = self.wb[SHEET_CLEANED_EFFECTIFS]
        headers = {
            str(cell.value).strip(): idx
            for idx, cell in enumerate(ws_eff[1], start=1)
            if cell.value
        }

        col_annee = headers.get("annee", 1)
        col_patho2 = headers.get("patho_niv2", 3)
        col_dept = headers.get("Département", 5)
        col_region = headers.get("Région", 6)
        col_eff = headers.get("Effectif", 9)

        c_annee = get_column_letter(col_annee)
        c_patho2 = get_column_letter(col_patho2)
        c_dept = get_column_letter(col_dept)
        c_region = get_column_letter(col_region)
        c_eff = get_column_letter(col_eff)

        last_row = min(ws_eff.max_row, 5000)

        annees = sorted(
            {
                int(r[col_annee - 1])
                for r in ws_eff.iter_rows(min_row=2, max_row=last_row, values_only=True)
                if r[col_annee - 1]
            },
            reverse=True,
        )
        pathos = sorted(
            {
                str(r[col_patho2 - 1]).strip()
                for r in ws_eff.iter_rows(min_row=2, max_row=last_row, values_only=True)
                if r[col_patho2 - 1] and "total" not in str(r[col_patho2 - 1]).lower()
            }
        )

        annee_defaut = annees[0] if annees else ""
        patho_defaut = pathos[0] if pathos else ""

        ws["A1"] = "Tableau de bord des effectifs"
        ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
        ws["A1"].fill = PatternFill(
            start_color=COULEURS["principal"], fill_type="solid"
        )
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells("A1:N1")
        ws.row_dimensions[1].height = 28

        ws["A3"] = "Année"
        ws["A3"].font = Font(bold=True, color="FFFFFF")
        ws["A3"].fill = PatternFill(
            start_color=COULEURS["principal"], fill_type="solid"
        )
        ws["B3"] = annee_defaut

        ws["D3"] = "Pathologie"
        ws["D3"].font = Font(bold=True, color="FFFFFF")
        ws["D3"].fill = PatternFill(
            start_color=COULEURS["principal"], fill_type="solid"
        )
        ws["E3"] = patho_defaut

        dv_annee = DataValidation(
            type="list",
            formula1=f'"{",".join(str(a) for a in annees)}"',
            allow_blank=False,
        )
        dv_patho = DataValidation(
            type="list", formula1=f'"{",".join(pathos)}"', allow_blank=False
        )
        ws.add_data_validation(dv_annee)
        ws.add_data_validation(dv_patho)
        dv_annee.add("B3")
        dv_patho.add("E3")

        ws["A5"] = "Département"
        ws["B5"] = "Effectif"
        ws["F5"] = "Région"
        ws["G5"] = "Effectif"

        for cell in ws["5:5"]:
            if cell.column in (1, 2, 6, 7):
                cell.font = Font(bold=True, color="FFFFFF", size=10)
                cell.fill = PatternFill(
                    start_color=COULEURS["principal"], fill_type="solid"
                )
                cell.alignment = Alignment(horizontal="center", vertical="center")

        dept_list = sorted(
            {
                str(r[col_dept - 1]).strip()
                for r in ws_eff.iter_rows(min_row=2, max_row=last_row, values_only=True)
                if r[col_dept - 1] and "total" not in str(r[col_dept - 1]).lower()
            }
        )
        region_list = sorted(
            {
                str(r[col_region - 1]).strip()
                for r in ws_eff.iter_rows(min_row=2, max_row=last_row, values_only=True)
                if r[col_region - 1] and "total" not in str(r[col_region - 1]).lower()
            }
        )

        top_n = 30

        for i, dept in enumerate(dept_list[:top_n]):
            r = 6 + i
            ws.cell(r, 1).value = dept
            ws.cell(r, 2).value = (
                f"=IFERROR(SUMIFS('{ws_eff.title}'!${c_eff}:${c_eff},"
                f"'{ws_eff.title}'!${c_annee}:${c_annee},$B$3,"
                f"'{ws_eff.title}'!${c_patho2}:${c_patho2},$E$3,"
                f"'{ws_eff.title}'!${c_dept}:${c_dept},$A{r}),0)"
            )
            ws.cell(r, 2).number_format = "#,##0"

        for i, region in enumerate(region_list[:top_n]):
            r = 6 + i
            ws.cell(r, 6).value = region
            ws.cell(r, 7).value = (
                f"=IFERROR(SUMIFS('{ws_eff.title}'!${c_eff}:${c_eff},"
                f"'{ws_eff.title}'!${c_annee}:${c_annee},$B$3,"
                f"'{ws_eff.title}'!${c_patho2}:${c_patho2},$E$3,"
                f"'{ws_eff.title}'!${c_region}:${c_region},$F{r}),0)"
            )
            ws.cell(r, 7).number_format = "#,##0"

        if dept_list:
            chart_dept = BarChart()
            chart_dept.type = "bar"
            chart_dept.style = 10
            chart_dept.title = f"Top départements — {patho_defaut}"
            chart_dept.y_axis.title = "Département"
            chart_dept.x_axis.title = "Effectif"
            chart_dept.height = 12
            chart_dept.width = 20
            chart_dept.legend = None

            data_dept = Reference(
                ws, min_col=2, min_row=5, max_row=5 + min(top_n, len(dept_list))
            )
            cats_dept = Reference(
                ws, min_col=1, min_row=6, max_row=5 + min(top_n, len(dept_list))
            )
            chart_dept.add_data(data_dept, titles_from_data=True)
            chart_dept.set_categories(cats_dept)
            chart_dept.varyColors = True
            chart_dept.dataLabels = DataLabelList()
            chart_dept.dataLabels.showVal = True
            chart_dept.dataLabels.position = "outEnd"
            ws.add_chart(chart_dept, "J5")

        if region_list:
            chart_reg = BarChart()
            chart_reg.type = "bar"
            chart_reg.style = 10
            chart_reg.title = f"Top régions — {patho_defaut}"
            chart_reg.y_axis.title = "Région"
            chart_reg.x_axis.title = "Effectif"
            chart_reg.height = 12
            chart_reg.width = 20
            chart_reg.legend = None

            data_reg = Reference(
                ws, min_col=7, min_row=5, max_row=5 + min(top_n, len(region_list))
            )
            cats_reg = Reference(
                ws, min_col=6, min_row=6, max_row=5 + min(top_n, len(region_list))
            )
            chart_reg.add_data(data_reg, titles_from_data=True)
            chart_reg.set_categories(cats_reg)
            chart_reg.varyColors = True
            chart_reg.dataLabels = DataLabelList()
            chart_reg.dataLabels.showVal = True
            chart_reg.dataLabels.position = "outEnd"
            ws.add_chart(chart_reg, "J25")

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["F"].width = 28
        ws.column_dimensions["G"].width = 14
