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
        ws.freeze_panes = "A7"
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

        ws.merge_cells("A1:H1")
        ws["A1"] = "Dépenses national par pathologies"
        ws["A1"].font = Font(bold=True, size=18, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40

        # ✅ KPIs
        ws["A3"] = "Montant Total"
        ws["A3"].font = Font(bold=True, size=11, color="FFFFFF")
        ws["A3"].fill = PatternFill(start_color=COULEURS["accent"], fill_type="solid")

        ws["B3"] = "=IFERROR(SUM(B7:B500),0)"
        ws["B3"].number_format = "#,##0 €"
        ws["B3"].font = Font(bold=True, size=13, color=COULEURS["principal"])
        ws["B3"].fill = PatternFill(start_color="FFFAFAFA", fill_type="solid")

        ws["D3"] = "👥 Effectif Total"
        ws["D3"].font = Font(bold=True, size=11, color="FFFFFF")
        ws["D3"].fill = PatternFill(start_color=COULEURS["accent"], fill_type="solid")

        ws["E3"] = "=IFERROR(SUM(C7:C500),0)"
        ws["E3"].number_format = "#,##0"
        ws["E3"].font = Font(bold=True, size=13, color=COULEURS["principal"])
        ws["E3"].fill = PatternFill(start_color="FFFAFAFA", fill_type="solid")

        ws["A5"] = "Année"
        ws["A5"].font = Font(bold=True, color="FFFFFF")
        ws["A5"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
        ws["B5"] = annee_defaut
        ws["B5"].fill = PatternFill(start_color="FFFFFFFF", fill_type="solid")
        ws["B5"].font = Font(bold=True, size=11)

        ws["D5"] = "Poste"
        ws["D5"].font = Font(bold=True, color="FFFFFF")
        ws["D5"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
        ws["E5"] = poste_defaut
        ws["E5"].fill = PatternFill(start_color="FFFFFFFF", fill_type="solid")
        ws["E5"].font = Font(bold=True, size=11)

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
        dv_annee.add("B5")
        dv_poste.add("E5")

        ws["A7"] = "Pathologie niv2"
        ws["B7"] = "Montant"
        ws["C7"] = "Effectif"
        ws["D7"] = "Coût/patient"

        for cell in ws["7:7"]:
            if cell.column <= 4:
                cell.font = Font(bold=True, color="FFFFFF", size=11)
                cell.fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(
                    left=Side(style="thin", color="CCCCCC"),
                    right=Side(style="thin", color="CCCCCC"),
                    top=Side(style="thin", color="CCCCCC"),
                    bottom=Side(style="thin", color="CCCCCC"),
                )

        patho2_list = sorted(
            {
                str(r[col_patho2 - 1]).strip()
                for r in ws_dep.iter_rows(min_row=2, max_row=last_row, values_only=True)
                if r[col_patho2 - 1]
                and "total" not in str(r[col_patho2 - 1]).lower()
                and str(r[col_poste - 1]).strip() == poste_defaut
            }
        )

        start_row = 8
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
                f"=IFERROR(SUMIFS('{ws_dep.title}'!${c_montant}$2:${c_montant}$5000,"
                f"'{ws_dep.title}'!${c_annee}$2:${c_annee}$5000,$B$5,"
                f"'{ws_dep.title}'!${c_poste}$2:${c_poste}$5000,$E$5,"
                f"'{ws_dep.title}'!${c_patho2}$2:${c_patho2}$5000,$A{r}),0)"
            )
            ws.cell(r, 3).value = (
                f"=IFERROR(SUMIFS('{ws_dep.title}'!${c_effectif}$2:${c_effectif}$5000,"
                f"'{ws_dep.title}'!${c_annee}$2:${c_annee}$5000,$B$5,"
                f"'{ws_dep.title}'!${c_poste}$2:${c_poste}$5000,$E$5,"
                f"'{ws_dep.title}'!${c_patho2}$2:${c_patho2}$5000,$A{r}),0)"
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
            chart.title = f"Dépenses par pathologie — {poste_defaut}"
            chart.y_axis.title = ""
            chart.x_axis.title = "Montant (€)"
            chart.height = max(12, min(len(patho2_list) * 0.28, 24))
            chart.width = 22
            chart.legend = None

            data = Reference(ws, min_col=2, min_row=7, max_row=end_row)
            cats = Reference(ws, min_col=1, min_row=8, max_row=end_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.varyColors = True
            ws.add_chart(chart, "F7")

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
        ws.freeze_panes = "A7"
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

        # ✅ TITRE AMÉLIORÉ
        ws.merge_cells("A1:N1")
        ws["A1"] = "📊 Tableau de Bord des Effectifs"
        ws["A1"].font = Font(bold=True, size=18, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40

        # ✅ KPI
        ws["A3"] = "👥 Effectif Total"
        ws["A3"].font = Font(bold=True, size=11, color="FFFFFF")
        ws["A3"].fill = PatternFill(start_color=COULEURS["accent"], fill_type="solid")

        ws["B3"] = "=IFERROR(SUM(B7:B500),0)"
        ws["B3"].number_format = "#,##0"
        ws["B3"].font = Font(bold=True, size=13, color=COULEURS["principal"])
        ws["B3"].fill = PatternFill(start_color="FFFAFAFA", fill_type="solid")

        # ✅ FILTRES AMÉLIORÉS
        ws["A5"] = "Année"
        ws["A5"].font = Font(bold=True, color="FFFFFF")
        ws["A5"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
        ws["B5"] = annee_defaut
        ws["B5"].fill = PatternFill(start_color="FFFFFFFF", fill_type="solid")
        ws["B5"].font = Font(bold=True, size=11)

        ws["D5"] = "Pathologie"
        ws["D5"].font = Font(bold=True, color="FFFFFF")
        ws["D5"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
        ws["E5"] = patho_defaut
        ws["E5"].fill = PatternFill(start_color="FFFFFFFF", fill_type="solid")
        ws["E5"].font = Font(bold=True, size=11)

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
        dv_annee.add("B5")
        dv_patho.add("E5")

        ws["A7"] = "Département"
        ws["B7"] = "Effectif"
        ws["F7"] = "Région"
        ws["G7"] = "Effectif"

        for cell in ws["7:7"]:
            if cell.column in (1, 2, 6, 7):
                cell.font = Font(bold=True, color="FFFFFF", size=11)
                cell.fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(
                    left=Side(style="thin", color="CCCCCC"),
                    right=Side(style="thin", color="CCCCCC"),
                    top=Side(style="thin", color="CCCCCC"),
                    bottom=Side(style="thin", color="CCCCCC"),
                )

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

        top_n = 15

        for i, dept in enumerate(dept_list[:top_n]):
            r = 8 + i
            ws.cell(r, 1).value = dept
            ws.cell(r, 1).fill = PatternFill(start_color="FFF7F7F7" if i % 2 == 0 else "FFFFFFFF", fill_type="solid")
            ws.cell(r, 2).value = (
                f"=IFERROR(SUMIFS('{ws_eff.title}'!${c_eff}$2:${c_eff}$5000,"
                f"'{ws_eff.title}'!${c_annee}$2:${c_annee}$5000,$B$5,"
                f"'{ws_eff.title}'!${c_patho2}$2:${c_patho2}$5000,$E$5,"
                f"'{ws_eff.title}'!${c_dept}$2:${c_dept}$5000,$A{r}),0)"
            )
            ws.cell(r, 2).number_format = "#,##0"
            ws.cell(r, 2).fill = PatternFill(start_color="FFF7F7F7" if i % 2 == 0 else "FFFFFFFF", fill_type="solid")

        for i, region in enumerate(region_list[:top_n]):
            r = 8 + i
            ws.cell(r, 6).value = region
            ws.cell(r, 6).fill = PatternFill(start_color="FFF7F7F7" if i % 2 == 0 else "FFFFFFFF", fill_type="solid")
            ws.cell(r, 7).value = (
                f"=IFERROR(SUMIFS('{ws_eff.title}'!${c_eff}$2:${c_eff}$5000,"
                f"'{ws_eff.title}'!${c_annee}$2:${c_annee}$5000,$B$5,"
                f"'{ws_eff.title}'!${c_patho2}$2:${c_patho2}$5000,$E$5,"
                f"'{ws_eff.title}'!${c_region}$2:${c_region}$5000,$F{r}),0)"
            )
            ws.cell(r, 7).number_format = "#,##0"
            ws.cell(r, 7).fill = PatternFill(start_color="FFF7F7F7" if i % 2 == 0 else "FFFFFFFF", fill_type="solid")

        if dept_list:
            chart_dept = BarChart()
            chart_dept.type = "bar"
            chart_dept.style = 10
            chart_dept.title = f"Top Départements — {patho_defaut}"
            chart_dept.y_axis.title = ""
            chart_dept.x_axis.title = "Effectif"
            chart_dept.height = 14
            chart_dept.width = 22
            chart_dept.legend = None

            data_dept = Reference(ws, min_col=2, min_row=7, max_row=8 + min(top_n, len(dept_list)) - 1)
            cats_dept = Reference(ws, min_col=1, min_row=8, max_row=8 + min(top_n, len(dept_list)) - 1)
            chart_dept.add_data(data_dept, titles_from_data=True)
            chart_dept.set_categories(cats_dept)
            chart_dept.varyColors = True
            ws.add_chart(chart_dept, "J7")

        if region_list:
            chart_reg = BarChart()
            chart_reg.type = "bar"
            chart_reg.style = 10
            chart_reg.title = f"Top Régions — {patho_defaut}"
            chart_reg.y_axis.title = ""
            chart_reg.x_axis.title = "Effectif"
            chart_reg.height = 14
            chart_reg.width = 22
            chart_reg.legend = None

            data_reg = Reference(ws, min_col=7, min_row=7, max_row=8 + min(top_n, len(region_list)) - 1)
            cats_reg = Reference(ws, min_col=6, min_row=8, max_row=8 + min(top_n, len(region_list)) - 1)
            chart_reg.add_data(data_reg, titles_from_data=True)
            chart_reg.set_categories(cats_reg)
            chart_reg.varyColors = True
            ws.add_chart(chart_reg, "J23")

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["F"].width = 28
        ws.column_dimensions["G"].width = 14