import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


class OngletDepartement:
    """
    Gère l'onglet 'Département' qui permet l'analyse géographique
    et temporelle des effectifs par pathologie.
    """

    def __init__(self, wb, df, df_dep=None):
        self.wb = wb
        self.df = df.copy()
        self.df_dep = df_dep.copy() if df_dep is not None else None
        self.create()

    def create(self):
        for nom in [
            "Departement",
            "CalcDept",
            "CalcDepTop",
            "CalcDepSexeAge",
            "CalcDepAge",
        ]:
            if nom in self.wb.sheetnames:
                self.wb.remove(self.wb[nom])

        df = self.df.copy()
        df["Effectif"] = pd.to_numeric(df["Effectif"], errors="coerce").fillna(0)
        if "Population de référence" in df.columns:
            df["Population de référence"] = pd.to_numeric(
                df["Population de référence"], errors="coerce"
            ).fillna(0)
        else:
            df["Population de référence"] = 0

        def _vals(col, exclure=()):
            return sorted(
                {
                    str(v).strip()
                    for v in df[col].dropna().unique()
                    if v
                    and "total" not in str(v).lower()
                    and "tous" not in str(v).lower()
                    and not any(e in str(v).lower() for e in exclure)
                }
            )

        pathos = _vals("patho_niv1")
        deps = _vals("Département")
        ages = _vals("Classe d'age")
        sexes = _vals("Sexe", exclure=("ensemble",))
        if not sexes:
            sexes = sorted({str(v).strip() for v in df["Sexe"].dropna().unique() if v})
        annees = sorted(df["annee"].dropna().unique(), reverse=True)

        patho_def = pathos[0] if pathos else ""
        dept_def = deps[0] if deps else ""
        annee_def = int(annees[0]) if len(annees) else None

        sub = df[(df["patho_niv1"] == patho_def) & (df["annee"] == annee_def)]
        dep_def = sub.groupby("Département", observed=True)[
            ["Effectif", "Population de référence"]
        ].sum()
        dep_def["prev"] = dep_def["Effectif"] / dep_def[
            "Population de référence"
        ].replace(0, pd.NA)
        top_dep = dep_def["prev"].dropna().nlargest(15).sort_values().index.tolist()

        ws = self.wb.create_sheet("Departement")
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = "C00000"

        calc = self.wb.create_sheet("CalcDept")
        calc.sheet_state = "hidden"
        calct = self.wb.create_sheet("CalcDepTop")
        calct.sheet_state = "hidden"
        calca = self.wb.create_sheet("CalcDepAge")
        calca.sheet_state = "hidden"

        COLOR_BLEU = "FF4472C4"
        COLOR_ROUGE = "C00000"
        bord = Border(*([Side(style="thin", color="CCCCCC")] * 4))
        fond = PatternFill("solid", fgColor="FCE9E9")

        for row in range(1, 72):
            for col in range(1, 21):
                ws.cell(row, col).fill = fond

        L_PATHO_COL, L_DEP_COL = 28, 29
        for i, p in enumerate(pathos, start=2):
            ws.cell(i, L_PATHO_COL).value = p
        for i, d in enumerate(deps, start=2):
            ws.cell(i, L_DEP_COL).value = d
        ws.column_dimensions[get_column_letter(L_PATHO_COL)].hidden = True
        ws.column_dimensions[get_column_letter(L_DEP_COL)].hidden = True
        LP = get_column_letter(L_PATHO_COL)
        LD = get_column_letter(L_DEP_COL)

        def _filtre(cell_lbl, cell_val, label, valeur):
            ws[cell_lbl] = label
            ws[cell_lbl].font = Font(bold=True, color="FFFFFF")
            ws[cell_lbl].fill = PatternFill("solid", fgColor=COLOR_BLEU)
            ws[cell_lbl].alignment = Alignment(horizontal="center")
            ws[cell_val] = valeur
            ws[cell_val].font = Font(bold=True)
            ws[cell_val].alignment = Alignment(horizontal="center")

        _filtre("A3", "B3", "Pathologie", patho_def)
        _filtre("A4", "B4", "Département", dept_def)
        _filtre("A5", "B5", "Année", annee_def)

        ws["B3"].alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        ws.row_dimensions[3].height = 30

        dv_p = DataValidation(type="list", formula1=f"${LP}$2:${LP}${1 + len(pathos)}")
        dv_d = DataValidation(type="list", formula1=f"${LD}$2:${LD}${1 + len(deps)}")
        dv_y = DataValidation(
            type="list", formula1=f'"{",".join(str(int(a)) for a in annees)}"'
        )
        for dv, target in ((dv_p, "B3"), (dv_d, "B4"), (dv_y, "B5")):
            ws.add_data_validation(dv)
            dv.add(target)

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 50

        annees_asc = sorted(int(a) for a in annees)
        ws.merge_cells("I2:M2")
        ws["I2"] = "Variation annuelle du nombre de patients"
        ws["I2"].font = Font(bold=True, size=11, color="FFFFFF")
        ws["I2"].fill = PatternFill("solid", fgColor=COLOR_ROUGE)
        ws["I2"].alignment = Alignment(horizontal="center", vertical="center")

        entetes_var = ["Année", "Nb patient", "N-1 Nb patient", "Évolution %"]
        for j, txt in enumerate(entetes_var, start=9):
            c = ws.cell(3, j, txt)
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = PatternFill("solid", fgColor=COLOR_BLEU)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = bord

        for k, an in enumerate(annees_asc):
            r = 4 + k
            ws.cell(r, 9, an).border = bord
            ws.cell(r, 9).alignment = Alignment(horizontal="center")
            ws.cell(r, 10).value = (
                f"=SUMIFS(cleanedData_Effectifs!$G:$G,"
                f"cleanedData_Effectifs!$A:$A,I{r},"
                f"cleanedData_Effectifs!$C:$C,Departement!$B$3)"
            )
            ws.cell(r, 10).number_format = "#,##0"
            ws.cell(r, 10).border = bord
            if k == 0:
                ws.cell(r, 11).value = 0
                ws.cell(r, 12).value = 0
            else:
                ws.cell(r, 11).value = f"=J{r-1}"
                ws.cell(r, 12).value = f"=IFERROR((J{r}-K{r})/K{r},0)"
            ws.cell(r, 11).number_format = "#,##0"
            ws.cell(r, 11).border = bord
            ws.cell(r, 12).number_format = "0.00%"
            ws.cell(r, 12).border = bord

        ws.column_dimensions["I"].width = 10
        ws.column_dimensions["J"].width = 14
        ws.column_dimensions["K"].width = 16
        ws.column_dimensions["L"].width = 14

        calc.cell(1, 1, "Sexe")
        for j, age in enumerate(ages, start=2):
            calc.cell(1, j, age)

        for i, sexe in enumerate(sexes, start=2):
            calc.cell(i, 1, sexe)
            for j in range(2, 2 + len(ages)):
                col_age = get_column_letter(j)
                calc.cell(i, j).value = (
                    f"=SUMIFS(cleanedData_Effectifs!$G:$G,"
                    f"cleanedData_Effectifs!$A:$A,Departement!$B$5,"
                    f"cleanedData_Effectifs!$C:$C,Departement!$B$3,"
                    f"cleanedData_Effectifs!$D:$D,Departement!$B$4,"
                    f"cleanedData_Effectifs!$E:$E,{col_age}$1,"
                    f"cleanedData_Effectifs!$F:$F,$A{i})"
                )
                calc.cell(i, j).number_format = "#,##0"

        n_sexe, n_age = len(sexes), len(ages)

        chart1 = BarChart()
        chart1.type = "col"
        chart1.grouping = "clustered"
        chart1.title = "Effectif par Sexe x Classe d'âge"
        chart1.height = 11
        chart1.width = 16
        chart1.style = 12
        chart1.add_data(
            Reference(
                calc, min_col=2, max_col=1 + n_age, min_row=1, max_row=1 + n_sexe
            ),
            titles_from_data=True,
        )
        chart1.set_categories(Reference(calc, min_col=1, min_row=2, max_row=1 + n_sexe))
        chart1.x_axis.title = "Sexe"
        chart1.y_axis.title = "Effectif"
        chart1.legend.position = "t"
        ws.add_chart(chart1, "A8")

        calct.cell(1, 1, "Département")
        calct.cell(1, 2, "Effectif")
        calct.cell(1, 3, "Population")
        calct.cell(1, 4, "Prévalence")

        for i, dep in enumerate(top_dep, start=2):
            calct.cell(i, 1, dep)
            calct.cell(i, 2).value = (
                f"=SUMIFS(cleanedData_Effectifs!$G:$G,"
                f"cleanedData_Effectifs!$A:$A,Departement!$B$5,"
                f"cleanedData_Effectifs!$C:$C,Departement!$B$3,"
                f"cleanedData_Effectifs!$D:$D,A{i})"
            )
            calct.cell(i, 3).value = (
                f"=SUMIFS(cleanedData_Effectifs!$H:$H,"
                f"cleanedData_Effectifs!$A:$A,Departement!$B$5,"
                f"cleanedData_Effectifs!$C:$C,Departement!$B$3,"
                f"cleanedData_Effectifs!$D:$D,A{i})"
            )
            calct.cell(i, 4).value = f"=IF(C{i}=0,0,B{i}/C{i})"
            calct.cell(i, 4).number_format = "0.00%"

        n_dep = len(top_dep)

        chart2 = BarChart()
        chart2.type = "bar"
        chart2.title = "Top 15 — Taux de prévalence par Département"
        chart2.height = 12
        chart2.width = 16
        chart2.style = 10
        chart2.legend = None
        chart2.add_data(
            Reference(calct, min_col=4, min_row=1, max_row=n_dep + 1),
            titles_from_data=True,
        )
        chart2.set_categories(Reference(calct, min_col=1, min_row=2, max_row=n_dep + 1))
        chart2.x_axis.title = "Département"
        chart2.y_axis.title = None
        if chart2.series:
            chart2.series[0].graphicalProperties.solidFill = COLOR_ROUGE
        ws.add_chart(chart2, "I8")

        calca.cell(1, 1, "Classe d'âge")
        calca.cell(1, 2, "Prévalence")
        for i, age in enumerate(ages, start=2):
            calca.cell(i, 1, age)
            calca.cell(i, 2).value = (
                f"=IFERROR("
                f"SUMIFS(cleanedData_Effectifs!$G:$G,"
                f"cleanedData_Effectifs!$A:$A,Departement!$B$5,"
                f"cleanedData_Effectifs!$C:$C,Departement!$B$3,"
                f"cleanedData_Effectifs!$E:$E,A{i})"
                f"/"
                f"SUMIFS(cleanedData_Effectifs!$H:$H,"
                f"cleanedData_Effectifs!$A:$A,Departement!$B$5,"
                f"cleanedData_Effectifs!$C:$C,Departement!$B$3,"
                f"cleanedData_Effectifs!$E:$E,A{i})"
                f",0)"
            )
            calca.cell(i, 2).number_format = "0.00%"

        ws.row_dimensions[2].height = 22
        ws.row_dimensions[3].height = 24
        ws.row_dimensions[5].height = 22
