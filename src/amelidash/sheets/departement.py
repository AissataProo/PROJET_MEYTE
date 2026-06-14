from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from config import COULEURS, SHEET_CLEANED_EFFECTIFS
from components.filters import add_filter


class OngletDepartement:
    def __init__(self, wb, df):
        self.wb = wb
        self.df = df
        self.sheet_cleaned = SHEET_CLEANED_EFFECTIFS
        self.color_label = COULEURS["accent"]
        self.color_value = COULEURS["secondaire"]

    def create(self):
        if "Departement" in self.wb.sheetnames:
            del self.wb["Departement"]
        ws = self.wb.create_sheet("Departement")
        ws.sheet_view.showGridLines = False

        L_DEPT = "A"
        L_PATHO = "B"
        L_ANNEE = "C"
        L_CLASSE_AGE = "D"
        L_SEXE = "E"
        L_EFFECTIF = "F"

        depts = sorted(self.df["Département"].dropna().unique())
        annees = sorted(
            [str(int(a)) for a in self.df["annee"].dropna().unique()], reverse=True
        )
        ages = sorted(self.df["Classe d'age"].dropna().astype(str).str.strip().unique())
        pathos = sorted(
            [
                p
                for p in self.df["patho_niv1"].dropna().unique()
                if "total" not in str(p).lower()
            ]
        )

        ws.merge_cells("A1:F1")
        ws["A1"] = "Analyse Pathologie / Sexe par Département"
        ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor="FF004A99")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        add_filter(
            ws, "E2", "Département", depts, depts[0], self.color_value, self.color_label
        )
        add_filter(
            ws, "E3", "Année", annees, annees[0], self.color_value, self.color_label
        )
        add_filter(ws, "E4", "Âge", ages, ages[0], self.color_value, self.color_label)

        headers = ["Pathologie", "Effectif Homme", "Effectif Femme"]
        for col_idx, title in enumerate(headers, 1):
            cell = ws.cell(row=6, column=col_idx, value=title)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="FF4472C4")
            cell.alignment = Alignment(horizontal="center")
            ws.column_dimensions[get_column_letter(col_idx)].width = 28

        VAL_DEPT = "$F$2"
        VAL_ANNEE = "$F$3"
        VAL_AGE = "$F$4"

        for idx, patho in enumerate(pathos, start=7):
            ws.cell(idx, 1, patho)

            ws.cell(idx, 2).value = (
                f"=SUMIFS('{self.sheet_cleaned}'!${L_EFFECTIF}:${L_EFFECTIF},"
                f"'{self.sheet_cleaned}'!${L_DEPT}:${L_DEPT},{VAL_DEPT},"
                f"'{self.sheet_cleaned}'!${L_ANNEE}:${L_ANNEE},{VAL_ANNEE},"
                f"'{self.sheet_cleaned}'!${L_CLASSE_AGE}:${L_CLASSE_AGE},{VAL_AGE},"
                f"'{self.sheet_cleaned}'!${L_PATHO}:${L_PATHO},$A{idx},"
                f"'{self.sheet_cleaned}'!${L_SEXE}:${L_SEXE},\"H\")"
            )

            ws.cell(idx, 3).value = (
                f"=SUMIFS('{self.sheet_cleaned}'!${L_EFFECTIF}:${L_EFFECTIF},"
                f"'{self.sheet_cleaned}'!${L_DEPT}:${L_DEPT},{VAL_DEPT},"
                f"'{self.sheet_cleaned}'!${L_ANNEE}:${L_ANNEE},{VAL_ANNEE},"
                f"'{self.sheet_cleaned}'!${L_CLASSE_AGE}:${L_CLASSE_AGE},{VAL_AGE},"
                f"'{self.sheet_cleaned}'!${L_PATHO}:${L_PATHO},$A{idx},"
                f"'{self.sheet_cleaned}'!${L_SEXE}:${L_SEXE},\"F\")"
            )

        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Répartition par Sexe et Pathologie"
        chart.y_axis.title = "Effectif"
        chart.x_axis.title = "Pathologie"
        chart.height = 12
        chart.width = 16

        max_row = len(pathos) + 6
        data = Reference(ws, min_col=2, min_row=6, max_col=3, max_row=max_row)
        cats = Reference(ws, min_col=1, min_row=7, max_row=max_row)

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        ws.add_chart(chart, "G6")
