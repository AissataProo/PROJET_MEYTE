"""Onglet Departement : tableau effectifs par département avec graphique et filtre."""

from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill

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
        ws = self.wb.create_sheet("Departement")
        ws.sheet_view.showGridLines = False

        # 1. Titre du tableau
        ws["A1"] = "Effectifs par Département (Top 30)"
        ws["A1"].font = Font(bold=True, size=13, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color="FF004A99", fill_type="solid")

        # 2. Calcul des départements (Logique DataFrame)
        depts_sorted = sorted(
            self.df["Département"].dropna().unique(),
            key=lambda x: self.df[self.df["Département"] == x]["Effectif"].sum(),
            reverse=True,
        )[:30]

        # 3. Ajout du filtre interactif en E2
        add_filter(
            ws,
            "E2",
            "Département",
            list(depts_sorted),
            depts_sorted[0] if depts_sorted else "",
            self.color_value,
            self.color_label,
        )

        # 4. Remplissage du tableau avec SUMIFS dynamique lié au filtre $E$2
        for idx, dept in enumerate(depts_sorted, start=2):
            ws.cell(idx, 1).value = dept
            ws.cell(idx, 2).value = (
                f"=SUMIFS('{self.sheet_cleaned}'!E:E, "
                f"'{self.sheet_cleaned}'!B:B, A{idx}, "
                f"'{self.sheet_cleaned}'!D:D, $E$3)"
            )
            ws.cell(idx, 2).number_format = "#,##0"

        # 5. Création du graphique
        chart = BarChart()
        chart.type = "bar"
        chart.title = "Top 30 Départements"
        chart.height = 20
        chart.width = 18

        data_ref = Reference(ws, min_col=2, min_row=1, max_row=len(depts_sorted) + 1)
        cat_ref = Reference(ws, min_col=1, min_row=2, max_row=len(depts_sorted) + 1)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cat_ref)
        ws.add_chart(chart, "G2")

        # 6. Mise en forme
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 16