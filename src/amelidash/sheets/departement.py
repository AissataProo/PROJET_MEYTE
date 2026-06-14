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

        # 1. Titre
        ws["A1"] = "Effectif par Département (Sélectionné)"
        ws["A1"].font = Font(bold=True, size=13, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color="FF004A99", fill_type="solid")

        # 2. Liste des départements
        depts_sorted = sorted(self.df["Département"].dropna().unique())

        # 3. Filtre en E2 (Le moteur de sélection)
        add_filter(
            ws,
            "E2",
            "Département",
            list(depts_sorted),
            depts_sorted[0],
            self.color_value,
            self.color_label,
        )

        # 4. Remplissage avec condition SI pour le "Focus"
        # Le graphique ne montrera qu'une barre au lieu de 30 vides
        for idx, dept in enumerate(depts_sorted, start=2):
            ws.cell(idx, 1).value = dept
            # Formule : Si le département est celui sélectionné, on calcule, sinon 0
            ws.cell(idx, 2).value = (
                f"=IF(A{idx}=$E$2, SUMIFS('{self.sheet_cleaned}'!E:E, "
                f"'{self.sheet_cleaned}'!B:B, A{idx}), 0)"
            )
            ws.cell(idx, 2).number_format = "#,##0"

        # 5. Graphique dynamique
        chart = BarChart()
        chart.type = "col"  # Barres verticales plus standard pour un focus
        chart.title = "Effectif du département sélectionné"
        chart.height = 10
        chart.width = 12

        # Référence aux données calculées (si 0, la barre disparaît du graphique)
        data_ref = Reference(ws, min_col=2, min_row=2, max_row=len(depts_sorted) + 1)
        cat_ref = Reference(ws, min_col=1, min_row=2, max_row=len(depts_sorted) + 1)

        chart.add_data(data_ref, titles_from_data=False)
        chart.set_categories(cat_ref)
        ws.add_chart(chart, "G2")

        # 6. Mise en forme
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 15
