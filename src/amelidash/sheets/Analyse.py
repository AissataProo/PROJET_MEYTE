"""Onglet Analyses Unifiées - 3 Graphiques avec filtre Année."""

from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from components.filters import add_filter
from config import COULEURS


class OngletAnalysesUnifiees:
    """3 graphiques: Top 10 Effectif, Top 10 Montant, Répartition Sexe."""
    
    def __init__(self, wb, df_depenses, df_effectifs):
        self.wb = wb
        self.df_depenses = df_depenses
        self.df_effectifs = df_effectifs
        self.couleur_principale = COULEURS["principal"]
        self.couleur_accent = COULEURS["accent"]

    def create(self):
        ws = self.wb.create_sheet("Analyses")
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 85

        ws["A1"] = "📊 Analyses Unifiées"
        ws["A1"].font = Font(bold=True, size=16, color="FFFFFF", name="Calibri")
        ws["A1"].fill = PatternFill(start_color=self.couleur_principale, fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells("A1:H1")
        ws.row_dimensions[1].height = 32

        annees = sorted(self.df_depenses["annee"].unique(), reverse=True)
        annee_defaut = int(annees[0]) if annees else 2023

        add_filter(ws, "B2", "Année", [str(int(a)) for a in annees], str(annee_defaut), self.couleur_accent, self.couleur_principale)

        df_dep_filtered = self.df_depenses[self.df_depenses["annee"] == annee_defaut]
        df_eff_filtered = self.df_effectifs[self.df_effectifs["annee"] == annee_defaut]

        top_effectifs = df_eff_filtered.groupby("patho_niv1")["Effectif"].sum().nlargest(10).sort_values()
        
        row = 5
        for idx, (patho, effectif) in enumerate(top_effectifs.items(), 1):
            ws.cell(row + idx, 1).value = patho
            ws.cell(row + idx, 2).value = effectif
            ws.cell(row + idx, 2).number_format = "#,##0"

        chart_eff = BarChart()
        chart_eff.type = "bar"
        chart_eff.style = 11
        chart_eff.title = "Top 10 Pathologies (Effectifs)"
        chart_eff.height = 13
        chart_eff.width = 18
        chart_eff.legend = None

        data_eff = Reference(ws, min_col=2, min_row=5, max_row=5+len(top_effectifs))
        cats_eff = Reference(ws, min_col=1, min_row=6, max_row=5+len(top_effectifs))
        chart_eff.add_data(data_eff, titles_from_data=True)
        chart_eff.set_categories(cats_eff)
        chart_eff.dataLabels = DataLabelList()
        chart_eff.dataLabels.showVal = True
        ws.add_chart(chart_eff, "E5")

        top_montants = df_dep_filtered.groupby("patho_niv1")["montant"].sum().nlargest(10).sort_values()
        
        row = 5
        for idx, (patho, montant) in enumerate(top_montants.items(), 1):
            ws.cell(row + idx, 4).value = patho
            ws.cell(row + idx, 5).value = montant
            ws.cell(row + idx, 5).number_format = "#,##0 €"

        chart_mont = BarChart()
        chart_mont.type = "bar"
        chart_mont.style = 11
        chart_mont.title = "Top 10 Pathologies (Montant)"
        chart_mont.height = 13
        chart_mont.width = 18
        chart_mont.legend = None

        data_mont = Reference(ws, min_col=5, min_row=5, max_row=5+len(top_montants))
        cats_mont = Reference(ws, min_col=4, min_row=6, max_row=5+len(top_montants))
        chart_mont.add_data(data_mont, titles_from_data=True)
        chart_mont.set_categories(cats_mont)
        chart_mont.dataLabels = DataLabelList()
        chart_mont.dataLabels.showVal = True
        ws.add_chart(chart_mont, "M5")

        sexe_data = df_eff_filtered.groupby("Sexe")["Effectif"].sum()
        
        row = 20
        for idx, (sexe, effectif) in enumerate(sexe_data.items(), 1):
            ws.cell(row + idx, 1).value = sexe
            ws.cell(row + idx, 2).value = effectif
            ws.cell(row + idx, 2).number_format = "#,##0"

        chart_pie = PieChart()
        chart_pie.title = "Répartition Sexe"
        chart_pie.height = 12
        chart_pie.width = 16
        chart_pie.style = 10

        data_pie = Reference(ws, min_col=2, min_row=20, max_row=20+len(sexe_data))
        labels_pie = Reference(ws, min_col=1, min_row=21, max_row=20+len(sexe_data))
        chart_pie.add_data(data_pie, titles_from_data=True)
        chart_pie.set_categories(labels_pie)
        chart_pie.dataLabels = DataLabelList()
        chart_pie.dataLabels.showPercent = True
        ws.add_chart(chart_pie, "E20")

        ws.column_dimensions["A"].width = 35
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["D"].width = 35
        ws.column_dimensions["E"].width = 18