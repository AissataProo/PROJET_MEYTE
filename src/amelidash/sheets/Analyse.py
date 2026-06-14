"""Onglet Analyses Unifiées - 2 Graphiques côte à côte + Filtre Année."""

from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from components.filters import add_filter
from config import COULEURS


class OngletAnalysesUnifiees:
    """2 Graphiques côte à côte + Filtre Année."""
    
    def __init__(self, wb, df_depenses, df_effectifs):
        self.wb = wb
        self.df_depenses = df_depenses
        self.df_effectifs = df_effectifs
        self.couleur_principale = COULEURS["principal"]
        self.couleur_accent = COULEURS["accent"]

    def create(self):
        ws = self.wb.create_sheet("Analyses")
        ws_data = self.wb.create_sheet("Data_Analyses")
        ws_data.sheet_state = "hidden"

        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 85

        ws.merge_cells("A1:Z1")
        ws["A1"] = "📊 Analyses Unifiées"
        ws["A1"].font = Font(bold=True, size=16, color="FFFFFF", name="Calibri")
        ws["A1"].fill = PatternFill(start_color=self.couleur_principale, fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 35

        annees = sorted(self.df_depenses["annee"].unique(), reverse=True)
        annee_defaut = int(annees[0]) if annees else 2023

        add_filter(ws, "B3", "Année", [str(int(a)) for a in annees], str(annee_defaut), self.couleur_accent, self.couleur_principale)

        df_dep_filtered = self.df_depenses[self.df_depenses["annee"] == annee_defaut]
        df_eff_filtered = self.df_effectifs[self.df_effectifs["annee"] == annee_defaut]

        top_effectifs = df_eff_filtered[
            (~df_eff_filtered["patho_niv1"].str.contains("total", case=False, na=False)) &
            (df_eff_filtered["patho_niv1"].notna())
        ].groupby("patho_niv1")["Effectif"].sum().nlargest(10).sort_values()
        
        row = 1
        for idx, (patho, effectif) in enumerate(top_effectifs.items(), 1):
            ws_data.cell(row + idx, 1).value = patho
            ws_data.cell(row + idx, 2).value = effectif

        chart_eff = BarChart()
        chart_eff.type = "bar"
        chart_eff.style = 11
        chart_eff.title = "Top 10 Pathologies (Effectifs)"
        chart_eff.height = 16
        chart_eff.width = 20

        data_eff = Reference(ws_data, min_col=2, min_row=1, max_row=len(top_effectifs)+1)
        cats_eff = Reference(ws_data, min_col=1, min_row=2, max_row=len(top_effectifs)+1)
        chart_eff.add_data(data_eff, titles_from_data=True)
        chart_eff.set_categories(cats_eff)
        ws.add_chart(chart_eff, "A6")

        top_montants = df_dep_filtered[
            (~df_dep_filtered["patho_niv1"].str.contains("total", case=False, na=False)) &
            (df_dep_filtered["patho_niv1"].notna())
        ].groupby("patho_niv1")["montant"].sum().nlargest(10).sort_values()
        
        row = 1
        for idx, (patho, montant) in enumerate(top_montants.items(), 1):
            ws_data.cell(row + idx, 4).value = patho
            ws_data.cell(row + idx, 5).value = montant

        chart_mont = BarChart()
        chart_mont.type = "bar"
        chart_mont.style = 11
        chart_mont.title = "Top 10 Pathologies (Montant)"
        chart_mont.height = 16
        chart_mont.width = 20

        data_mont = Reference(ws_data, min_col=5, min_row=1, max_row=len(top_montants)+1)
        cats_mont = Reference(ws_data, min_col=4, min_row=2, max_row=len(top_montants)+1)
        chart_mont.add_data(data_mont, titles_from_data=True)
        chart_mont.set_categories(cats_mont)
        ws.add_chart(chart_mont, "O6")