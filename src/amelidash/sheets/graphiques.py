import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


class OngletGraphiques:
    def __init__(self, wb, df):
        self.wb = wb
        self.df = df.copy()
        for nom in ["Graphiques", "Calc", "Source"]:
            if nom in self.wb.sheetnames:
                self.wb.remove(self.wb[nom])

        # CRUCIAL : Il manquait la création et le remplissage de la feuille Source
        self.ws_src = self.wb.create_sheet("Source")
        self._write_source_data()
        self.ws_src.sheet_state = "hidden"

        self.create()

    def _write_source_data(self):
        for c, col in enumerate(self.df.columns, 1):
            self.ws_src.cell(1, c, col)
        for r, row in enumerate(self.df.values, 2):
            for c, val in enumerate(row, 1):
                self.ws_src.cell(r, c, val)

    def create(self):
        ws = self.wb.create_sheet("Graphiques")
        calc = self.wb.create_sheet("Calc")

        # --- Mise en page du titre ---
        ws["B1"] = "Analyse Régionale des Pathologies"
        ws["B1"].font = Font(bold=True, size=14, color="FFFFFF")
        ws["B1"].fill = PatternFill("solid", fgColor="006B4F")  # Vert Ameli
        ws["B1"].alignment = Alignment(horizontal="center")
        ws.merge_cells("B1:J1")

        # --- Détection automatique des colonnes ---
        # Permet de trouver la bonne lettre (A, B, C...) peu importe l'ordre
        col_map = {
            col: get_column_letter(idx + 1) for idx, col in enumerate(self.df.columns)
        }
        l_reg = col_map.get("Région", "G")
        l_path = col_map.get("patho_niv1", "C")
        l_age = col_map.get("Classe d'age", "E")
        l_eff = col_map.get("Effectif", "J")
        l_pop = col_map.get("Population de référence", "I")

        # --- Filtres uniques ---
        pathos = sorted(self.df["patho_niv1"].dropna().astype(str).unique())
        ages = sorted(self.df["Classe d'age"].dropna().astype(str).unique())
        regions = sorted(self.df["Région"].dropna().astype(str).unique())

        ws["A3"], ws["B3"] = "Pathologie", pathos[0]
        ws["A4"], ws["B4"] = "Classe d'âge", ages[0]
        ws["A3"].font = ws["A4"].font = Font(bold=True)

        dv_p = DataValidation(
            type="list", formula1=f'"{",".join(pathos)}"', allow_blank=False
        )
        dv_a = DataValidation(
            type="list", formula1=f'"{",".join(ages)}"', allow_blank=False
        )
        ws.add_data_validation(dv_p)
        ws.add_data_validation(dv_a)
        dv_p.add("B3")
        dv_a.add("B4")

        # --- Tableau de calcul caché ---
        calc["A1"], calc["B1"], calc["C1"], calc["D1"] = (
            "Région",
            "Effectif",
            "Population",
            "Prévalence (%)",
        )

        for i, region in enumerate(regions, 2):
            calc.cell(i, 1, region)

            # Somme des effectifs pour la région, la patho et l'âge sélectionnés
            formule_eff = f"=SUMIFS(Source!{l_eff}:{l_eff}, Source!{l_reg}:{l_reg}, A{i}, Source!{l_path}:{l_path}, Graphiques!$B$3, Source!{l_age}:{l_age}, Graphiques!$B$4)"
            calc.cell(i, 2, formule_eff)

            # Somme de la population de référence
            formule_pop = f"=SUMIFS(Source!{l_pop}:{l_pop}, Source!{l_reg}:{l_reg}, A{i}, Source!{l_path}:{l_path}, Graphiques!$B$3, Source!{l_age}:{l_age}, Graphiques!$B$4)"
            calc.cell(i, 3, formule_pop)

            # Calcul du Taux de prévalence (Effectif / Population)
            formule_prev = f"=IF(C{i}=0, 0, B{i}/C{i})"
            cell_prev = calc.cell(i, 4, formule_prev)
            cell_prev.number_format = "0.00%"  # Format pourcentage propre

        # --- Graphique 1 : Taux de prévalence par région ---
        chart = BarChart()
        chart.type = "col"
        chart.title = "Taux de prévalence par Région"
        chart.height = 10
        chart.width = 18

        # Esthétique : Épuration du design
        chart.style = 2
        chart.legend = None  # Pas de légende car le titre suffit
        chart.y_axis.majorGridlines = None  # Suppression du quadrillage lourd
        chart.y_axis.delete = True  # On cache l'axe Y pour alléger

        # Affichage des valeurs directement sur les barres
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True

        # Injection des données (Colonne D = Prévalence)
        data = Reference(calc, min_col=4, min_row=1, max_row=len(regions) + 1)
        cats = Reference(calc, min_col=1, min_row=2, max_row=len(regions) + 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        # Force la couleur des barres avec votre bleu ou vert
        if chart.series:
            chart.series[0].graphicalProperties.solidFill = "006B4F"

        ws.add_chart(chart, "A7")
        calc.sheet_state = "hidden"
