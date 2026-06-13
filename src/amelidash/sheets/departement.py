from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from components.filters import add_filter


class OngletDepartement:

    def __init__(self, excel_manager, df):
        self.excel_manager = excel_manager
        self.df = df
        self.wb = excel_manager.wb
        self.source_sheet_name = self.wb.sheetnames[
            0
        ]  # Correspond à 'cleanedData_Effectifs'

    def create(self):
        ws = self.excel_manager.create_sheet("Departement", 2)
        ws.sheet_view.showGridLines = False

        # 1. Ajout du filtre interactif (en E2)
        # La cellule de sélection sera $E$2
        add_filter(
            sheet=ws,
            label="Departement",
            default_value="Tous",
            start_row=2,
            start_col=4,  # Colonne D
            formula="='Filtres'!$D$2:$D$100",
        )

        # 2. Titre du tableau
        ws["A1"] = "Effectifs par Departement (Top 30)"
        ws["A1"].font = Font(bold=True, size=13, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color="FF004A99", fill_type="solid")

        # 3. Calcul des départements (Logique DataFrame)
        depts_sorted = sorted(
            self.df["Departement"].dropna().unique(),
            key=lambda x: self.df[self.df["Departement"] == x]["Effectif"].sum(),
            reverse=True,
        )[:30]

        # 4. Remplissage du tableau avec SUMIFS dynamique lié au filtre $E$2
        for idx, dept in enumerate(depts_sorted, start=2):
            ws.cell(idx, 1).value = dept
            ws.cell(idx, 2).value = (
                f"=SUMIFS('{self.source_sheet_name}'!E:E, "
                f"'{self.source_sheet_name}'!B:B, A{idx}, "
                f"'{self.source_sheet_name}'!D:D, $E$2)"
            )
            ws.cell(idx, 2).number_format = "#,##0"

        # 5. Création du graphique
        chart = BarChart()
        chart.type = "bar"
        chart.title = "Top 30 Departements"
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
