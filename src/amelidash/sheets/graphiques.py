# """Onglet Graphiques : tableaux et graphiques par pathologie et sous-poste (dépenses) + par pathologie et classe d'âge (effectifs)."""

from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from config import COULEURS, SHEET_CLEANED_DEPENSES, SHEET_CLEANED_EFFECTIFS


class OngletGraphiques:
    def __init__(self, excel_manager, df):
        self.excel_manager = excel_manager
        self.df = df
        self.wb = excel_manager.wb

    def create(self):
        ws = self.excel_manager.create_sheet("Graphiques", 2)
        ws.sheet_view.showGridLines = False
        ws_dep = self.wb[SHEET_CLEANED_DEPENSES]
        # 1. Lire les en-têtes et identifier les colonnes
        headers = {}
        for idx, cell in enumerate(ws_dep[1], start=1):
            if cell.value:
                headers[str(cell.value).strip()] = idx

        col_poste = headers.get("poste de dépense", 5)
        col_sous_poste = headers.get("sous poste", 6)
        col_patho1 = headers.get("patho_niv1", 2)
        col_montant = headers.get("montant", 7)

        col_poste_letter = get_column_letter(col_poste)
        col_sous_poste_letter = get_column_letter(col_sous_poste)
        col_patho1_letter = get_column_letter(col_patho1)
        col_montant_letter = get_column_letter(col_montant)

        # 2. Extraire les listes uniques (postes et patho_niv1)
        postes_list = sorted(
            set(
                str(r[col_poste - 1]).strip()
                for r in ws_dep.iter_rows(min_row=2, values_only=True)
                if r[col_poste - 1] and "total" not in str(r[col_poste - 1]).lower()
            )
        )

        patho1_list = sorted(
            set(
                str(r[col_patho1 - 1]).strip()
                for r in ws_dep.iter_rows(min_row=2, values_only=True)
                if r[col_patho1 - 1] and "total" not in str(r[col_patho1 - 1]).lower()
            )
        )

        # 3. Titre section 1 : Par groupe de pathologies
        ws["A1"] = "Par groupe de pathologies"
        ws["A1"].font = Font(bold=True, size=13, color=COULEURS["noir"])

        # 4. Titre section 2 : Par grand poste de dépenses + filtre
        ws["H1"] = "Par grand poste de depenses"
        ws["H1"].font = Font(bold=True, color=COULEURS["blanc"], size=10)
        ws["H1"].fill = PatternFill(
            start_color=COULEURS["principal"], fill_type="solid"
        )

        ws["H2"] = postes_list[0] if postes_list else ""
        ws["H2"].font = Font(bold=True, size=10)

        dv_poste = DataValidation(
            type="list",
            formula1=f'"{",".join(postes_list)}"',
            allow_blank=False,
        )
        ws.add_data_validation(dv_poste)
        dv_poste.add("H2")

        # 5. Tableau 1 : Pathologie -> Montant (avec SUMIFS + filtre poste)
        HDR1 = 4
        for col_i, lbl in enumerate(["Pathologie", "Montant (€)"], start=1):
            cell = ws.cell(HDR1, col_i)
            cell.value = lbl
            cell.font = Font(bold=True, color=COULEURS["blanc"], size=9)
            cell.fill = PatternFill(
                start_color=COULEURS["principal"], fill_type="solid"
            )

        for idx, patho1 in enumerate(patho1_list):
            r = HDR1 + 1 + idx
            ws.cell(r, 1).value = patho1
            ws.cell(r, 2).value = (
                f"=IFERROR(SUMIFS('{ws_dep.title}'!${col_montant_letter}:${col_montant_letter},"
                f"'{ws_dep.title}'!${col_poste_letter}:${col_poste_letter},$H$2,"
                f"'{ws_dep.title}'!${col_patho1_letter}:${col_patho1_letter},A{r}),0)"
            )
            ws.cell(r, 2).number_format = "#,##0"

        END1 = HDR1 + len(patho1_list)

        # Graphique 1 : Barres (pathologie)
        chart1 = BarChart()
        chart1.type = "bar"
        chart1.title = None
        chart1.height = max(16, len(patho1_list) * 0.6)
        chart1.width = 26

        data_ref1 = Reference(ws, min_col=2, min_row=HDR1, max_row=END1)
        cat_ref1 = Reference(ws, min_col=1, min_row=HDR1 + 1, max_row=END1)

        chart1.add_data(data_ref1, titles_from_data=True)
        chart1.set_categories(cat_ref1)
        chart1.series[0].graphicalProperties.solidFill = COULEURS["principal"][2:]
        ws.add_chart(chart1, "A6")

        # 7. Section 2 : Par sous-poste de dépenses
        SECTION2_START = END1 + 6
        ws[f"A{SECTION2_START}"] = "Par sous-poste de depenses"
        ws[f"A{SECTION2_START}"].font = Font(bold=True, size=11, color=COULEURS["noir"])

        HDR2 = SECTION2_START + 3

        for col_i, lbl in enumerate(["Sous-poste", "Montant (€)"], start=1):
            cell = ws.cell(HDR2, col_i)
            cell.value = lbl
            cell.font = Font(bold=True, color=COULEURS["blanc"], size=9)
            cell.fill = PatternFill(
                start_color=COULEURS["principal"], fill_type="solid"
            )

        # 8. Extraire les sous-postes uniques (sans total)
        sous_poste_list = sorted(
            set(
                str(r[col_sous_poste - 1]).strip()
                for r in ws_dep.iter_rows(min_row=2, values_only=True)
                if r[col_sous_poste - 1]
                and "total" not in str(r[col_sous_poste - 1]).lower()
            )
        )

        # 9. Tableau 2 : Sous-poste -> Montant (avec SUMIFS + filtre poste)
        for idx, sous_poste in enumerate(sous_poste_list):
            r = HDR2 + 1 + idx
            ws.cell(r, 1).value = sous_poste
            ws.cell(r, 2).value = (
                f"=IFERROR(SUMIFS('{ws_dep.title}'!${col_montant_letter}:${col_montant_letter},"
                f"'{ws_dep.title}'!${col_poste_letter}:${col_poste_letter},$H$2,"
                f"'{ws_dep.title}'!${col_sous_poste_letter}:${col_sous_poste_letter},A{r}),0)"
            )
            ws.cell(r, 2).number_format = "#,##0"

        END2 = HDR2 + len(sous_poste_list)

        # 10. Graphique 2 : Donut (sous-poste)
        chart2 = DoughnutChart()
        chart2.title = None
        chart2.height = 14
        chart2.width = 18

        data_ref2 = Reference(ws, min_col=2, min_row=HDR2, max_row=END2)
        cat_ref2 = Reference(ws, min_col=1, min_row=HDR2 + 1, max_row=END2)

        chart2.add_data(data_ref2, titles_from_data=True)
        chart2.set_categories(cat_ref2)
        chart2.dataLabels.showPercent = True
        ws.add_chart(chart2, "D" + str(SECTION2_START + 2))

        # 11. Largeurs de colonnes
        ws.column_dimensions["A"].width = 50
        ws.column_dimensions["B"].width = 16


class OngletGraphiquesEffectifs:

    def __init__(self, excel_manager, df):
        self.excel_manager = excel_manager
        self.df = df
        self.wb = excel_manager.wb

    def create(self):
        ws = self.excel_manager.create_sheet("Graphiques_Effectifs", 3)
        ws.sheet_view.showGridLines = False

        ws_eff = self.wb[SHEET_CLEANED_EFFECTIFS]

        # 1. Lire les en-têtes et identifier les colonnes
        headers = {}
        for idx, cell in enumerate(ws_eff[1], start=1):
            if cell.value:
                headers[str(cell.value).strip()] = idx

        col_patho = headers.get("patho_niv1", 2)
        col_age = headers.get("Classe d'age", 5)
        col_effectif = headers.get("Effectif", 7)
        col_dept = headers.get("Département", 4)

        col_patho_letter = get_column_letter(col_patho)
        col_age_letter = get_column_letter(col_age)
        col_effectif_letter = get_column_letter(col_effectif)
        col_dept_letter = get_column_letter(col_dept)

        # 2. Extraire les listes uniques
        patho_list = sorted(
            set(
                str(r[col_patho - 1]).strip()
                for r in ws_eff.iter_rows(min_row=2, values_only=True)
                if r[col_patho - 1] and "total" not in str(r[col_patho - 1]).lower()
            )
        )

        age_list = sorted(
            set(
                str(r[col_age - 1]).strip()
                for r in ws_eff.iter_rows(min_row=2, values_only=True)
                if r[col_age - 1]
            )
        )

        dept_list = sorted(
            set(
                str(r[col_dept - 1]).strip()
                for r in ws_eff.iter_rows(min_row=2, values_only=True)
                if r[col_dept - 1]
            )
        )

        dept_defaut = dept_list[0] if dept_list else "Tous"

        # 3. Titre
        ws["A1"] = "Effectifs par pathologie et classe d'âge"
        ws["A1"].font = Font(bold=True, size=13, color=COULEURS["blanc"])
        ws["A1"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")

        # 4. Filtre Département
        ws["A3"] = "Département"
        ws["A3"].font = Font(bold=True, color=COULEURS["blanc"], size=9)
        ws["A3"].fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")

        ws["B3"] = dept_defaut
        ws["B3"].font = Font(bold=True, size=10)

        dv_dept = DataValidation(
            type="list",
            formula1=f'"{",".join(dept_list)}"',
            allow_blank=False,
        )
        ws.add_data_validation(dv_dept)
        dv_dept.add("B3")

        # 5. Tableau pivot : Pathologie x Classe d'âge (en %)
        start_col = 5  # Colonne E
        HDR_TABLE = 5

        # En-têtes colonnes (classes d'âge)
        for col_idx, age in enumerate(age_list):
            cell = ws.cell(HDR_TABLE, start_col + col_idx)
            cell.value = age
            cell.font = Font(bold=True, color=COULEURS["blanc"], size=9)
            cell.fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")

        # En-tête ligne (Pathologie)
        cell_patho_header = ws.cell(HDR_TABLE, start_col - 1)
        cell_patho_header.value = "Pathologie"
        cell_patho_header.font = Font(bold=True, color=COULEURS["blanc"], size=9)
        cell_patho_header.fill = PatternFill(start_color=COULEURS["principal"], fill_type="solid")

        # Données : calcul des pourcentages par pathologie
        for row_idx, patho in enumerate(patho_list):
            ws.cell(HDR_TABLE + 1 + row_idx, start_col - 1).value = patho
            for col_idx, age in enumerate(age_list):
                cell = ws.cell(HDR_TABLE + 1 + row_idx, start_col + col_idx)
                # Pourcentage = (effectif patho+age / total patho) * 100
                cell.value = (
                    f"=IFERROR("
                    f"(SUMIFS('{ws_eff.title}'!${col_effectif_letter}:${col_effectif_letter},"
                    f"'{ws_eff.title}'!${col_patho_letter}:${col_patho_letter},A{HDR_TABLE + 1 + row_idx},"
                    f"'{ws_eff.title}'!${col_age_letter}:${col_age_letter},E{HDR_TABLE + col_idx},"
                    f"'{ws_eff.title}'!${col_dept_letter}:${col_dept_letter},$B$3))/"
                    f"SUMIFS('{ws_eff.title}'!${col_effectif_letter}:${col_effectif_letter},"
                    f"'{ws_eff.title}'!${col_patho_letter}:${col_patho_letter},A{HDR_TABLE + 1 + row_idx},"
                    f"'{ws_eff.title}'!${col_dept_letter}:${col_dept_letter},$B$3)"
                    f"*100,0)"
                )
                cell.number_format = '0.0"%"'
        END_TABLE = HDR_TABLE + len(patho_list)
        # 6. Graphique : Barres empilées 100% (Pathologie x Classe d'âge)
        chart = BarChart()
        chart.type = "col"  # vertical
        chart.grouping = "percentStacked"  # 100% empilé
        chart.overlap = 100  # OBLIGATOIRE pour empilé dans openpyxl
        chart.title = f"Répartition des classes d'âge par Pathologie (100%) - {dept_defaut}"
        chart.style = 10
        chart.y_axis.title = "%"
        chart.x_axis.title = "Pathologie"
        # Données
        data_ref = Reference(
            ws,
            min_col=start_col,
            min_row=HDR_TABLE,
            max_row=END_TABLE,
            max_col=start_col + len(age_list) - 1
        )
        chart.add_data(data_ref, titles_from_data=True)

        # Catégories = pathologies
        cats_ref = Reference(
            ws,
            min_col=start_col - 1,
            min_row=HDR_TABLE + 1,
            max_row=END_TABLE
        )
        chart.set_categories(cats_ref)

        # Légende
        chart.legend.position = "r"
        chart.legend.include_in_layout = False

        # Étiquettes de pourcentage
        from openpyxl.chart.label import DataLabelList
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True
        chart.dataLabels.showVal = False

        chart.height = 16
        chart.width = 24
        ws.add_chart(chart, "A8")

        # 7. Largeurs de colonnes
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 18
        for col_idx in range(start_col - 1, start_col + len(age_list)):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 12