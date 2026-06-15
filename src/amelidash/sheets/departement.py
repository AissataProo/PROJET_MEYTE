import pandas as pd
from openpyxl.chart import (
    BarChart,
    PieChart,
    Reference,
    LineChart,
    ScatterChart,
    Series,
)
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from config import SHEET_CLEANED_EFFECTIFS


class OngletDepartement:
    """Gère la création de l'onglet départemental avec filtres simples (Dept + Année).
    Affiche les pathologies les plus présentes par département."""

    def __init__(self, wb, df, df_dep=None):
        self.wb = wb
        self.df = df.copy()
        self.df_dep = df_dep.copy() if df_dep is not None else None
        self.sheet_eff = SHEET_CLEANED_EFFECTIFS

    def _create_filter_sheet(self):
        """Crée feuille FiltresDept avec listes uniques pour les filtres."""

        if "FiltresDept" in self.wb.sheetnames:
            self.wb.remove(self.wb["FiltresDept"])

        wf = self.wb.create_sheet("FiltresDept", 0)
        wf.sheet_state = "hidden"

        df = self.df.copy()
        df["Effectif"] = pd.to_numeric(df["Effectif"], errors="coerce").fillna(0)

        deps = sorted(df["Département"].dropna().unique())
        annees = sorted([int(a) for a in df["annee"].dropna().unique()], reverse=True)
        pathos = sorted(df["patho_niv1"].dropna().unique())
        sexes = sorted(
            [
                s
                for s in df["Sexe"].dropna().unique()
                if "ensemble" not in str(s).lower()
            ]
        )
        ages = sorted(df["Classe d'age"].dropna().unique())

        # Colonnes : A=deps, B=annees, C=pathos, D=sexes, E=ages
        for i, val in enumerate(deps, 1):
            wf.cell(i, 1, val)
        for i, val in enumerate(annees, 1):
            wf.cell(i, 2, val)
        for i, val in enumerate(pathos, 1):
            wf.cell(i, 3, val)
        for i, val in enumerate(sexes, 1):
            wf.cell(i, 4, val)
        for i, val in enumerate(ages, 1):
            wf.cell(i, 5, val)

        return {
            "deps": deps,
            "annees": annees,
            "pathos": pathos,
            "sexes": sexes,
            "ages": ages,
        }

    def _get_column_mapping(self):
        """Récupère le mapping colonnes de la feuille nettoyée."""
        ws_eff = self.wb[self.sheet_eff]
        hdr = {
            str(c.value).strip(): get_column_letter(i)
            for i, c in enumerate(ws_eff[1], start=1)
            if c.value
        }
        return {
            "annee": hdr.get("annee"),
            "sexe": hdr.get("Sexe"),
            "patho": hdr.get("patho_niv1"),
            "age": hdr.get("Classe d'age"),
            "dept": hdr.get("Département"),
            "eff": hdr.get("Effectif"),
            "region": hdr.get("Région"),
            "prevalence": hdr.get("Prévalence"),
            "pop_ref": hdr.get("Population de référence"),
        }

    def _add_filter_with_validation(self, ws, row, label, list_range):
        """Ajoute label + cellule de filtre avec DataValidation."""
        COLOR_BLEU = "FF4472C4"

        ws[f"A{row}"] = label
        ws[f"A{row}"].fill = PatternFill("solid", fgColor=COLOR_BLEU)
        ws[f"A{row}"].font = Font(bold=True, color="FFFFFF")

        dv = DataValidation(
            type="list", formula1=list_range, allow_blank=False, showDropDown=True
        )
        dv.error = f"Sélectionner une valeur valide"
        dv.errorTitle = "Entrée invalide"
        dv.prompt = f"Choisir {label}"
        dv.promptTitle = label

        ws.add_data_validation(dv)
        dv.add(f"B{row}")

        return ws[f"B{row}"]

    def _create_chart(self, chart_type, title):
        """Crée et formate un graphique."""
        if chart_type == "BarChart":
            chart = BarChart()
        elif chart_type == "LineChart":
            chart = LineChart()
            chart.smooth = True
        else:
            chart = PieChart()

        chart.title = title
        chart.style = 12
        chart.legend.position = "r"
        return chart

    def create(self):
        """Crée l'onglet Departement avec filtres Département + Année."""

        # 1. Crée feuille de filtres
        lists = self._create_filter_sheet()
        col_map = self._get_column_mapping()

        # 2. Nettoyage des anciennes feuilles
        for nom in [
            "Departement",
            "CalcTopPatho",
            "CalcPathoSexe",
            "CalcVariation",
            "CalcPrevAge",
        ]:
            if nom in self.wb.sheetnames:
                self.wb.remove(self.wb[nom])

        # 3. Création des feuilles
        ws = self.wb.create_sheet("Departement")
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A8"

        calc_top_patho = self.wb.create_sheet("CalcTopPatho")
        calc_top_patho.sheet_state = "hidden"

        calc_patho_sexe = self.wb.create_sheet("CalcPathoSexe")
        calc_patho_sexe.sheet_state = "hidden"

        calc_variation = self.wb.create_sheet("CalcVariation")
        calc_variation.sheet_state = "hidden"

        calc_prev_age = self.wb.create_sheet("CalcPrevAge")
        calc_prev_age.sheet_state = "hidden"

        COLOR_BLEU = "FF4472C4"
        COLOR_ROUGE = "C00000"

        # 4. FILTRES : Département + Année seulement
        n_deps = len(lists["deps"])
        n_annees = len(lists["annees"])

        self._add_filter_with_validation(
            ws, 3, "Département", f"FiltresDept!$A$1:$A${n_deps}"
        )
        self._add_filter_with_validation(
            ws, 4, "Année", f"FiltresDept!$B$1:$B${n_annees}"
        )

        ws["B3"].value = lists["deps"][0]
        ws["B4"].value = lists["annees"][0]

        # 5. INDICATEURS CLÉS
        ws.merge_cells("E3:H3")
        ws["E3"] = "Indicateurs clés"
        ws["E3"].fill = PatternFill("solid", fgColor=COLOR_BLEU)
        ws["E3"].font = Font(bold=True, color="FFFFFF")

        ws["E4"] = "Total patients :"
        ws["F4"].value = (
            f"=SUMIFS('{self.sheet_eff}'!{col_map['eff']}:{col_map['eff']},"
            f"'{self.sheet_eff}'!{col_map['dept']}:{col_map['dept']},B3,"
            f"'{self.sheet_eff}'!{col_map['annee']}:{col_map['annee']},B4)"
        )

        ws["E5"] = "Population de référence :"
        ws["F5"].value = (
            f"=SUMIFS('{self.sheet_eff}'!{col_map['pop_ref']}:{col_map['pop_ref']},"
            f"'{self.sheet_eff}'!{col_map['dept']}:{col_map['dept']},B3,"
            f"'{self.sheet_eff}'!{col_map['annee']}:{col_map['annee']},B4)"
        )

        ws["E6"] = "Taux couverture :"
        ws["F6"].value = f"=IF(F5=0,0,F4/F5*100)"
        ws["F6"].number_format = '0.00"%"'

        ws["E7"] = "Pathologie principale :"
        ws["F7"].value = (
            "=INDEX(CalcTopPatho!A:A,MATCH(MAX(CalcTopPatho!B:B),CalcTopPatho!B:B,0))"
        )

        # 6. TOP PATHOLOGIES PAR EFFECTIF (Bar chart horizontal)
        calc_top_patho["A1"] = "Pathologie"
        calc_top_patho["B1"] = "Effectif"

        for i, patho in enumerate(lists["pathos"], 2):
            calc_top_patho.cell(i, 1, patho)
            calc_top_patho.cell(i, 2).value = (
                f"=SUMIFS('{self.sheet_eff}'!{col_map['eff']}:{col_map['eff']},"
                f"'{self.sheet_eff}'!{col_map['patho']}:{col_map['patho']},A{i},"
                f"'{self.sheet_eff}'!{col_map['dept']}:{col_map['dept']},Departement!$B$3,"
                f"'{self.sheet_eff}'!{col_map['annee']}:{col_map['annee']},Departement!$B$4)"
            )

        bar_patho = self._create_chart("BarChart", "Top pathologies par effectif")
        bar_patho.type = "bar"  # Horizontal
        bar_patho.add_data(
            Reference(
                calc_top_patho, min_col=2, min_row=1, max_row=len(lists["pathos"]) + 1
            ),
            titles_from_data=True,
        )
        bar_patho.set_categories(
            Reference(
                calc_top_patho, min_col=1, min_row=2, max_row=len(lists["pathos"]) + 1
            )
        )
        bar_patho.height, bar_patho.width = 14, 16
        ws.add_chart(bar_patho, "A8")

        # 7. PATHOLOGIES PAR SEXE
        calc_patho_sexe["A1"] = "Pathologie"
        for i, s in enumerate(lists["sexes"], 2):
            calc_patho_sexe.cell(1, i, s)

        for r, p in enumerate(lists["pathos"], 2):
            calc_patho_sexe.cell(r, 1, p)
            for c, s in enumerate(lists["sexes"], 2):
                sexe_col = get_column_letter(c)
                calc_patho_sexe.cell(r, c).value = (
                    f"=SUMIFS('{self.sheet_eff}'!{col_map['eff']}:{col_map['eff']},"
                    f"'{self.sheet_eff}'!{col_map['patho']}:{col_map['patho']},$A{r},"
                    f"'{self.sheet_eff}'!{col_map['sexe']}:{col_map['sexe']},'{calc_patho_sexe.title}'!${sexe_col}$1,"
                    f"'{self.sheet_eff}'!{col_map['dept']}:{col_map['dept']},Departement!$B$3,"
                    f"'{self.sheet_eff}'!{col_map['annee']}:{col_map['annee']},Departement!$B$4)"
                )

        chart_patho_sexe = self._create_chart(
            "BarChart", "Effectif par pathologie et sexe"
        )
        chart_patho_sexe.add_data(
            Reference(
                calc_patho_sexe,
                min_col=2,
                max_col=1 + len(lists["sexes"]),
                min_row=1,
                max_row=len(lists["pathos"]) + 1,
            ),
            titles_from_data=True,
        )
        chart_patho_sexe.set_categories(
            Reference(
                calc_patho_sexe, min_col=1, min_row=2, max_row=len(lists["pathos"]) + 1
            )
        )
        chart_patho_sexe.height, chart_patho_sexe.width = 12, 18
        ws.add_chart(chart_patho_sexe, "K8")

        # 8. VARIATION ANNUELLE DU NOMBRE DE PATIENTS
        calc_variation["A1"] = "Année"
        calc_variation["B1"] = "Nb patient"
        calc_variation["C1"] = "N-1"
        calc_variation["D1"] = "Var %"

        for k, an in enumerate(sorted(lists["annees"]), 1):
            r = k + 1
            calc_variation.cell(r, 1, an)
            calc_variation.cell(r, 2).value = (
                f"=SUMIFS('{self.sheet_eff}'!{col_map['eff']}:{col_map['eff']},"
                f"'{self.sheet_eff}'!{col_map['annee']}:{col_map['annee']},A{r},"
                f"'{self.sheet_eff}'!{col_map['dept']}:{col_map['dept']},Departement!$B$3)"
            )
            if k > 1:
                calc_variation.cell(r, 3).value = f"=B{r-1}"
                calc_variation.cell(r, 4).value = f"=IFERROR((B{r}-C{r})/C{r},0)"
                calc_variation.cell(r, 4).number_format = "0.00%"

        # Afficher tableau variation dans le dashboard
        ws.merge_cells("K3:M3")
        ws["K3"] = "Variation annuelle"
        ws["K3"].fill = PatternFill("solid", fgColor=COLOR_BLEU)
        ws["K3"].font = Font(bold=True, color="FFFFFF")

        # Headers du tableau (sans fusion, écrire directement)
        headers_var = ["Année", "Nb patient", "Var %"]
        for j, h in enumerate(headers_var, 11):
            cell = ws.cell(4, j)
            cell.value = h
            cell.fill = PatternFill("solid", fgColor=COLOR_BLEU)
            cell.font = Font(bold=True, color="FFFFFF")

        for k, an in enumerate(sorted(lists["annees"]), 1):
            r = k + 4
            ws.cell(r, 11, an)
            ws.cell(r, 12).value = f"=CalcVariation!B{k+1}"
            if k > 1:
                ws.cell(r, 13).value = f"=IFERROR((L{r}-L{r-1})/L{r-1},0)"
                ws.cell(r, 13).number_format = "0.00%"

        # 9. PRÉVALENCE PAR CLASSE D'ÂGE
        calc_prev_age["A1"] = "Classe d'âge"
        calc_prev_age["B1"] = "Prévalence (%)"

        for i, age in enumerate(lists["ages"], 2):
            calc_prev_age.cell(i, 1, age)
            if col_map["prevalence"]:
                calc_prev_age.cell(i, 2).value = (
                    f"=IFERROR(SUMIFS('{self.sheet_eff}'!{col_map['prevalence']}:{col_map['prevalence']},"
                    f"'{self.sheet_eff}'!{col_map['age']}:{col_map['age']},A{i},"
                    f"'{self.sheet_eff}'!{col_map['dept']}:{col_map['dept']},Departement!$B$3,"
                    f"'{self.sheet_eff}'!{col_map['annee']}:{col_map['annee']},Departement!$B$4)"
                    f"/COUNTIFS('{self.sheet_eff}'!{col_map['age']}:{col_map['age']},A{i},"
                    f"'{self.sheet_eff}'!{col_map['dept']}:{col_map['dept']},Departement!$B$3,"
                    f"'{self.sheet_eff}'!{col_map['annee']}:{col_map['annee']},Departement!$B$4),0)"
                )
            else:
                calc_prev_age.cell(i, 2).value = 0

        line_prev_age = self._create_chart("LineChart", "Prévalence par classe d'âge")
        line_prev_age.add_data(
            Reference(
                calc_prev_age, min_col=2, min_row=1, max_row=len(lists["ages"]) + 1
            ),
            titles_from_data=True,
        )
        line_prev_age.set_categories(
            Reference(
                calc_prev_age, min_col=1, min_row=2, max_row=len(lists["ages"]) + 1
            )
        )
        line_prev_age.height, line_prev_age.width = 10, 16
        ws.add_chart(line_prev_age, "K28")

        # Mise en forme
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["E"].width = 20
        ws.column_dimensions["F"].width = 18
