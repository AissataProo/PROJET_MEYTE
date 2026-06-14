import pandas as pd
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from config import SHEET_CLEANED_EFFECTIFS


class OngletDepartement:
    """Gère la création de l’onglet départemental : initialise le classeur,
    stocke les DataFrames d’entrée (données principales et dépenses éventuelles)
    et prépare la génération de l’onglet via la méthode `create`."""

    def __init__(self, wb, df, df_dep=None):
        self.wb = wb
        self.df = df.copy()
        self.df_dep = df_dep.copy() if df_dep is not None else None

    def create(self):

        # 1. Nettoyage des anciennes feuilles
        for nom in ["Departement", "CalcDeptPatho", "CalcDeptSexe", "CalcDemo"]:
            if nom in self.wb.sheetnames:
                self.wb.remove(self.wb[nom])

        df = self.df.copy()
        df["Effectif"] = pd.to_numeric(df["Effectif"], errors="coerce").fillna(0)

        ws_eff = self.wb[SHEET_CLEANED_EFFECTIFS]
        sheet_eff = SHEET_CLEANED_EFFECTIFS

        # Mapping colonnes
        hdr = {
            str(c.value).strip(): get_column_letter(i)
            for i, c in enumerate(ws_eff[1], start=1)
            if c.value
        }

        L_annee = hdr["annee"]
        L_sexe = hdr["Sexe"]
        L_patho = hdr["patho_niv1"]
        L_age = hdr["Classe d'age"]
        L_dept = hdr["Département"]
        L_eff = hdr["Effectif"]
        L_pop = hdr["Population de référence"]

        # Listes utiles
        pathos = sorted(df["patho_niv1"].dropna().unique())
        deps = sorted(df["Département"].dropna().unique())
        sexes = sorted(
            [
                s
                for s in df["Sexe"].dropna().unique()
                if "ensemble" not in str(s).lower()
            ]
        )
        annees = sorted(df["annee"].dropna().unique(), reverse=True)
        ages = sorted(df["Classe d'age"].dropna().unique())

        # 2. Création des feuilles
        ws = self.wb.create_sheet("Departement")
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A8"

        calcp = self.wb.create_sheet("CalcDeptPatho")
        calcp.sheet_state = "hidden"

        calcs = self.wb.create_sheet("CalcDeptSexe")
        calcs.sheet_state = "hidden"

        cal = self.wb.create_sheet("CalcDemo")
        cal.sheet_state = "hidden"

        COLOR_BLEU = "FF4472C4"
        COLOR_ROUGE = "C00000"

        # --- Filtres (Pathologie, Département, Année) ---
        for i, (lbl, val) in enumerate(
            [
                ("Pathologie", pathos[0]),
                ("Département", deps[0]),
                ("Année", int(annees[0])),
            ],
            3,
        ):
            ws[f"A{i}"] = lbl
            ws[f"A{i}"].fill = PatternFill("solid", fgColor=COLOR_BLEU)
            ws[f"A{i}"].font = Font(bold=True, color="FFFFFF")
            ws[f"B{i}"] = val
            ws[f"B{i}"].font = Font(bold=True)

        # --- Génération CalcDeptPatho (Pathologie x Sexe, filtré sur Département + Année) ---
        calcp["A1"] = "Pathologie"
        for i, s in enumerate(sexes, 2):
            calcp.cell(1, i, s)

        for r, p in enumerate(pathos, 2):
            calcp.cell(r, 1, p)
            for c, s in enumerate(sexes, 2):
                sexe_col = get_column_letter(c)
                # critère sexe = cellule d’en-tête de CalcDeptPatho (B1, C1, ...)
                calcp.cell(r, c).value = (
                    f"=SUMIFS('{sheet_eff}'!{L_eff}:{L_eff},"
                    f"'{sheet_eff}'!{L_patho}:{L_patho},$A{r},"
                    f"'{sheet_eff}'!{L_sexe}:{L_sexe},'{calcp.title}'!${sexe_col}$1,"
                    f"'{sheet_eff}'!{L_dept}:{L_dept},Departement!$B$4,"
                    f"'{sheet_eff}'!{L_annee}:{L_annee},Departement!$B$5)"
                )

        # --- Graphique Patho x Sexe ---
        chart1 = BarChart()
        chart1.title = "Effectif par pathologie et sexe"
        chart1.add_data(
            Reference(
                calcp,
                min_col=2,
                max_col=1 + len(sexes),
                min_row=1,
                max_row=len(pathos) + 1,
            ),
            titles_from_data=True,
        )
        chart1.set_categories(
            Reference(calcp, min_col=1, min_row=2, max_row=len(pathos) + 1)
        )
        ws.add_chart(chart1, "A8")

        # --- Bloc indicateurs ---
        ws.merge_cells("I3:M3")
        ws["I3"] = "Indicateurs clés"
        ws["I3"].fill = PatternFill("solid", fgColor=COLOR_BLEU)
        ws["I3"].font = Font(bold=True, color="FFFFFF")

        ws["I4"] = "Total patients :"
        # On filtre sur Département + Année + Pathologie (B3, B4, B5)
        ws["J4"].value = (
            f"=SUMIFS('{sheet_eff}'!{L_eff}:{L_eff},"
            f"'{sheet_eff}'!{L_dept}:{L_dept},B4,"
            f"'{sheet_eff}'!{L_annee}:{L_annee},B5,"
            f"'{sheet_eff}'!{L_patho}:{L_patho},B3)"
        )

        # --- Variation annuelle ---
        ws.merge_cells("I7:M7")
        ws["I7"] = "Variation annuelle du nombre de patients"
        ws["I7"].fill = PatternFill("solid", fgColor=COLOR_ROUGE)
        ws["I7"].font = Font(bold=True, color="FFFFFF")

        headers = ["Année", "Nb patient", "N-1", "Var %"]
        for j, h in enumerate(headers, 9):
            cell = ws.cell(8, j, h)
            cell.fill = PatternFill("solid", fgColor=COLOR_BLEU)
            cell.font = Font(bold=True, color="FFFFFF")

        for k, an in enumerate(sorted(int(a) for a in annees)):
            r = 9 + k
            ws.cell(r, 9, an)
            ws.cell(r, 10).value = (
                f"=SUMIFS('{sheet_eff}'!{L_eff}:{L_eff},"
                f"'{sheet_eff}'!{L_annee}:{L_annee},I{r},"
                f"'{sheet_eff}'!{L_patho}:{L_patho},B3,"
                f"'{sheet_eff}'!{L_dept}:{L_dept},B4)"
            )
            if k > 0:
                ws.cell(r, 11).value = f"=J{r-1}"
                ws.cell(r, 12).value = f"=IFERROR((J{r}-K{r})/K{r},0)"
                ws.cell(r, 12).number_format = "0.00%"

        # --- Camembert Sexe (filtré sur Pathologie + Département + Année) ---
        calcs["A1"], calcs["B1"] = "Sexe", "Effectif"
        for i, s in enumerate(sexes, 2):
            calcs.cell(i, 1, s)
            calcs.cell(i, 2).value = (
                f"=SUMIFS('{sheet_eff}'!{L_eff}:{L_eff},"
                f"'{sheet_eff}'!{L_sexe}:{L_sexe},A{i},"
                f"'{sheet_eff}'!{L_dept}:{L_dept},Departement!$B$4,"
                f"'{sheet_eff}'!{L_annee}:{L_annee},Departement!$B$5,"
                f"'{sheet_eff}'!{L_patho}:{L_patho},Departement!$B$3)"
            )

        pie = PieChart()
        pie.title = "Répartition Homme / Femme"
        pie.add_data(
            Reference(calcs, min_col=2, min_row=1, max_row=len(sexes) + 1),
            titles_from_data=True,
        )
        pie.set_categories(
            Reference(calcs, min_col=1, min_row=2, max_row=len(sexes) + 1)
        )
        pie.dataLabels = DataLabelList(showPercent=True)
        ws.add_chart(pie, "I18")

        # --- Pyramide des âges (filtrée sur Pathologie + Département + Année) ---
        cal["A1"], cal["B1"], cal["C1"] = "Classe d'âge", "Hommes", "Femmes"

        for i, age in enumerate(ages, 2):
            cal.cell(i, 1, age)

            # Hommes en négatif
            cal.cell(i, 2).value = (
                f"=-SUMIFS('{sheet_eff}'!{L_eff}:{L_eff},"
                f"'{sheet_eff}'!{L_age}:{L_age},A{i},"
                f"'{sheet_eff}'!{L_sexe}:{L_sexe},\"*Hom*\","
                f"'{sheet_eff}'!{L_patho}:{L_patho},Departement!$B$3,"
                f"'{sheet_eff}'!{L_annee}:{L_annee},Departement!$B$5,"
                f"'{sheet_eff}'!{L_dept}:{L_dept},Departement!$B$4)"
            )

            # Femmes en positif
            cal.cell(i, 3).value = (
                f"=SUMIFS('{sheet_eff}'!{L_eff}:{L_eff},"
                f"'{sheet_eff}'!{L_age}:{L_age},A{i},"
                f"'{sheet_eff}'!{L_sexe}:{L_sexe},\"*Fem*\","
                f"'{sheet_eff}'!{L_patho}:{L_patho},Departement!$B$3,"
                f"'{sheet_eff}'!{L_annee}:{L_annee},Departement!$B$5,"
                f"'{sheet_eff}'!{L_dept}:{L_dept},Departement!$B$4)"
            )

        pyra = BarChart()
        pyra.type = "bar"
        pyra.grouping = "stacked"
        pyra.title = "Pyramide des âges"
        pyra.y_axis.title = "Effectif"
        pyra.x_axis.title = "Classe d'âge"

        data = Reference(cal, min_col=2, max_col=3, min_row=1, max_row=len(ages) + 1)
        cats = Reference(cal, min_col=1, min_row=2, max_row=len(ages) + 1)
        pyra.add_data(data, titles_from_data=True)
        pyra.set_categories(cats)
        pyra.height, pyra.width = 12, 18

        ws.add_chart(pyra, "A25")

        # Mise en forme finale
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 40
        ws.column_dimensions["I"].width = 20
