"""Onglet Postes : détails des dépenses par pathologie + camembert répartition des postes."""

from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from config import SHEET_CLEANED_DEPENSES


class OngletPostes:
    def __init__(self, wb, df):
        self.wb = wb
        self.df = df
        self.sheet_cleaned_dep = SHEET_CLEANED_DEPENSES
        self.create()

    def create(self):
        ws = self.wb.create_sheet("Postes")
        ws.sheet_view.showGridLines = False

        ws_dep = self.wb[self.sheet_cleaned_dep]
        sheet_dep = self.sheet_cleaned_dep

        headers = {}
        for idx, cell in enumerate(ws_dep[1], start=1):
            if cell.value is not None:
                headers[str(cell.value).strip()] = idx

        col_annee = headers.get("annee", 1)
        col_patho1 = headers.get("patho_niv1", 2)
        col_poste = headers.get("poste de dépense", 4)
        col_montant = headers.get("montant", 6)

        L_annee = get_column_letter(col_annee)
        L_patho1 = get_column_letter(col_patho1)
        L_poste = get_column_letter(col_poste)
        L_montant = get_column_letter(col_montant)

        annees_list = sorted(self.df["annee"].dropna().unique(), reverse=True)
        annee_def = int(annees_list[0]) if len(annees_list) else None

        postes_list = sorted(
            {
                str(v).strip()
                for v in self.df["poste de dépense"].dropna().unique()
                if v
                and "total" not in str(v).lower()
                and "soins" not in str(v).lower()
                and str(v).strip().lower() != "nan"
            }
        )
        poste_def = postes_list[0] if postes_list else "Sélectionner"

        pathos_list = sorted(
            {
                str(v).strip()
                for v in self.df["patho_niv1"].dropna().unique()
                if v and "total" not in str(v).lower()
            }
        )

        COLOR_ROUGE = "C00000"
        COLOR_BLEU = "FF4472C4"

        bord_h = Border(*([Side(style="thin", color="4472C4")] * 4))
        bord_d = Border(*([Side(style="thin", color="CCCCCC")] * 4))
        alt1 = PatternFill("solid", fgColor="FFF5F5F5")
        alt2 = PatternFill("solid", fgColor="FFFFFFFF")

        # Fond léger rouge + onglet
        ws.sheet_properties.tabColor = COLOR_ROUGE
        fond = PatternFill("solid", fgColor="FCE9E9")
        for row in range(1, 60):
            for col in range(1, 16):
                ws.cell(row, col).fill = fond

        # --- Filtres ---
        ws["A2"] = "Année"
        ws["A2"].font = Font(bold=True, color="FFFFFF")
        ws["A2"].fill = PatternFill("solid", fgColor=COLOR_BLEU)
        ws["A2"].alignment = Alignment(horizontal="center")
        ws["B2"] = annee_def
        ws["B2"].font = Font(bold=True)
        ws["B2"].alignment = Alignment(horizontal="center")

        ws["A4"] = "Poste"
        ws["A4"].font = Font(bold=True, color="FFFFFF")
        ws["A4"].fill = PatternFill("solid", fgColor=COLOR_BLEU)
        ws["A4"].alignment = Alignment(horizontal="center")
        ws["B4"] = poste_def
        ws["B4"].font = Font(bold=True)
        ws["B4"].alignment = Alignment(horizontal="center")

        dv_a = DataValidation(
            type="list",
            formula1=f'"{",".join(str(int(a)) for a in annees_list)}"',
            allow_blank=False,
        )
        dv_p = DataValidation(
            type="list", formula1=f'"{",".join(postes_list)}"', allow_blank=False
        )
        ws.add_data_validation(dv_a)
        ws.add_data_validation(dv_p)
        dv_a.add("B2")
        dv_p.add("B4")

        # --- KPI total ---
        ws["A6"] = "Total des dépenses (poste)"
        ws["A6"].font = Font(bold=True, size=11, color="FFFFFF")
        ws["A6"].fill = PatternFill("solid", fgColor=COLOR_ROUGE)
        ws["A6"].alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells("A6:B6")
        ws["C6"] = (
            f"=IFERROR(SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
            f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2,"
            f"'{sheet_dep}'!${L_poste}:${L_poste},$B$4),0)"
        )
        ws["C6"].font = Font(bold=True, size=13, color=COLOR_ROUGE)
        ws["C6"].number_format = "#,##0 €"

        # --- Détails : dépense par pathologie ---
        ws.merge_cells("A8:C8")
        ws["A8"] = "Détails des prestations — dépense par pathologie"
        ws["A8"].font = Font(bold=True, size=12, color="FFFFFF")
        ws["A8"].fill = PatternFill("solid", fgColor=COLOR_ROUGE)
        ws["A8"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[8].height = 22

        ws.cell(9, 1, "Pathologie")
        ws.cell(9, 2, "Dépense (€)")
        ws.cell(9, 3, "Part %")
        for c in [ws.cell(9, 1), ws.cell(9, 2), ws.cell(9, 3)]:
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = PatternFill("solid", fgColor=COLOR_BLEU)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = bord_h
        ws.row_dimensions[9].height = 20

        for idx, patho in enumerate(pathos_list):
            r = 10 + idx
            coul = alt1 if idx % 2 == 0 else alt2
            ws.cell(r, 1, patho)
            ws.cell(r, 1).fill = coul
            ws.cell(r, 1).border = bord_d
            ws.cell(r, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.cell(r, 1).font = Font(size=10)

            ws.cell(r, 2).value = (
                f"=IFERROR(SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
                f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2,"
                f"'{sheet_dep}'!${L_poste}:${L_poste},$B$4,"
                f"'{sheet_dep}'!${L_patho1}:${L_patho1},A{r}),0)"
            )
            ws.cell(r, 2).number_format = "#,##0 €"
            ws.cell(r, 2).fill = coul
            ws.cell(r, 2).border = bord_d
            ws.cell(r, 2).alignment = Alignment(horizontal="right")
            ws.cell(r, 2).font = Font(size=10)

            ws.cell(r, 3).value = f"=IFERROR(B{r}/$C$6,0)"
            ws.cell(r, 3).number_format = "0.0%"
            ws.cell(r, 3).fill = coul
            ws.cell(r, 3).border = bord_d
            ws.cell(r, 3).alignment = Alignment(horizontal="right")
            ws.cell(r, 3).font = Font(size=10)

        # --- Camembert : répartition des postes (par année) ---
        PIE_COL = 27
        L_PIE = get_column_letter(PIE_COL)
        ws.cell(1, PIE_COL, "Poste")
        ws.cell(1, PIE_COL + 1, "Montant")
        for idx, poste in enumerate(postes_list):
            r = 2 + idx
            ws.cell(r, PIE_COL, poste)
            ws.cell(r, PIE_COL + 1).value = (
                f"=IFERROR(SUMIFS('{sheet_dep}'!${L_montant}:${L_montant},"
                f"'{sheet_dep}'!${L_annee}:${L_annee},$B$2,"
                f"'{sheet_dep}'!${L_poste}:${L_poste},$" + L_PIE + f"${r}),0)"
            )
            ws.cell(r, PIE_COL + 1).number_format = "#,##0"
        n_post = len(postes_list)

        pie = PieChart()
        pie.title = "Répartition des postes"
        pie.height = 11
        pie.width = 14
        pie.style = 10
        pie.add_data(
            Reference(ws, min_col=PIE_COL + 1, min_row=1, max_row=1 + n_post),
            titles_from_data=True,
        )
        pie.set_categories(
            Reference(ws, min_col=PIE_COL, min_row=2, max_row=1 + n_post)
        )
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        ws.add_chart(pie, "E8")

        ws.column_dimensions["A"].width = 48
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions[L_PIE].hidden = True
        ws.column_dimensions[get_column_letter(PIE_COL + 1)].hidden = True