"""Onglet Postes : tableau de bord depenses par poste avec filtres et graphique."""

from openpyxl.chart import BarChart, Reference
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
        col_poste = headers.get("poste de dépense", 5)
        col_montant = headers.get("montant", 7)
        col_effectif = headers.get("effectif", 8)

        col_annee_letter = get_column_letter(col_annee)
        col_poste_letter = get_column_letter(col_poste)
        col_montant_letter = get_column_letter(col_montant)
        col_effectif_letter = get_column_letter(col_effectif)

        annees_list = sorted(self.df["annee"].dropna().unique(), reverse=True)
        annee_sel = annees_list[0] if len(annees_list) else None

        postes_base = sorted(
            {
                str(v).strip()
                for v in self.df["poste de dépense"].dropna().unique()
                if v and "total" not in str(v).lower()
            }
        )

        postes_with_amount = []
        for poste in postes_base:
            montant = 0
            for row in ws_dep.iter_rows(
                min_row=2, max_row=ws_dep.max_row, values_only=True
            ):
                if (
                    row[col_poste - 1]
                    and str(row[col_poste - 1]).strip() == poste
                    and str(row[col_annee - 1]).strip() == str(annee_sel)
                ):
                    try:
                        montant += float(row[col_montant - 1] or 0)
                    except Exception:
                        pass
            postes_with_amount.append((poste, montant))

        postes_with_amount.sort(key=lambda x: x[1], reverse=True)
        postes_list = [p for p, _ in postes_with_amount]

        COLOR_VERT = "FF006B4F"
        COLOR_BLEU = "FF4472C4"

        ws.merge_cells("A1:H1")
        ws["A1"] = "Depenses par Poste"
        ws["A1"].font = Font(bold=True, size=13, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color=COLOR_VERT, fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        ws["B3"] = "Annee"
        ws["B3"].font = Font(bold=True, color="FFFFFF")
        ws["B3"].fill = PatternFill(start_color=COLOR_BLEU, fill_type="solid")
        ws["B3"].alignment = Alignment(horizontal="center")

        ws["C3"] = annees_list[0] if annees_list else None
        ws["C3"].fill = PatternFill(start_color="FFFFFFFF", fill_type="solid")
        ws["C3"].font = Font(bold=True, size=11)
        ws["C3"].alignment = Alignment(horizontal="center")

        annees_str = ",".join(str(int(a)) for a in annees_list)
        dv_annee = DataValidation(
            type="list", formula1=f'"{annees_str}"', allow_blank=False
        )
        ws.add_data_validation(dv_annee)
        dv_annee.add("C3")

        ws["E3"] = "Poste"
        ws["E3"].font = Font(bold=True, color="FFFFFF")
        ws["E3"].fill = PatternFill(start_color=COLOR_BLEU, fill_type="solid")
        ws["E3"].alignment = Alignment(horizontal="center")

        ws["F3"] = postes_list[0] if postes_list else ""
        ws["F3"].fill = PatternFill(start_color="FFFFFFFF", fill_type="solid")
        ws["F3"].font = Font(bold=True, size=11)
        ws["F3"].alignment = Alignment(horizontal="center")

        postes_str = ",".join(postes_list)
        dv_poste = DataValidation(
            type="list", formula1=f'"{postes_str}"', allow_blank=False
        )
        ws.add_data_validation(dv_poste)
        dv_poste.add("F3")

        TABLE_HDR = 6
        TABLE_START = TABLE_HDR + 1

        border = Border(
            left=Side(style="thin", color="CCCCCC"),
            right=Side(style="thin", color="CCCCCC"),
            top=Side(style="thin", color="CCCCCC"),
            bottom=Side(style="thin", color="CCCCCC"),
        )

        couleur_alt_1 = PatternFill(start_color="FFF5F5F5", fill_type="solid")
        couleur_alt_2 = PatternFill(start_color="FFFFFFFF", fill_type="solid")

        def make_header_cell(cell, label):
            cell.value = label
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color=COLOR_BLEU, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        def style_data_cell(cell, idx=0, num_format=None, align="left"):
            cell.fill = couleur_alt_1 if idx % 2 == 0 else couleur_alt_2
            cell.font = Font(color="000000")
            cell.border = border
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if num_format:
                cell.number_format = num_format

        for col_i, lbl in enumerate(
            ["Poste", "Montant (E)", "Effectif", "Cout/patient"], start=1
        ):
            make_header_cell(ws.cell(TABLE_HDR, col_i), lbl)

        ws.row_dimensions[TABLE_HDR].height = 22

        for idx, poste in enumerate(postes_list):
            r = TABLE_START + idx

            c = ws.cell(r, 1)
            c.value = poste
            style_data_cell(c, idx, align="left")

            c = ws.cell(r, 2)
            c.value = (
                f"=IFERROR(SUMIFS('{sheet_dep}'!${col_montant_letter}:${col_montant_letter},"
                f"'{sheet_dep}'!${col_annee_letter}:${col_annee_letter},$C$3,"
                f"'{sheet_dep}'!${col_poste_letter}:${col_poste_letter},$F$3),0)"
            )
            style_data_cell(c, idx, "#,##0", "right")

            c = ws.cell(r, 3)
            c.value = (
                f"=IFERROR(SUMIFS('{sheet_dep}'!${col_effectif_letter}:${col_effectif_letter},"
                f"'{sheet_dep}'!${col_annee_letter}:${col_annee_letter},$C$3,"
                f"'{sheet_dep}'!${col_poste_letter}:${col_poste_letter},$F$3),0)"
            )
            style_data_cell(c, idx, "#,##0", "right")

            c = ws.cell(r, 4)
            c.value = f"=IFERROR(B{r}/C{r},0)"
            style_data_cell(c, idx, "#,##0.00", "right")

            ws.row_dimensions[r].height = 18

        last_row = TABLE_START + len(postes_list) - 1
        chart = BarChart()
        chart.type = "col"
        chart.title = "Depenses par Poste"
        chart.y_axis.title = "Montant (E)"
        chart.x_axis.title = "Postes"
        chart.height = 12
        chart.width = 20

        data = Reference(ws, min_col=2, min_row=TABLE_HDR, max_row=last_row)
        cats = Reference(ws, min_col=1, min_row=TABLE_START, max_row=last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, "G6")

        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 18
