from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from components.filters import add_filter


class OngletDepenses:
    def __init__(self, wb, df):
        self.wb = wb
        self.df = df

    def create(self):
        ws = self.wb.create_sheet("Depenses", 1)
        ws.sheet_view.showGridLines = False

        ws_dep = self.wb["cleanedData_Depenses"]
        sheet_dep = "cleanedData_Depenses"

        headers = {}
        for idx, cell in enumerate(ws_dep[1], start=1):
            if cell.value is not None:
                headers[str(cell.value).strip()] = idx

        col_annee = headers.get("annee", 1)
        col_poste = headers.get("poste de dépense", 5)
        col_patho3 = headers.get("patho_niv3", 4)
        col_montant = headers.get("montant", 7)
        col_effectif = headers.get("Effectif", 8)

        col_annee_letter = get_column_letter(col_annee)
        col_poste_letter = get_column_letter(col_poste)
        col_patho3_letter = get_column_letter(col_patho3)
        col_montant_letter = get_column_letter(col_montant)
        col_effectif_letter = get_column_letter(col_effectif)

        annees_list = sorted(self.df["annee"].dropna().unique(), reverse=True)
        if not annees_list:
            annees_list = [2023]

        postes_list = sorted(
            {
                str(v).strip()
                for v in self.df["poste de dépense"].dropna().unique()
                if "total" not in str(v).lower()
                and str(v).strip().lower() != "dépenses"
            }
        )
        if not postes_list:
            postes_list = ["Poste 1"]

        poste_defaut = postes_list[0]
        annee_defaut = int(annees_list[0])

        COLOR_VERT = "FF006B4F"
        COLOR_BLEU = "FF4472C4"

        ws.merge_cells("B1:L1")
        ws["B1"] = "Analyse des Pathologies"
        ws["B1"].font = Font(bold=True, size=14, color="FFFFFF")
        ws["B1"].fill = PatternFill(start_color=COLOR_VERT, fill_type="solid")
        ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        add_filter(ws, "B2", "Poste", postes_list, poste_defaut, COLOR_BLEU, COLOR_VERT)
        add_filter(
            ws,
            "D2",
            "Annee",
            [str(int(a)) for a in annees_list],
            str(annee_defaut),
            COLOR_BLEU,
            COLOR_VERT,
        )

        patho3_rows = []
        seen = set()

        for row in ws_dep.iter_rows(min_row=2, values_only=True):
            val_annee = row[col_annee - 1]
            val_poste = row[col_poste - 1]
            val_patho3 = row[col_patho3 - 1]
            if val_annee and val_poste and val_patho3:
                if (
                    int(val_annee) == annee_defaut
                    and str(val_poste).strip() == str(poste_defaut).strip()
                ):
                    label = str(val_patho3).strip()
                    is_total = "total" in label.lower()
                    if not is_total and label not in seen:
                        seen.add(label)
                        patho3_rows.append(label)

        border_header = Border(
            left=Side(style="thin", color="4472C4"),
            right=Side(style="thin", color="4472C4"),
            top=Side(style="thin", color="4472C4"),
            bottom=Side(style="thin", color="4472C4"),
        )
        border_data = Border(
            left=Side(style="thin", color="CCCCCC"),
            right=Side(style="thin", color="CCCCCC"),
            top=Side(style="thin", color="CCCCCC"),
            bottom=Side(style="thin", color="CCCCCC"),
        )

        couleur_alt_1 = PatternFill(start_color="FFF5F5F5", fill_type="solid")
        couleur_alt_2 = PatternFill(start_color="FFFFFFFF", fill_type="solid")
        fill_green = PatternFill(start_color="FF006B4F", fill_type="solid")
        fill_blue = PatternFill(start_color="FF4472C4", fill_type="solid")

        def style_header(cell, label):
            cell.value = label
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = fill_blue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_header

        def style_data(cell, idx=0, num_format=None, align="left"):
            cell.fill = couleur_alt_1 if idx % 2 == 0 else couleur_alt_2
            cell.font = Font(size=10)
            cell.border = border_data
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if num_format:
                cell.number_format = num_format

        ws.merge_cells("A8:B8")
        ws["A8"] = "Total des dépenses"
        ws["A8"].font = Font(bold=True, size=11, color="FFFFFF")
        ws["A8"].fill = fill_green
        ws["A8"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[8].height = 22

        # TOTAL DÉPENSES
        ws["C8"] = (
            f"=IFERROR(SUMIFS('{sheet_dep}'!${col_montant_letter}$2:${col_montant_letter}$5000,"
            f"'{sheet_dep}'!${col_annee_letter}$2:${col_annee_letter}$5000,E$2,"
            f"'{sheet_dep}'!${col_poste_letter}$2:${col_poste_letter}$5000,C$2),0)"
        )
        ws["C8"].font = Font(bold=True, size=13, color="FF006B4F")
        ws["C8"].number_format = "#,##0 €"
        ws["C8"].alignment = Alignment(horizontal="left")

        entetes = ["Pathologie", "Montant", "Effectif", "Coût/patient"]
        for i, texte in enumerate(entetes, start=1):
            style_header(ws.cell(10, i), texte)
        ws.row_dimensions[10].height = 22

        for idx, patho3 in enumerate(patho3_rows):
            row = 12 + idx
            couleur = couleur_alt_1 if idx % 2 == 0 else couleur_alt_2

            ws[f"A{row}"] = patho3
            ws[f"A{row}"].fill = couleur
            ws[f"A{row}"].border = border_data
            ws[f"A{row}"].alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )
            ws[f"B{row}"] = (
                f"=IFERROR(SUMIFS('{sheet_dep}'!${col_montant_letter}$2:${col_montant_letter}$5000,"
                f"'{sheet_dep}'!${col_annee_letter}$2:${col_annee_letter}$5000,E$2,"
                f"'{sheet_dep}'!${col_poste_letter}$2:${col_poste_letter}$5000,C$2,"
                f"'{sheet_dep}'!${col_patho3_letter}$2:${col_patho3_letter}$5000,A{row}),0)"
            )
            ws[f"B{row}"].number_format = "#,##0"
            ws[f"B{row}"].fill = couleur
            ws[f"B{row}"].border = border_data
            ws[f"B{row}"].alignment = Alignment(horizontal="right", vertical="center")
            ws[f"C{row}"] = (
                f"=IFERROR(SUMIFS('{sheet_dep}'!${col_effectif_letter}$2:${col_effectif_letter}$5000,"
                f"'{sheet_dep}'!${col_annee_letter}$2:${col_annee_letter}$5000,E$2,"
                f"'{sheet_dep}'!${col_poste_letter}$2:${col_poste_letter}$5000,C$2,"
                f"'{sheet_dep}'!${col_patho3_letter}$2:${col_patho3_letter}$5000,A{row}),0)"
            )
            ws[f"C{row}"].number_format = "#,##0"
            ws[f"C{row}"].fill = couleur
            ws[f"C{row}"].border = border_data
            ws[f"C{row}"].alignment = Alignment(horizontal="right", vertical="center")

            # cOÛT/PATIENT
            ws[f"D{row}"] = f"=IF(C{row}=0,0,B{row}/C{row})"
            ws[f"D{row}"].number_format = "#,##0.00 €"
            ws[f"D{row}"].fill = couleur
            ws[f"D{row}"].border = border_data
            ws[f"D{row}"].alignment = Alignment(horizontal="right", vertical="center")

            ws.row_dimensions[row].height = 25

        ws.column_dimensions["A"].width = 60
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 35
        ws.column_dimensions["D"].width = 18