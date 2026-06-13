from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from components.filters import add_filter


class OngletDepenses:

    def __init__(self, excel_manager, df):
        self.excel_manager = excel_manager
        self.df = df
        self.wb = excel_manager.wb

    def create(self):
        ws = self.excel_manager.create_sheet("Depenses", 1)
        ws.sheet_view.showGridLines = False

        ws_dep = self.wb["cleanedData_depenses"]
        sheet_dep = "cleanedData_depenses"

        headers = {}
        for idx, cell in enumerate(ws_dep[1], start=1):
            if cell.value is not None:
                headers[str(cell.value).strip()] = idx

        col_annee = headers.get("annee", 1)
        col_patho = headers.get("patho_niv1", 2)
        col_poste = headers.get("poste de depense", 5)
        col_montant = headers.get("montant", 7)
        col_effectif = headers.get("Effectif", 8)

        col_annee_letter = get_column_letter(col_annee)
        col_patho_letter = get_column_letter(col_patho)
        col_poste_letter = get_column_letter(col_poste)
        col_montant_letter = get_column_letter(col_montant)
        col_effectif_letter = get_column_letter(col_effectif)

        annees_list = sorted(self.df["annee"].dropna().unique(), reverse=True)
        if not annees_list:
            annees_list = [2023]

        postes_list = sorted(
            set(
                str(v).strip()
                for v in self.df["poste de depense"].dropna().unique()
                if "total" not in str(v).lower()
            )
        )
        if not postes_list:
            postes_list = ["Poste 1"]

        pathos_list = sorted(self.df["patho_niv1"].dropna().unique())
        if not pathos_list:
            pathos_list = ["Pathologie 1"]

        poste_defaut = postes_list[0]
        annee_defaut = int(annees_list[0])

        COLOR_VERT = "FF006B4F"
        COLOR_BLEU = "FF4472C4"

        ws.merge_cells("A1:H1")
        ws["A1"] = "Depenses par pathologie"
        ws["A1"].font = Font(bold=True, size=13, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color=COLOR_VERT, fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        add_filter(ws, "B2", "Poste", postes_list, poste_defaut, COLOR_BLEU, COLOR_VERT)
        add_filter(
            ws,
            "D2",
            "Annee",
            [str(int(a)) for a in annees_list],
            str(int(annee_defaut)),
            COLOR_BLEU,
            COLOR_VERT,
        )

        patho_rows = []
        seen = set()

        for row in ws_dep.iter_rows(min_row=2, values_only=True):
            val_annee = row[col_annee - 1]
            val_poste = row[col_poste - 1]
            val_patho = row[col_patho - 1]

            if val_annee and val_poste and val_patho:
                if int(val_annee) == annee_defaut:
                    if str(val_poste).strip() == str(poste_defaut).strip():
                        label = str(val_patho).strip()
                        is_total = (
                            "total" in label.lower()
                            or "soins courants" in label.lower()
                        )
                        if not is_total and label not in seen:
                            seen.add(label)
                            patho_rows.append(label)

        TABLE_HDR = 4
        TABLE_START = TABLE_HDR + 1

        border = Border(
            left=Side(style="thin", color="CCCCCC"),
            right=Side(style="thin", color="CCCCCC"),
            top=Side(style="thin", color="CCCCCC"),
            bottom=Side(style="thin", color="CCCCCC"),
        )
        border_header = Border(
            left=Side(style="thin", color="4472C4"),
            right=Side(style="thin", color="4472C4"),
            top=Side(style="thin", color="4472C4"),
            bottom=Side(style="thin", color="4472C4"),
        )

        couleur_alt_1 = PatternFill(start_color="FFF5F5F5", fill_type="solid")
        couleur_alt_2 = PatternFill(start_color="FFFFFFFF", fill_type="solid")

        def make_header_cell(cell, label):
            cell.value = label
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color=COLOR_BLEU, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_header

        def style_data_cell(cell, idx=0, num_format=None, align="left"):
            cell.fill = couleur_alt_1 if idx % 2 == 0 else couleur_alt_2
            cell.font = Font(color="000000")
            cell.border = border
            cell.alignment = Alignment(horizontal=align, vertical="center")
            if num_format:
                cell.number_format = num_format

        for col_i, lbl in enumerate(
            ["Pathologie", "Effectifs", "Montant (E)", "Cout/patient"], start=1
        ):
            make_header_cell(ws.cell(TABLE_HDR, col_i), lbl)
        ws.row_dimensions[TABLE_HDR].height = 22

        for idx, label in enumerate(patho_rows):
            r = TABLE_START + idx

            c = ws.cell(r, 1)
            c.value = label
            style_data_cell(c, idx, align="left")

            c = ws.cell(r, 2)
            c.value = (
                f"=IFERROR(SUMIFS('{sheet_dep}'!${col_effectif_letter}:${col_effectif_letter},"
                f"'{sheet_dep}'!${col_patho_letter}:${col_patho_letter},A{r},"
                f"'{sheet_dep}'!${col_poste_letter}:${col_poste_letter},B$3,"
                f"'{sheet_dep}'!${col_annee_letter}:${col_annee_letter},D$3),0)"
            )
            style_data_cell(c, idx, "#,##0", "right")

            c = ws.cell(r, 3)
            c.value = (
                f"=IFERROR(SUMIFS('{sheet_dep}'!${col_montant_letter}:${col_montant_letter},"
                f"'{sheet_dep}'!${col_patho_letter}:${col_patho_letter},A{r},"
                f"'{sheet_dep}'!${col_poste_letter}:${col_poste_letter},B$3,"
                f"'{sheet_dep}'!${col_annee_letter}:${col_annee_letter},D$3),0)"
            )
            style_data_cell(c, idx, "#,##0", "right")

            c = ws.cell(r, 4)
            c.value = f"=IFERROR(C{r}/B{r},0)"
            style_data_cell(c, idx, "#,##0.00", "right")

            ws.row_dimensions[r].height = 18

        ws.column_dimensions["A"].width = 55
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 18
