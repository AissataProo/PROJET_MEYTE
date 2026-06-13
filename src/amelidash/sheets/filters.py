# """Onglet Filtres : création des filtres déroulants pour dépenses et effectifs."""

from components.filters import add_filter
from config import COULEURS


class OngletFiltres:

    def __init__(self, excel_manager, df):
        self.excel_manager = excel_manager
        self.df = df
        # Utiliser les couleurs de config.py
        self.color_label = COULEURS["accent"]      # "FF4CAF50"
        self.color_value = COULEURS["secondaire"]  # "FF008080"

    def create_depenses(self):
        ws = self.excel_manager.create_sheet("Filtres", 0)
        ws.sheet_state = "hidden"

        postes = sorted(self.df["poste de dépense"].dropna().unique())
        annees = sorted(self.df["annee"].dropna().unique(), reverse=True)
        pathos = sorted(self.df["patho_niv1"].dropna().unique())
        sous_postes = sorted(self.df["sous poste"].dropna().unique())
        annees_str = [str(int(a)) for a in annees]

        add_filter(
            ws,
            "A1",
            "Poste",
            list(postes),
            postes[0] if postes else "",
            self.color_value,
            self.color_label,
        )
        add_filter(
            ws,
            "B1",
            "Annee",
            annees_str,
            annees_str[0] if annees_str else "",
            self.color_value,
            self.color_label,
        )
        add_filter(
            ws,
            "C1",
            "Pathologie",
            list(pathos),
            pathos[0] if pathos else "",
            self.color_value,
            self.color_label,
        )
        add_filter(
            ws,
            "D1",
            "Sous-poste",
            list(sous_postes),
            sous_postes[0] if sous_postes else "",
            self.color_value,
            self.color_label,
        )

        widths = {"A": 35, "B": 15, "C": 35, "D": 35}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

    def create_effectifs(self):
        ws = self.excel_manager.create_sheet("Filtres", 0)
        ws.sheet_state = "hidden"

        # Dictionnaire de configuration
        filters = {
            "A1": ("Pathologie", sorted(self.df["patho_niv1"].dropna().unique())),
            "B1": (
                "Annee",
                [
                    str(int(a))
                    for a in sorted(self.df["annee"].dropna().unique(), reverse=True)
                ],
            ),
            "C1": ("Région", sorted(self.df["Région"].dropna().unique())),  # avec accent
            "D1": ("Département", sorted(self.df["Département"].dropna().unique())),  # avec accent
            "E1": ("Classe d'age", sorted(self.df["Classe d'age"].dropna().unique())),
            "F1": ("Sexe", sorted(self.df["Sexe"].dropna().unique())),
        }

        for cell, (label, values) in filters.items():
            add_filter(
                ws,
                cell,
                label,
                list(values),
                values[0] if values else "",
                self.color_value,
                self.color_label,
            )

        widths = {"A": 35, "B": 15, "C": 30, "D": 35, "E": 20, "F": 15}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width