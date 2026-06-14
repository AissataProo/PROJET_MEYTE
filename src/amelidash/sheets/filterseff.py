from components.filters import add_filter
from config import COULEURS


class OngletFiltreseff:
    """Crée l’onglet caché contenant les filtres dédiés aux effectifs :
    vérifie la présence des colonnes nécessaires, génère les listes de valeurs
    propres et triées pour chaque filtre (pathologie, année, région, département,
    classe d’âge) et applique la mise en forme associée aux zones de sélection."""

    def __init__(self, wb, df):
        self.wb = wb
        self.df = df
        self.color_label = COULEURS["accent"]
        self.color_value = COULEURS["secondaire"]

    def create_effectifs(self):
        # 1. Gestion sécurisée de la feuille
        sheet_name = "Filtres_Effectifs"
        if sheet_name in self.wb.sheetnames:
            ws = self.wb[sheet_name]
        else:
            ws = self.wb.create_sheet(sheet_name)
        ws.sheet_state = "hidden"

        # 2. Vérification des colonnes (Debug)
        required_cols = ["patho_niv1", "annee", "Région", "Département", "Classe d'age"]
        for col in required_cols:
            if col not in self.df.columns:
                raise KeyError(
                    f"La colonne '{col}' est absente. Colonnes trouvées : {list(self.df.columns)}"
                )

        # 3. Création des filtres
        filters = {
            "A1": ("Pathologie", self.df["patho_niv1"]),
            "B1": ("Annee", self.df["annee"]),
            "C1": ("Région", self.df["Région"]),
            "D1": ("Département", self.df["Département"]),
            "E1": ("Classe d'age", self.df["Classe d'age"]),
        }

        for cell, (label, col_data) in filters.items():
            # Nettoyage et tri des valeurs
            values = sorted(col_data.dropna().unique())
            clean_values = [str(v) for v in values]

            add_filter(
                ws,
                cell,
                label,
                clean_values,
                clean_values[0] if clean_values else "",
                self.color_value,
                self.color_label,
            )

        # 4. Largeurs
        widths = {"A": 35, "B": 15, "C": 30, "D": 30, "E": 20}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
