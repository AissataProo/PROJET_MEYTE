from components.filters import add_filter
from config import COULEURS


class OngletFiltres:
    """Construit l’onglet de synthèse des dépenses à partir des données nettoyées :
    calcule le total des montants, détaille les dépenses par sous‑pathologie (niveau 2),
    compare les années disponibles et génère un graphique camembert illustrant la
    répartition des dépenses."""

    def __init__(self, wb, df):
        self.wb = wb
        self.df = df
        self.color_label = COULEURS["accent"]
        self.color_value = COULEURS["secondaire"]

    def create_depenses(self):
        # La feuille est bien créée
        sheet_name = "Filtres_Depenses"
        if sheet_name in self.wb.sheetnames:
            ws = self.wb[sheet_name]
        else:
            ws = self.wb.create_sheet(sheet_name)
        ws.sheet_state = "hidden"

        # TOUT LE RESTE DOIT ÊTRE DÉCALÉ VERS LA DROITE
        required_cols = ["poste de dépense", "annee", "patho_niv1", "sous poste"]
        for col in required_cols:
            if col not in self.df.columns:
                raise KeyError(f"La colonne '{col}' est absente pour les dépenses.")

        postes = sorted([str(v) for v in self.df["poste de dépense"].dropna().unique()])
        annees = sorted(
            [str(int(a)) for a in self.df["annee"].dropna().unique()], reverse=True
        )
        pathos = sorted([str(v) for v in self.df["patho_niv1"].dropna().unique()])
        sous_postes = sorted([str(v) for v in self.df["sous poste"].dropna().unique()])

        filters = [
            ("A1", "Poste", postes),
            ("B1", "Annee", annees),
            ("C1", "Pathologie", pathos),
            ("D1", "Sous-poste", sous_postes),
        ]

        for cell, label, values in filters:
            add_filter(
                ws,
                cell,
                label,
                values,
                values[0] if values else "",
                self.color_value,
                self.color_label,
            )

        widths = {"A": 35, "B": 15, "C": 35, "D": 35}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
