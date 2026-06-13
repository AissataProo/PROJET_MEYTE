"""Point d'entree principal du Dashboard Ameli (sans ExcelManager)."""

from openpyxl import Workbook

from config import OUTPUT_DASHBOARD, SHEET_CLEANED_DEPENSES, SHEET_CLEANED_EFFECTIFS
from sheets.graphiques import OngletGraphiques, OngletGraphiquesEffectifs
from data import get_cleaned_effectifs, get_cleaned_depenses
from sheets.filters import OngletFiltres
from sheets.depenses import OngletDepenses
from sheets.region import OngletRegion
from sheets.departement import OngletDepartement
from sheets.postes import OngletPostes


def main():
    print("Dashboard Ameli")
    print("=" * 70)

    try:
        print("\nChargement des donnees...")
        df_effectifs = get_cleaned_effectifs()
        df_depenses = get_cleaned_depenses()

        print(f"   Effectifs: {len(df_effectifs)} lignes")
        print(f"   Depenses: {len(df_depenses)} lignes")

        print(f"\nCreation de {OUTPUT_DASHBOARD.name}...")
        wb = Workbook()
        wb.remove(wb.active)

        print("   - cleanedData_depenses...")
        ws_dep = wb.create_sheet(SHEET_CLEANED_DEPENSES, 0)
        for row in [df_depenses.columns.tolist()] + df_depenses.values.tolist():
            ws_dep.append(row)

        print("   - Filtres...")
        OngletFiltres(wb, df_depenses).create_depenses()

        print("   - Depenses...")
        OngletDepenses(wb, df_depenses).create()

        print("   - Graphiques...")
        OngletGraphiques(wb, df_depenses).create()

        print("   - Postes...")
        OngletPostes(wb, df_depenses)

        print("   - cleanedData_effectifs...")
        ws_eff = wb.create_sheet(SHEET_CLEANED_EFFECTIFS, 0)
        for row in [df_effectifs.columns.tolist()] + df_effectifs.values.tolist():
            ws_eff.append(row)

        print("   - Filtres effectifs...")
        OngletFiltres(wb, df_effectifs).create_effectifs()

        print("   - Region...")
        OngletRegion(wb, df_effectifs).create()

        print("   - Departement...")
        OngletDepartement(wb, df_effectifs).create()

        print("   - Graphiques effectifs...")
        OngletGraphiquesEffectifs(wb, df_effectifs).create()

        wb.save(OUTPUT_DASHBOARD)

        print(f"\nSucces - Fichier cree : {OUTPUT_DASHBOARD}")
        return 0

    except Exception as e:
        print(f"\nErreur: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main())