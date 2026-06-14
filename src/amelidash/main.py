from openpyxl import Workbook
from config import OUTPUT_DASHBOARD, SHEET_CLEANED_DEPENSES, SHEET_CLEANED_EFFECTIFS
from data import get_cleaned_effectifs, get_cleaned_depenses
from sheets.filtersdep import OngletFiltres
from sheets.filterseff import OngletFiltreseff
from sheets.PathobySexe import OngletSexePatho
from sheets.graphiques import OngletGraphiques
from sheets.depenses import OngletDepenses
from sheets.region import OngletRegion
from sheets.departement import OngletDepartement
from sheets.postes import OngletPostes
from sheets.Analyse import OngletAnalysesUnifiees


def main():
    print("Dashboard Ameli")
    print("=" * 70)

    try:
        print("\nChargement des donnees...")
        df_effectifs = get_cleaned_effectifs()
        df_depenses = get_cleaned_depenses()

        print(f"   Effectifs: {len(df_effectifs)} lignes (total)")
        print(f"   Depenses: {len(df_depenses)} lignes")

        print(f"\nCreation de {OUTPUT_DASHBOARD.name}...")
        wb = Workbook()
        wb.remove(wb.active)

        print("   - cleanedData_Depenses...")
        ws_dep = wb.create_sheet(SHEET_CLEANED_DEPENSES, 0)
        for row in [df_depenses.columns.tolist()] + df_depenses.values.tolist():
            ws_dep.append(row)

        print("   - cleanedData_Effectifs...")
        ws_eff = wb.create_sheet(SHEET_CLEANED_EFFECTIFS, 1)
        df_eff_limited = df_effectifs.head(5000)
        for row in [df_eff_limited.columns.tolist()] + df_eff_limited.values.tolist():
            ws_eff.append(row)

        print("   - Filtres depenses...")
        OngletFiltres(wb, df_depenses).create_depenses()

        print("   - Filtres effectifs...")
        OngletFiltreseff(wb, df_effectifs).create_effectifs()

        print("   - Depenses...")
        OngletDepenses(wb, df_depenses).create()

        print("   - Graphiques depenses...")
        OngletGraphiques(wb, df_depenses).create()

        print("   - Postes...")
        OngletPostes(wb, df_depenses).create()

        print("   - Region...")
        OngletRegion(wb, df_effectifs).create()

        print("   - Departement...")
        OngletDepartement(wb, df_effectifs).create()

        print("   - Sexe et Pathologie...")
        OngletSexePatho(wb, df_effectifs).create()

        print("   - Analyses unifiees...")
        OngletAnalysesUnifiees(wb, df_depenses, df_effectifs).create()

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
