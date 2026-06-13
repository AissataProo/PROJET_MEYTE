"""Point d'entrée principal du Dashboard AMeli (fichier fusionné)."""

from config import OUTPUT_DASHBOARD, SHEET_CLEANED_DEPENSES, SHEET_CLEANED_EFFECTIFS
from data import get_cleaned_effectifs, get_cleaned_depenses, load_data , compute_list_lengths
from sheets.filters import OngletFiltres
from sheets.graphiques import OngletGraphiques, OngletGraphiquesEffectifs
from sheets.depenses import OngletDepenses
from sheets.region import OngletRegion
from sheets.departement import OngletDepartement
from sheets.postes import OngletPostes


def main():
    print("Dashboard Ameli (Fichier Fusionné)")
    print("=" * 70)

    try:
        print("\nChargement des donnees...")
        df_effectifs = get_cleaned_effectifs()
        df_depenses = get_cleaned_depenses()

        print(f"   Effectifs: {len(df_effectifs)} lignes")
        print(f"   Depenses: {len(df_depenses)} lignes")

        print(f"\nCreation de {OUTPUT_DASHBOARD.name}...")
        excel = ExcelManager(OUTPUT_DASHBOARD)

        print("   - cleanedData_Depenses...")
        ws_dep = excel.create_sheet(SHEET_CLEANED_DEPENSES, 0)
        excel.write_dataframe(ws_dep, df_depenses)
        excel.auto_adjust_columns(ws_dep)

        print("   - Filtres (Dépenses)...")
        filtres_dep = OngletFiltres(excel, df_depenses)
        filtres_dep.create_depenses()

        print("   - Depenses...")
        depenses = OngletDepenses(excel, df_depenses)
        depenses.create()

        print("   - Graphiques (Dépenses)...")
        graphs_dep = OngletGraphiques(excel, df_depenses)
        graphs_dep.create()

        print("   - Postes...")
        postes = OngletPostes(excel, df_depenses)
        postes.create()

        print("   - cleanedData_Effectifs...")
        ws_eff = excel.create_sheet(SHEET_CLEANED_EFFECTIFS, 0)
        excel.write_dataframe(ws_eff, df_effectifs)
        excel.auto_adjust_columns(ws_eff)

        print("   - Filtres (Effectifs)...")
        filtres_eff = OngletFiltres(excel, df_effectifs)
        filtres_eff.create_effectifs()

        print("   - Region...")
        region = OngletRegion(excel, df_effectifs)
        region.create()

        print("   - Departement...")
        dept = OngletDepartement(excel, df_effectifs)
        dept.create()

        print("   - Graphiques (Effectifs)...")
        graphs_eff = OngletGraphiquesEffectifs(excel, df_effectifs)
        graphs_eff.create()

        excel.save()

        print(f"\n✓ Fichier créé : {OUTPUT_DASHBOARD}")
        return 0

    except Exception as e:
        print(f"\nErreur: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main())