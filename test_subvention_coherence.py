#!/usr/bin/env python3
"""
Script de test pour vérifier la cohérence des calculs de subvention
entre SubventionDialog, ProjectDetailsDialog et CompteResultatDisplay
"""

import sqlite3
import sys
import os

# Ajouter le répertoire du projet au path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from subvention_dialog import SubventionDialog

DB_PATH = 'gestion_budget.db'

def test_subvention_consistency():
    """Test de cohérence des calculs de subvention"""
    
    print("=== TEST DE COHÉRENCE DES CALCULS DE SUBVENTION ===")
    print()
    
    # Données du test
    projet_id = 9  # TESTE
    year = 2024
    month = 9  # Septembre 2024
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Récupérer les paramètres de la subvention
    cursor.execute('''
        SELECT nom, mode_simplifie, montant_forfaitaire, depenses_temps_travail, coef_temps_travail, 
               depenses_externes, coef_externes, depenses_autres_achats, coef_autres_achats, 
               depenses_dotation_amortissements, coef_dotation_amortissements, cd, taux,
               date_debut_subvention, date_fin_subvention, montant_subvention_max, depenses_eligibles_max
        FROM subventions WHERE projet_id = ?
    ''', (projet_id,))
    
    subvention_row = cursor.fetchone()
    if not subvention_row:
        print("Aucune subvention trouvée pour le projet TESTE")
        return
    
    # Construire le dictionnaire de données
    subvention_data = {
        'nom': subvention_row[0],
        'mode_simplifie': subvention_row[1] or 0,
        'montant_forfaitaire': subvention_row[2] or 0,
        'depenses_temps_travail': subvention_row[3] or 0,
        'coef_temps_travail': subvention_row[4] or 1,
        'depenses_externes': subvention_row[5] or 0,
        'coef_externes': subvention_row[6] or 1,
        'depenses_autres_achats': subvention_row[7] or 0,
        'coef_autres_achats': subvention_row[8] or 1,
        'depenses_dotation_amortissements': subvention_row[9] or 0,
        'coef_dotation_amortissements': subvention_row[10] or 1,
        'cd': subvention_row[11] or 1,
        'taux': subvention_row[12] or 100,
        'date_debut_subvention': subvention_row[13],
        'date_fin_subvention': subvention_row[14],
        'montant_subvention_max': subvention_row[15],
        'depenses_eligibles_max': subvention_row[16]
    }
    
    print(f"Subvention: {subvention_data['nom']}")
    print(f"Période: {subvention_data['date_debut_subvention']} - {subvention_data['date_fin_subvention']}")
    print(f"Mode: {'Simplifié' if subvention_data['mode_simplifie'] else 'Détaillé'}")
    print(f"Coefficients: temps={subvention_data['coef_temps_travail']}, cd={subvention_data['cd']}, taux={subvention_data['taux']}%")
    print()
    
    # Test 1: Calcul de l'assiette éligible totale
    print("1. TEST ASSIETTE ÉLIGIBLE TOTALE")
    assiette_totale = SubventionDialog._calculate_total_eligible_expenses(
        cursor, projet_id, subvention_data, 
        subvention_data['date_debut_subvention'], 
        subvention_data['date_fin_subvention']
    )
    print(f"   Assiette éligible totale: {assiette_totale:,.2f} €")
    
    # Test 2: Calcul mensuel (Septembre 2024)
    print("\n2. TEST CALCUL MENSUEL (Septembre 2024)")
    montant_septembre = SubventionDialog.calculate_distributed_subvention(
        projet_id, subvention_data, year, month
    )
    print(f"   Montant septembre 2024: {montant_septembre:,.2f} €")
    
    # Test 3: Calcul annuel (2024)
    print("\n3. TEST CALCUL ANNUEL (2024)")
    montant_2024 = SubventionDialog.calculate_distributed_subvention(
        projet_id, subvention_data, year, None
    )
    print(f"   Montant année 2024: {montant_2024:,.2f} €")
    
    # Test 4: Calcul annuel (2025)
    print("\n4. TEST CALCUL ANNUEL (2025)")
    montant_2025 = SubventionDialog.calculate_distributed_subvention(
        projet_id, subvention_data, 2025, None
    )
    print(f"   Montant année 2025: {montant_2025:,.2f} €")
    
    # Test 5: Somme des années doit être proche du montant total attendu
    print("\n5. TEST COHÉRENCE TOTALE")
    montant_total_calcule = montant_2024 + montant_2025
    montant_total_attendu = assiette_totale * (subvention_data['taux'] / 100.0)
    
    print(f"   Somme 2024 + 2025: {montant_total_calcule:,.2f} €")
    print(f"   Montant attendu (assiette × taux): {montant_total_attendu:,.2f} €")
    print(f"   Différence: {abs(montant_total_calcule - montant_total_attendu):,.2f} €")
    
    # Test 6: Vérification détaillée des périodes éligibles
    print("\n6. DÉTAIL DES PÉRIODES ÉLIGIBLES")
    periodes_test = [
        (2024, 1, "Janvier 2024"),  # Avant période subvention
        (2024, 9, "Septembre 2024"),  # Dans période subvention
        (2025, 1, "Janvier 2025"),   # Dans période subvention  
        (2025, 6, "Juin 2025"),      # Dans période subvention
        (2025, 7, "Juillet 2025"),   # Après période subvention
    ]
    
    for year_t, month_t, nom_periode in periodes_test:
        montant_periode = SubventionDialog.calculate_distributed_subvention(
            projet_id, subvention_data, year_t, month_t
        )
        print(f"   {nom_periode}: {montant_periode:,.2f} €")
    
    conn.close()
    
    print("\n=== FIN DU TEST ===")

if __name__ == "__main__":
    test_subvention_consistency()