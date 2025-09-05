import sqlite3
from subvention_dialog import SubventionDialog

conn = sqlite3.connect('gestion_budget.db')
cursor = conn.cursor()

# Données du projet TASCII
projet_id = 3

subvention_data = {
    'nom': 'AEAG',
    'mode_simplifie': 0,
    'montant_forfaitaire': 0,
    'depenses_temps_travail': 1,
    'coef_temps_travail': 0.48,
    'depenses_externes': 1,
    'coef_externes': 0.49,
    'depenses_autres_achats': 1,
    'coef_autres_achats': 0.49,
    'depenses_dotation_amortissements': 0,
    'coef_dotation_amortissements': 1.0,
    'cd': 1.54,
    'taux': 50.0,
    'date_debut_subvention': '04/2026',
    'date_fin_subvention': '03/2027'
}

print('=== VÉRIFICATION ANNÉE 2025 ===')

# Test pour 2025 (qui devrait être pris en compte pour le total)
depenses_2025 = SubventionDialog._calculate_period_eligible_expenses(cursor, projet_id, subvention_data, 2025, None)
print(f'Dépenses éligibles 2025: {depenses_2025:.2f}€')

depenses_2026 = SubventionDialog._calculate_period_eligible_expenses(cursor, projet_id, subvention_data, 2026, None)
print(f'Dépenses éligibles 2026: {depenses_2026:.2f}€')

depenses_2027 = SubventionDialog._calculate_period_eligible_expenses(cursor, projet_id, subvention_data, 2027, None)
print(f'Dépenses éligibles 2027: {depenses_2027:.2f}€')

total_manuel = depenses_2025 + depenses_2026 + depenses_2027
print(f'Total dépenses éligibles (calcul manuel): {total_manuel:.2f}€')

# Comparaison avec le calcul automatique
import datetime
debut_subv = datetime.datetime.strptime('04/2026', '%m/%Y')
fin_subv = datetime.datetime.strptime('03/2027', '%m/%Y')

depenses_totales_auto = SubventionDialog._calculate_total_eligible_expenses(cursor, projet_id, subvention_data, debut_subv, fin_subv)
print(f'Total dépenses éligibles (calcul auto projet complet): {depenses_totales_auto:.2f}€')

# Calcul de répartition avec toutes les années
if total_manuel > 0:
    proportion_2025 = depenses_2025 / total_manuel
    proportion_2026 = depenses_2026 / total_manuel 
    proportion_2027 = depenses_2027 / total_manuel
    
    print(f'\\nProportions correctes:')
    print(f'  2025: {proportion_2025:.4f} ({proportion_2025*100:.2f}%)')
    print(f'  2026: {proportion_2026:.4f} ({proportion_2026*100:.2f}%)')
    print(f'  2027: {proportion_2027:.4f} ({proportion_2027*100:.2f}%)')
    
    # Montant théorique
    montant_theorique = SubventionDialog._calculate_detailed_subvention_amount(cursor, projet_id, subvention_data)
    
    # Répartition correcte
    montant_2025 = montant_theorique * proportion_2025
    montant_2026 = montant_theorique * proportion_2026
    montant_2027 = montant_theorique * proportion_2027
    
    print(f'\\nRépartition correcte:')
    print(f'  2025: {montant_2025:.2f}€')
    print(f'  2026: {montant_2026:.2f}€')
    print(f'  2027: {montant_2027:.2f}€')
    print(f'  Total: {montant_2025 + montant_2026 + montant_2027:.2f}€')
    
    # Mais la subvention ne s'applique que sur 2026-2027 !
    print(f'\\nSubvention RÉELLE (seulement période subvention 2026-2027):')
    print(f'  2026: {montant_2026:.2f}€')
    print(f'  2027: {montant_2027:.2f}€')
    print(f'  Total subvention: {montant_2026 + montant_2027:.2f}€')

conn.close()
