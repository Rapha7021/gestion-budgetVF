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

print('=== DEBUG DÉTAILLÉ DES CALCULS ===')

# Test 1: Montant théorique total
montant_theorique = SubventionDialog._calculate_detailed_subvention_amount(cursor, projet_id, subvention_data)
print(f'1. Montant théorique total: {montant_theorique:.2f}€')

# Test 2: Dépenses éligibles totales (période complète de subvention)
import datetime
debut_subv = datetime.datetime.strptime('04/2026', '%m/%Y')
fin_subv = datetime.datetime.strptime('03/2027', '%m/%Y')

depenses_totales = SubventionDialog._calculate_total_eligible_expenses(cursor, projet_id, subvention_data, debut_subv, fin_subv)
print(f'2. Dépenses éligibles totales (période subvention): {depenses_totales:.2f}€')

# Test 3: Dépenses éligibles par année
depenses_2026 = SubventionDialog._calculate_period_eligible_expenses(cursor, projet_id, subvention_data, 2026, None)
print(f'3a. Dépenses éligibles 2026: {depenses_2026:.2f}€')

depenses_2027 = SubventionDialog._calculate_period_eligible_expenses(cursor, projet_id, subvention_data, 2027, None)
print(f'3b. Dépenses éligibles 2027: {depenses_2027:.2f}€')

# Test 4: Calcul des proportions
if depenses_totales > 0:
    proportion_2026 = depenses_2026 / depenses_totales
    proportion_2027 = depenses_2027 / depenses_totales
    print(f'4a. Proportion 2026: {proportion_2026:.4f} ({proportion_2026*100:.2f}%)')
    print(f'4b. Proportion 2027: {proportion_2027:.4f} ({proportion_2027*100:.2f}%)')
    
    # Test 5: Montants répartis
    montant_2026 = montant_theorique * proportion_2026
    montant_2027 = montant_theorique * proportion_2027
    print(f'5a. Montant réparti 2026: {montant_2026:.2f}€')
    print(f'5b. Montant réparti 2027: {montant_2027:.2f}€')
    print(f'5c. Total réparti: {montant_2026 + montant_2027:.2f}€')

# Test 6: Calcul des dépenses brutes par année pour comprendre
print(f'\n=== DÉPENSES BRUTES PAR ANNÉE ===')

# Temps de travail 2026
cursor.execute('''SELECT SUM(tt.jours * cc.montant_charge) 
                  FROM temps_travail tt
                  JOIN categorie_cout cc ON cc.libelle = tt.categorie AND cc.annee = tt.annee
                  WHERE tt.projet_id = ? AND tt.annee = ?''', (projet_id, 2026))
temps_brut_2026 = cursor.fetchone()[0] or 0

# Temps de travail 2027  
cursor.execute('''SELECT SUM(tt.jours * cc.montant_charge) 
                  FROM temps_travail tt
                  JOIN categorie_cout cc ON cc.libelle = tt.categorie AND cc.annee = tt.annee
                  WHERE tt.projet_id = ? AND tt.annee = ?''', (projet_id, 2027))
temps_brut_2027 = cursor.fetchone()[0] or 0

print(f'Temps travail brut 2026: {temps_brut_2026:.2f}€')
print(f'Temps travail brut 2027: {temps_brut_2027:.2f}€')
print(f'Total temps brut: {temps_brut_2026 + temps_brut_2027:.2f}€')

# Avec coefficients appliqués correctement
temps_eligible_2026 = temps_brut_2026 * subvention_data['coef_temps_travail'] * subvention_data['cd']
temps_eligible_2027 = temps_brut_2027 * subvention_data['coef_temps_travail'] * subvention_data['cd']

print(f'Temps travail éligible 2026 (coef {subvention_data["coef_temps_travail"]} × CD {subvention_data["cd"]}): {temps_eligible_2026:.2f}€')
print(f'Temps travail éligible 2027 (coef {subvention_data["coef_temps_travail"]} × CD {subvention_data["cd"]}): {temps_eligible_2027:.2f}€')

conn.close()
