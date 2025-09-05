import sqlite3
from subvention_dialog import SubventionDialog
import datetime

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

print('=== ANALYSE MOIS PAR MOIS ===')

# Analyser la période de subvention mois par mois
debut_subv = datetime.datetime.strptime('04/2026', '%m/%Y')
fin_subv = datetime.datetime.strptime('03/2027', '%m/%Y')

current_date = debut_subv.replace(day=1)
total_period = 0

print('Calcul dépenses éligibles par mois (période subvention):')
while current_date <= fin_subv:
    montant_mois = SubventionDialog._calculate_period_eligible_expenses(
        cursor, projet_id, subvention_data, current_date.year, current_date.month
    )
    total_period += montant_mois
    print(f'  {current_date.month:02d}/{current_date.year}: {montant_mois:.2f}€')
    
    # Passer au mois suivant
    if current_date.month == 12:
        current_date = current_date.replace(year=current_date.year + 1, month=1)
    else:
        current_date = current_date.replace(month=current_date.month + 1)

print(f'Total période subvention: {total_period:.2f}€')

print('\n=== COMPARAISON AVEC CALCULS ANNUELS ===')

# Comparaison avec les calculs annuels
depenses_2026_annee = SubventionDialog._calculate_period_eligible_expenses(cursor, projet_id, subvention_data, 2026, None)
depenses_2027_annee = SubventionDialog._calculate_period_eligible_expenses(cursor, projet_id, subvention_data, 2027, None)

print(f'Dépenses 2026 (année complète): {depenses_2026_annee:.2f}€')
print(f'Dépenses 2027 (année complète): {depenses_2027_annee:.2f}€')
print(f'Total années: {depenses_2026_annee + depenses_2027_annee:.2f}€')

print('\n=== VÉRIFICATION DES DONNÉES BRUTES ===')

# Vérifier qu'il y a bien des données pour chaque mois de la période subvention
print('Vérification des temps de travail par mois:')
cursor.execute('''SELECT annee, mois, SUM(jours) as total_jours
                  FROM temps_travail 
                  WHERE projet_id = ? AND annee IN (2026, 2027)
                  GROUP BY annee, mois
                  ORDER BY annee, mois''', (projet_id,))

for annee, mois, jours in cursor.fetchall():
    print(f'  {mois} {annee}: {jours} jours')

print('\nVérification des dépenses externes par mois:')
cursor.execute('''SELECT annee, mois, SUM(montant) as total_montant
                  FROM depenses 
                  WHERE projet_id = ? AND annee IN (2026, 2027)
                  GROUP BY annee, mois
                  ORDER BY annee, mois''', (projet_id,))

for annee, mois, montant in cursor.fetchall():
    print(f'  {mois} {annee}: {montant}€')

conn.close()
