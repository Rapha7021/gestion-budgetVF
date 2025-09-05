import sqlite3
from subvention_dialog import SubventionDialog

conn = sqlite3.connect('gestion_budget.db')
cursor = conn.cursor()

print('=== Analyse du projet TASCII ===')

# Vérifier les données du projet TASCII (qui correspond probablement à R0049)
cursor.execute('SELECT id, nom, date_debut, date_fin FROM projets WHERE nom LIKE "%TASCII%" OR id = 3')
projet = cursor.fetchone()
if projet:
    projet_id, nom, debut, fin = projet
    print(f'Projet: {nom} (ID: {projet_id})')
    print(f'Période projet: {debut} à {fin}')
    
    # Vérifier les subventions de ce projet  
    cursor.execute('''SELECT nom, date_debut_subvention, date_fin_subvention, mode_simplifie, 
                             montant_forfaitaire, taux, cd,
                             depenses_temps_travail, coef_temps_travail,
                             depenses_externes, coef_externes,
                             depenses_autres_achats, coef_autres_achats,
                             depenses_dotation_amortissements, coef_dotation_amortissements
                      FROM subventions WHERE projet_id = ?''', (projet_id,))
    
    subventions = cursor.fetchall()
    print(f'\nNombre de subventions: {len(subventions)}')
    
    for i, subv in enumerate(subventions):
        nom_subv, debut_subv, fin_subv, mode_simp, montant_forf, taux, cd, dep_temps, coef_temps, dep_ext, coef_ext, dep_autres, coef_autres, dep_amort, coef_amort = subv
        print(f'\nSubvention {i+1}: {nom_subv}')
        print(f'  Période subvention: {debut_subv} à {fin_subv}')
        print(f'  Mode simplifié: {bool(mode_simp)}')
        if mode_simp:
            print(f'  Montant forfaitaire: {montant_forf}€')
        else:
            print(f'  Taux: {taux}%, CD: {cd}')
            print(f'  Temps travail: {bool(dep_temps)} (coef: {coef_temps})')
            print(f'  Dépenses externes: {bool(dep_ext)} (coef: {coef_ext})')
            print(f'  Autres achats: {bool(dep_autres)} (coef: {coef_autres})')
            print(f'  Amortissements: {bool(dep_amort)} (coef: {coef_amort})')
        
        # Préparer les données pour le calcul
        subvention_data = {
            'nom': nom_subv,
            'mode_simplifie': mode_simp,
            'montant_forfaitaire': montant_forf or 0,
            'depenses_temps_travail': dep_temps or 0,
            'coef_temps_travail': coef_temps or 1,
            'depenses_externes': dep_ext or 0,
            'coef_externes': coef_ext or 1,
            'depenses_autres_achats': dep_autres or 0,
            'coef_autres_achats': coef_autres or 1,
            'depenses_dotation_amortissements': dep_amort or 0,
            'coef_dotation_amortissements': coef_amort or 1,
            'cd': cd or 1,
            'taux': taux or 100,
            'date_debut_subvention': debut_subv,
            'date_fin_subvention': fin_subv
        }
        
        print(f'\n  === Test de répartition pour {nom_subv} ===')
        
        # Tester la répartition par année
        for annee in [2025, 2026, 2027, 2028]:
            result = SubventionDialog.calculate_distributed_subvention(projet_id, subvention_data, annee, None)
            print(f'    {annee}: {result:.2f}€')
        
        # Calculer le total
        total_reparti = 0
        for annee in [2025, 2026, 2027, 2028]:
            result = SubventionDialog.calculate_distributed_subvention(projet_id, subvention_data, annee, None)
            total_reparti += result
        
        print(f'    TOTAL RÉPARTI: {total_reparti:.2f}€')
        
        # Comparer avec le montant théorique
        if mode_simp:
            montant_theorique = montant_forf
        else:
            # Calculer le montant théorique en mode détaillé
            montant_theorique = SubventionDialog._calculate_detailed_subvention_amount(cursor, projet_id, subvention_data)
        
        print(f'    MONTANT THÉORIQUE: {montant_theorique:.2f}€')
        print(f'    DIFFÉRENCE: {abs(total_reparti - montant_theorique):.2f}€')
    
    print(f'\n=== Analyse des dépenses par année ===')
    
    # Temps de travail par année
    cursor.execute('''SELECT tt.annee, SUM(tt.jours * cc.montant_charge) as cout_total
                      FROM temps_travail tt
                      JOIN categorie_cout cc ON cc.libelle = tt.categorie AND cc.annee = tt.annee
                      WHERE tt.projet_id = ?
                      GROUP BY tt.annee
                      ORDER BY tt.annee''', (projet_id,))
    
    temps_travail_par_annee = cursor.fetchall()
    print('\nTemps de travail par année:')
    total_temps = 0
    for annee, cout in temps_travail_par_annee:
        print(f'  {annee}: {cout:.2f}€')
        total_temps += cout
    print(f'  Total temps: {total_temps:.2f}€')
    
    # Dépenses externes par année
    cursor.execute('''SELECT annee, SUM(montant) as total_depenses
                      FROM depenses
                      WHERE projet_id = ?
                      GROUP BY annee
                      ORDER BY annee''', (projet_id,))
    
    depenses_par_annee = cursor.fetchall()
    print('\nDépenses externes par année:')
    total_depenses = 0
    for annee, montant in depenses_par_annee:
        print(f'  {annee}: {montant:.2f}€')
        total_depenses += montant
    print(f'  Total dépenses: {total_depenses:.2f}€')
    
    # Autres dépenses par année
    cursor.execute('''SELECT annee, SUM(montant) as total_autres
                      FROM autres_depenses
                      WHERE projet_id = ?
                      GROUP BY annee
                      ORDER BY annee''', (projet_id,))
    
    autres_par_annee = cursor.fetchall()
    print('\nAutres dépenses par année:')
    total_autres = 0
    for annee, montant in autres_par_annee:
        print(f'  {annee}: {montant:.2f}€')
        total_autres += montant
    print(f'  Total autres: {total_autres:.2f}€')
    
    print(f'\nTOTAL GÉNÉRAL DÉPENSES: {total_temps + total_depenses + total_autres:.2f}€')
    
else:
    print('Projet TASCII non trouvé')

conn.close()
