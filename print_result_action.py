
import sqlite3
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QPushButton

DB_PATH = 'gestion_budget.db'

def print_result_action(parent, projet_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Récupérer toutes les années du projet
    cursor.execute('SELECT DISTINCT annee FROM temps_travail WHERE projet_id=? ORDER BY annee', (projet_id,))
    annees = [row[0] for row in cursor.fetchall()]
    if not annees:
        conn.close()
        dlg = QDialog(parent)
        dlg.setWindowTitle("Synthèse budget")
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Aucune donnée disponible pour ce projet."))
        dlg.setLayout(layout)
        dlg.exec()
        return

    dlg = QDialog(parent)
    dlg.setWindowTitle(f"Synthèse budget - Projet {projet_id}")
    layout = QVBoxLayout()
    layout.addWidget(QLabel(f"Projet ID: {projet_id}"))

    for annee in annees:
        layout.addWidget(QLabel(f"Année: {annee}"))

        # Récupérer les catégories et directions
        cursor.execute('SELECT DISTINCT direction FROM temps_travail WHERE projet_id=? AND annee=?', (projet_id, annee))
        directions = [row[0] for row in cursor.fetchall()]
        cursor.execute('SELECT DISTINCT categorie FROM temps_travail WHERE projet_id=? AND annee=?', (projet_id, annee))
        categories = [row[0] for row in cursor.fetchall()]

        # Récupérer les jours par direction/catégorie
        cursor.execute('SELECT direction, categorie, SUM(jours) FROM temps_travail WHERE projet_id=? AND annee=? GROUP BY direction, categorie', (projet_id, annee))
        jours_data = cursor.fetchall()

        # Récupérer les jours totaux par direction
        cursor.execute('SELECT direction, SUM(jours) FROM temps_travail WHERE projet_id=? AND annee=? GROUP BY direction', (projet_id, annee))
        jours_direction = {row[0]: row[1] for row in cursor.fetchall()}

        # Récupérer le nombre total de jours
        cursor.execute('SELECT SUM(jours) FROM temps_travail WHERE projet_id=? AND annee=?', (projet_id, annee))
        total_jours = cursor.fetchone()[0] or 0

        # Récupérer les coûts par catégorie
        cursor.execute('SELECT categorie, montant_charge, cout_production, cout_complet FROM categorie_cout WHERE annee=?', (annee,))
        couts = {row[0]: row[1:] for row in cursor.fetchall()}

        # Préparer les lignes du tableau avec calcul des totaux
        table_data = []
        for direction, categorie, jours in jours_data:
            montant_charge, cout_direct, cout_complet = couts.get(categorie, (None, None, None))
            montant_charge_total = jours * montant_charge if montant_charge is not None else ''
            cout_direct_total = jours * cout_direct if cout_direct is not None else ''
            cout_complet_total = jours * cout_complet if cout_complet is not None else ''
            table_data.append([
                f"{categorie} / {direction}",
                f"{jours:.2f}",
                f"{montant_charge_total:.2f}" if montant_charge_total != '' else '',
                f"{cout_direct_total:.2f}" if cout_direct_total != '' else '',
                f"{cout_complet_total:.2f}" if cout_complet_total != '' else ''
            ])

        # Ajouter les totaux par direction
        for direction in directions:
            table_data.append([
                f"Total {direction}",
                f"{jours_direction[direction]:.2f}", '', '', ''
            ])

        # Ajouter le total général
        table_data.append(["Total général", f"{total_jours:.2f}", '', '', ''])

        # Récupérer les dépenses et recettes
        cursor.execute('SELECT SUM(montant) FROM depenses WHERE projet_id=? AND annee=?', (projet_id, annee))
        total_depenses = cursor.fetchone()[0] or 0
        cursor.execute('SELECT SUM(montant) FROM recettes WHERE projet_id=? AND annee=?', (projet_id, annee))
        total_recettes = cursor.fetchone()[0] or 0

        # Détail des dépenses
        cursor.execute('SELECT mois, montant, detail FROM depenses WHERE projet_id=? AND annee=?', (projet_id, annee))
        depenses_details = cursor.fetchall()
        # Détail des recettes
        cursor.execute('SELECT mois, montant, detail FROM recettes WHERE projet_id=? AND annee=?', (projet_id, annee))
        recettes_details = cursor.fetchall()

        # Ajout des dépenses dans le tableau
        table_data.append(["Dépenses", '', '', '', ''])
        table_data.append(["Total dépenses", f"{total_depenses:.2f} €", '', '', ''])
        for mois, montant, detail in depenses_details:
            table_data.append([f"{mois}", f"{montant:.2f} €", detail, '', ''])

        # Ajout des recettes dans le tableau
        table_data.append(["Recettes", '', '', '', ''])
        table_data.append(["Total recettes", f"{total_recettes:.2f} €", '', '', ''])
        for mois, montant, detail in recettes_details:
            table_data.append([f"{mois}", f"{montant:.2f} €", detail, '', ''])

        # Tableau pour l'année
        table = QTableWidget(len(table_data), 5)
        table.setHorizontalHeaderLabels([
            "Catégorie/Direction", "Nombre jours", "Montant chargé", "Coût direct", "Coût complet"
        ])
        for i, row in enumerate(table_data):
            for j, val in enumerate(row):
                table.setItem(i, j, QTableWidgetItem(str(val)))
        layout.addWidget(table)

    conn.close()

    btn = QPushButton("Fermer")
    btn.clicked.connect(dlg.accept)
    layout.addWidget(btn)
    dlg.setLayout(layout)
    dlg.exec()
