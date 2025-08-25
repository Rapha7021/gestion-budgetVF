from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton, QTableWidget, QTableWidgetItem,
    QComboBox, QHBoxLayout, QStackedWidget, QWidget
)
from PyQt6.QtCore import Qt
import sqlite3
import re

class BudgetEditDialog(QDialog):
    def __init__(self, projet_id, parent=None):
        super().__init__(parent)
        self.projet_id = projet_id
        self.setWindowTitle("Budget du Projet")
        self.setMinimumSize(1100, 700)
        main_layout = QVBoxLayout()

        # --- Sélection de l'année en haut ---
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (self.projet_id,))
        row = cursor.fetchone()
        conn.close()
        annees = set()
        if row:
            for date in row:
                if date:
                    date_str = str(date)
                    # Cherche un groupe de 4 chiffres (année) dans la date
                    match = re.search(r'(\d{4})', date_str)
                    if match:
                        annees.add(match.group(1))
        if not annees:
            annees = {"2024"}  # Valeur par défaut AAAA

        annee_label = QLabel("Année du projet :")
        self.annee_combo = QComboBox()
        self.annee_combo.addItems(sorted(annees))
        annee_layout = QHBoxLayout()
        annee_layout.addWidget(annee_label)
        annee_layout.addWidget(self.annee_combo)

        # --- Boutons en haut ---
        btn_layout = QHBoxLayout()
        self.btn_temps = QPushButton("Temps de travail")
        self.btn_recettes = QPushButton("Recettes")
        self.btn_depenses = QPushButton("Dépenses")
        btn_layout.addWidget(self.btn_temps)
        btn_layout.addWidget(self.btn_recettes)
        btn_layout.addWidget(self.btn_depenses)

        # --- Regroupe année + boutons ---
        top_layout = QHBoxLayout()
        top_layout.addLayout(annee_layout)
        top_layout.addLayout(btn_layout)
        main_layout.addLayout(top_layout)

        # --- QStackedWidget pour les panneaux ---
        self.stacked = QStackedWidget()
        main_layout.addWidget(self.stacked)

        # --- Panneau Temps de travail ---
        temps_widget = QWidget()
        temps_layout = QVBoxLayout(temps_widget)

        # --- Récupération des membres d'équipe et direction (table 'equipe') ---
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("SELECT type, direction, nombre FROM equipe WHERE projet_id=?", (self.projet_id,))
        equipe_rows = cursor.fetchall()
        conn.close()
        # --- Grouper les membres par direction ---
        directions = {}
        membre_idx = 1
        for type_, direction, nombre in equipe_rows:
            if direction not in directions:
                directions[direction] = []
            for _ in range(int(nombre)):
                directions[direction].append((f"Membre {membre_idx}", type_))
                membre_idx += 1

        # --- Récupération des investissements du projet ---
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("SELECT montant, date_achat, duree FROM investissements WHERE projet_id=?", (self.projet_id,))
        investissements = cursor.fetchall()
        conn.close()

        # --- Tableau par direction ---
        self.tables = []
        for direction, membres in directions.items():
            direction_label = QLabel(f"Direction : {direction}")
            temps_layout.addWidget(direction_label)
            colonnes = [
                "Membre",
                "Catégorie",
                "Nombre de jours",
                "Montant chargé",
                "Coût de production",
                "Coût complet"
            ]
            table = QTableWidget()
            table.setRowCount(len(membres))
            table.setColumnCount(len(colonnes))
            table.setHorizontalHeaderLabels(colonnes)

            # Récupère l'année sélectionnée
            def get_annee():
                try:
                    return int(self.annee_combo.currentText())
                except Exception:
                    return None

            # Récupère les coûts pour une catégorie et une année
            def get_couts(categorie, annee):
                conn = sqlite3.connect('gestion_budget.db')
                cursor = conn.cursor()
                cursor.execute("SELECT montant_charge, cout_production, cout_complet FROM categorie_cout WHERE annee=? AND categorie=?", (annee, categorie))
                res = cursor.fetchone()
                conn.close()
                if res:
                    return res
                return (0, 0, 0)

            # Remplit le tableau
            for row, (nom, categorie) in enumerate(membres):
                item_nom = QTableWidgetItem(nom)
                item_nom.setFlags(item_nom.flags() & ~Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 0, item_nom)
                item_cat = QTableWidgetItem(categorie)
                item_cat.setFlags(item_cat.flags() & ~Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 1, item_cat)
                item_jours = QTableWidgetItem("")
                item_jours.setFlags(item_jours.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 2, item_jours)
                # Colonnes calculées (initialisées à vide)
                table.setItem(row, 3, QTableWidgetItem(""))
                table.setItem(row, 4, QTableWidgetItem(""))
                table.setItem(row, 5, QTableWidgetItem(""))

            # Fonction de calcul automatique
            def update_row(row, table=table):
                annee = get_annee()
                categorie = table.item(row, 1).text()
                try:
                    jours = float(table.item(row, 2).text())
                except Exception:
                    jours = 0
                montant_charge, cout_production, cout_complet = get_couts(categorie, annee)
                table.item(row, 3).setText(f"{jours * montant_charge:.2f}")
                table.item(row, 4).setText(f"{jours * cout_production:.2f}")
                table.item(row, 5).setText(f"{jours * cout_complet:.2f}")

            def on_item_changed(item, table=table):
                if item.column() == 2:
                    update_row(item.row(), table)

            table.itemChanged.connect(on_item_changed)

            # Met à jour tous les calculs si l'année change
            def update_all_rows(table=table):
                for row in range(table.rowCount()):
                    update_row(row, table)
            self.annee_combo.currentTextChanged.connect(lambda _: update_all_rows(table))

            temps_layout.addWidget(table)
            self.tables.append(table)

        aide_label = QLabel("Remplissez le nombre de jours pour chaque membre. Les montants sont calculés automatiquement selon l'année et la catégorie.")
        temps_layout.addWidget(aide_label)

        self.stacked.addWidget(temps_widget)

        # --- Panneau Recettes (vide pour l'instant, mais utilisera self.annee_combo) ---
        recettes_widget = QWidget()
        recettes_layout = QVBoxLayout(recettes_widget)
        recettes_layout.addWidget(QLabel("Panneau Recettes (à compléter)"))
        self.stacked.addWidget(recettes_widget)

        # --- Panneau Dépenses (vide pour l'instant, mais utilisera self.annee_combo) ---
        depenses_widget = QWidget()
        depenses_layout = QVBoxLayout(depenses_widget)
        depenses_layout.addWidget(QLabel("Panneau Dépenses (à compléter)"))
        self.stacked.addWidget(depenses_widget)

        # --- Connexion des boutons ---
        self.btn_temps.clicked.connect(lambda: self.stacked.setCurrentIndex(0))
        self.btn_recettes.clicked.connect(lambda: self.stacked.setCurrentIndex(1))
        self.btn_depenses.clicked.connect(lambda: self.stacked.setCurrentIndex(2))

        self.setLayout(main_layout)
        self.btn_temps.clicked.connect(lambda: self.stacked.setCurrentIndex(0))
        self.btn_recettes.clicked.connect(lambda: self.stacked.setCurrentIndex(1))
        self.btn_depenses.clicked.connect(lambda: self.stacked.setCurrentIndex(2))

        self.setLayout(main_layout)
