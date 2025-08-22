from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QTableWidget, QTableWidgetItem
from PyQt6.QtCore import Qt
import sqlite3
import re

class BudgetEditDialog(QDialog):
    def __init__(self, projet_id, parent=None):
        super().__init__(parent)
        self.projet_id = projet_id
        self.setWindowTitle("Modifier le budget du projet")
        self.setMinimumSize(900, 600)  # Agrandissement de la taille minimale
        self.resize(1100, 700)         # Taille initiale plus grande
        layout = QVBoxLayout()

        # Fonction pour extraire l'année
        def extract_year(date_str):
            if not date_str:
                return ""
            # Cherche 4 chiffres consécutifs
            match = re.search(r"\b(\d{4})\b", str(date_str))
            return match.group(1) if match else str(date_str)

        # Récupérer les années du projet
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (self.projet_id,))
        row = cursor.fetchone()
        conn.close()
        if row:
            annee1 = extract_year(row[0]) if row[0] else "Année 1"
            annee2 = extract_year(row[1]) if row[1] else "Année 2"
        else:
            annee1, annee2 = "Année 1", "Année 2"

        # Détermine le nombre de colonnes à afficher
        if annee1 == annee2:
            colonnes = [annee1]
        else:
            colonnes = [annee1, annee2]

        # Tableau des données budgétaires
        self.table = QTableWidget()
        self.table.setRowCount(10)
        self.table.setColumnCount(len(colonnes))
        self.table.setHorizontalHeaderLabels(colonnes)
        lignes = [
            "Nombre jours TOTAL",
            "Nombre jours par direction",
            "Valorisation en € des nombres de jours",
            "Cout direct",
            "Cout complet",
            "(cout financeur)",
            "Dépenses",
            "Subvention",
            "Crédit d’impot",
            "CA CIR"
        ]
        self.table.setVerticalHeaderLabels(lignes)
        # Rendre toutes les cellules éditables
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = QTableWidgetItem("")
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, col, item)

        layout.addWidget(self.table)
        # Ici tu pourras ajouter d'autres widgets ou boutons si nécessaire

        # === Tableau équipe projet ===
        equipe_label = QLabel("Equipe projet (saisir le nombre de jours par membre) :")
        layout.addWidget(equipe_label)

        # Récupère l'équipe du projet (avec le nombre de personnes par catégorie)
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("SELECT type, nombre FROM equipe WHERE projet_id=?", (self.projet_id,))
        equipe_rows = cursor.fetchall()
        conn.close()

        # Prépare la liste des membres (chaque membre dissocié)
        membres = []
        membre_idx = 1
        for categorie, nombre in equipe_rows:
            for _ in range(int(nombre)):
                membres.append((f"Membre {membre_idx}", categorie))
                membre_idx += 1

        # Coûts fictifs par catégorie
        couts = {
            "Ingénieur": 500,
            "Technicien": 350,
            "Manager": 600,
            "Assistant": 250,
            "Autre": 300
        }

        equipe_table = QTableWidget()
        equipe_table.setRowCount(len(membres))
        equipe_table.setColumnCount(4)
        equipe_table.setHorizontalHeaderLabels(["Membre", "Catégorie", "Nombre jours", "Coût total (€)"])

        for row, (nom, categorie) in enumerate(membres):
            # Membre (Membre 1, Membre 2, ...)
            item_nom = QTableWidgetItem(nom)
            item_nom.setFlags(item_nom.flags() & ~Qt.ItemFlag.ItemIsEditable)
            equipe_table.setItem(row, 0, item_nom)
            # Catégorie
            item_cat = QTableWidgetItem(str(categorie))
            item_cat.setFlags(item_cat.flags() & ~Qt.ItemFlag.ItemIsEditable)
            equipe_table.setItem(row, 1, item_cat)
            # Nombre jours (éditable)
            item_jours = QTableWidgetItem("")
            item_jours.setFlags(item_jours.flags() | Qt.ItemFlag.ItemIsEditable)
            equipe_table.setItem(row, 2, item_jours)
            # Coût total (non éditable, calculé)
            item_cout = QTableWidgetItem("")
            item_cout.setFlags(item_cout.flags() & ~Qt.ItemFlag.ItemIsEditable)
            equipe_table.setItem(row, 3, item_cout)

        # Fonction de calcul automatique du coût
        def update_cost(row):
            try:
                cat = equipe_table.item(row, 1).text()
                jours = float(equipe_table.item(row, 2).text())
                cout_cat = couts.get(cat, couts["Autre"])
                total = jours * cout_cat
                equipe_table.item(row, 3).setText(f"{total:.2f}")
            except Exception:
                equipe_table.item(row, 3).setText("")

        def on_item_changed(item):
            if item.column() == 2:
                update_cost(item.row())

        equipe_table.itemChanged.connect(on_item_changed)

        layout.addWidget(equipe_table)

        self.setLayout(layout)
        self.setLayout(layout)
        self.setLayout(layout)
