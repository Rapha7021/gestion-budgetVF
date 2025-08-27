from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton, QTableWidget, QTableWidgetItem,
    QComboBox, QHBoxLayout, QStackedWidget, QWidget, QMessageBox
)
from PyQt6.QtGui import QDoubleValidator
from PyQt6.QtCore import Qt
import sqlite3
import re
from calendar import month_name

class BudgetEditDialog(QDialog):
     
    def __init__(self, projet_id, parent=None):
        # --- Création table recettes si besoin ---
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS recettes (
                projet_id INTEGER,
                annee INTEGER,
                mois TEXT,
                montant REAL,
                detail TEXT,
                PRIMARY KEY (projet_id, annee, mois)
            )
        """)
        conn.commit()
        conn.close()
           # --- Création table depenses si besoin ---
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS depenses (
                projet_id INTEGER,
                annee INTEGER,
                mois TEXT,
                montant REAL,
                detail TEXT,
                PRIMARY KEY (projet_id, annee, mois)
            )
        """)
        conn.commit()
        conn.close()
        # --- Création table autres_depenses si besoin ---
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS autres_depenses (
                projet_id INTEGER,
                annee INTEGER,
                ligne_index INTEGER,
                mois TEXT,
                montant REAL,
                detail TEXT,
                PRIMARY KEY (projet_id, annee, ligne_index, mois)
            )
        """)
            
        conn.commit()
        conn.close()
        super().__init__(parent)
        self.projet_id = projet_id
        self.setWindowTitle("Budget du Projet")
        self.setMinimumSize(1100, 700)
        main_layout = QVBoxLayout()

        # --- Création table temps_travail si besoin ---
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS temps_travail (
                projet_id INTEGER,
                annee INTEGER,
                direction TEXT,
                categorie TEXT,
                mois TEXT,
                jours REAL,
                PRIMARY KEY (projet_id, annee, direction, categorie, mois)
            )
        """)
        conn.commit()
        conn.close()

        # --- Stockage des valeurs par année ---
        self.budget_data = {}  # {annee: {(direction, categorie, mois): jours}}
        self.recettes_data = {}  # {annee: {(categorie, mois): (montant, detail)}}
        self.depenses_data = {}  # {annee: {(categorie, mois): (montant, detail)}}
        self.autres_depenses_data = {}  # {annee: {(ligne_index, mois): (montant, detail)}}
        self.modified = False  # Pour savoir si des modifs non enregistrées existent
        self.recettes_modified = False
        self.depenses_modified = False
        self.autres_depenses_modified = False
        self.depenses_categories = [
            "Salaires",
            "Achats",
            "Prestations",
            "Frais de fonctionnement",
            "Autres dépenses"
        ]
        
        # Définit les catégories de recettes
        self.recettes_categories = [
            "Subventions publiques",
            "Financement privé", 
            "Ventes/Prestations",
            "Partenariats",
            "Autres recettes"
        ]

        # --- Sélection de l'année en haut ---
        import re
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (self.projet_id,))
        row = cursor.fetchone()
        conn.close()
        self.date_debut = str(row[0]) if row and row[0] else None
        self.date_fin = str(row[1]) if row and row[1] else None
        annees = []
        if row:
            def extract_year(date_str):
                # Cherche un groupe de 4 chiffres (année) dans la date
                match = re.search(r'(\d{4})', str(date_str))
                if match:
                    return int(match.group(1))
                return None
            year_start = extract_year(row[0])
            year_end = extract_year(row[1])
            if year_start and year_end:
                # Ajoute toutes les années entre début et fin inclus
                annees = [str(y) for y in range(year_start, year_end + 1)]
            elif year_start:
                annees = [str(year_start)]
            elif year_end:
                annees = [str(year_end)]
        if not annees:
            annees = ["2024"]  # Valeur par défaut AAAA

        annee_label = QLabel("Année du projet :")
        self.annee_combo = QComboBox()
        self.annee_combo.addItems(annees)
        annee_layout = QHBoxLayout()
        annee_layout.addWidget(annee_label)
        annee_layout.addWidget(self.annee_combo)

        # --- Boutons en haut ---
        btn_layout = QHBoxLayout()
        self.btn_temps = QPushButton("Temps de travail")
        self.btn_depenses = QPushButton("Dépenses externes")
        self.btn_autres_depenses = QPushButton("Autres dépenses")
        self.btn_recettes = QPushButton("Recettes")
        self.btns = [self.btn_temps, self.btn_recettes, self.btn_depenses, self.btn_autres_depenses]
        btn_layout.addWidget(self.btn_temps)
        btn_layout.addWidget(self.btn_depenses)
        btn_layout.addWidget(self.btn_autres_depenses)
        btn_layout.addWidget(self.btn_recettes)

        # Ajout du surlignage et des connexions dans le constructeur
        def update_button_styles(active_idx):
            for i, btn in enumerate(self.btns):
                if i == active_idx:
                    btn.setStyleSheet("background-color: #87CEFA; font-weight: bold;")
                else:
                    btn.setStyleSheet("")
        self.update_button_styles = update_button_styles
        self.btn_temps.clicked.connect(lambda: (self.stacked.setCurrentIndex(0), self.update_button_styles(0)))
        self.btn_recettes.clicked.connect(lambda: (self.stacked.setCurrentIndex(1), self.update_button_styles(1)))
        self.btn_depenses.clicked.connect(lambda: (self.stacked.setCurrentIndex(2), self.update_button_styles(2)))
        self.btn_autres_depenses.clicked.connect(lambda: (self.stacked.setCurrentIndex(3), self.update_button_styles(3)))
        self.update_button_styles(0)

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
        self.temps_layout = QVBoxLayout(temps_widget)

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

        # --- Tableau unique avec colonnes dynamiques par mois ---
        mois_fr = [
            "", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
            "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
        ]
        def get_months_for_year(year):
            # Retourne la liste des mois (numéro, nom français) à afficher pour l'année donnée
            months = []
            if not self.date_debut or not self.date_fin:
                return [(m, mois_fr[m]) for m in range(1, 13)]
            def extract_ym(date_str):
                match = re.search(r'(\d{2})[/-](\d{4})', date_str)
                if match:
                    month = int(match.group(1))
                    year = int(match.group(2))
                    return year, month
                match = re.search(r'(\d{4})[/-](\d{2})', date_str)
                if match:
                    year = int(match.group(1))
                    month = int(match.group(2))
                    return year, month
                return None, None
            y_start, m_start = extract_ym(self.date_debut)
            y_end, m_end = extract_ym(self.date_fin)
            year = int(year)
            if y_start is None or y_end is None:
                return [(m, mois_fr[m]) for m in range(1, 13)]
            if year == y_start and year == y_end:
                months = [(m, mois_fr[m]) for m in range(m_start, m_end + 1)]
            elif year == y_start:
                months = [(m, mois_fr[m]) for m in range(m_start, 13)]
            elif year == y_end:
                months = [(m, mois_fr[m]) for m in range(1, m_end + 1)]
            elif y_start < year < y_end:
                months = [(m, mois_fr[m]) for m in range(1, 13)]
            return months

        def build_table_for_year(year):
            # Sauvegarde les valeurs de l'année courante avant de changer
            if hasattr(self, "current_year") and hasattr(self, "table_budget"):
                self.save_table_to_memory(self.current_year)
            self.current_year = year

            # Efface l'ancien tableau si présent
            for i in reversed(range(self.temps_layout.count())):
                widget = self.temps_layout.itemAt(i).widget()
                if widget:
                    widget.setParent(None)
            # Colonnes : Catégorie + mois
            months = get_months_for_year(year)
            colonnes = ["Catégorie"] + [mois for _, mois in months]
            table = QTableWidget()
            total_rows = sum(len(membres) + 1 for membres in directions.values())
            table.setRowCount(total_rows)
            table.setColumnCount(len(colonnes))
            table.setHorizontalHeaderLabels(colonnes)
            table.setColumnWidth(0, 200)
            row_idx = 0
            self.direction_rows = set()
            double_validator = QDoubleValidator(0.0, 9999999.99, 2)
            double_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
            for direction, membres in directions.items():
                item_dir = QTableWidgetItem(direction)
                item_dir.setFlags(item_dir.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item_dir.setBackground(Qt.GlobalColor.lightGray)
                table.setItem(row_idx, 0, item_dir)
                for col in range(1, len(colonnes)):
                    empty = QTableWidgetItem("")
                    empty.setFlags(empty.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    empty.setBackground(Qt.GlobalColor.lightGray)
                    table.setItem(row_idx, col, empty)
                self.direction_rows.add(row_idx)
                row_idx += 1
                for _, categorie in membres:
                    item_cat = QTableWidgetItem(categorie)
                    item_cat.setFlags(item_cat.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    table.setItem(row_idx, 0, item_cat)
                    # Colonnes mois éditables
                    for col in range(1, len(colonnes)):
                        item_mois = QTableWidgetItem("")
                        item_mois.setFlags(item_mois.flags() | Qt.ItemFlag.ItemIsEditable)
                        table.setItem(row_idx, col, item_mois)
                    row_idx += 1
            # Ajout du contrôle de saisie sur les cellules jours
            def on_item_changed(item):
                if item.row() in self.direction_rows:
                    return
                col = item.column()
                if col > 0:
                    value = item.text()
                    state = double_validator.validate(value, 0)[0]
                    if state != QDoubleValidator.State.Acceptable and value != "":
                        item.setText("")
                        return
                self.modified = True
            table.itemChanged.connect(on_item_changed)

            # Restaure les valeurs de l'année si elles existent
            self.restore_table_from_memory(year, table, colonnes, directions)
            
            # Si pas de données en mémoire, charge depuis la base de données
            if year not in self.budget_data or not self.budget_data[year]:
                self.load_data_from_db_for_year(year, table, colonnes)

            # Détection de modification
            def on_item_changed(item):
                if item.row() in self.direction_rows:
                    return
                self.modified = True
            table.itemChanged.connect(on_item_changed)

            self.temps_layout.addWidget(table)
            aide_label = QLabel("Remplissez le nombre de jours pour chaque membres du projet.")
            self.temps_layout.addWidget(aide_label)
            self.table_budget = table
            self.colonnes_budget = colonnes
            self.directions_budget = directions

            # Ajout du bouton Enregistrer
            btn_save = QPushButton("Enregistrer")
            def save_to_db():
                self.save_table_to_memory(year)
                conn = sqlite3.connect('gestion_budget.db')
                cursor = conn.cursor()
                # Efface les anciennes valeurs pour ce projet et cette année
                cursor.execute("DELETE FROM temps_travail WHERE projet_id=? AND annee=?", (self.projet_id, int(year)))
                # Insère toutes les valeurs de l'année
                for key, jours in self.budget_data.get(year, {}).items():
                    row_index, mois = key
                    # Récupère direction et catégorie depuis le tableau
                    current_direction = None
                    for r in range(row_index + 1):
                        if r in self.direction_rows:
                            current_direction = self.table_budget.item(r, 0).text()
                    categorie = self.table_budget.item(row_index, 0).text()
                    
                    cursor.execute("""
                        INSERT OR REPLACE INTO temps_travail (projet_id, annee, direction, categorie, mois, jours)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (self.projet_id, int(year), current_direction, categorie, mois, jours))
                conn.commit()
                conn.close()
                self.modified = False
                QMessageBox.information(self, "Enregistrement", "Les données ont été enregistrées.")
            btn_save.clicked.connect(save_to_db)
            self.temps_layout.addWidget(btn_save)



        # Initialisation du tableau pour l'année sélectionnée
        self.load_data_from_db()  # Charge les données depuis la base
        build_table_for_year(self.annee_combo.currentText())
        self.annee_combo.currentTextChanged.connect(lambda year: build_table_for_year(year))

        self.stacked.addWidget(temps_widget)

        # --- Panneau Recettes ---
        recettes_widget = QWidget()
        self.recettes_layout = QVBoxLayout(recettes_widget)
        self.recettes_table = None


        def build_recettes_table_for_year(year):
            # Sauvegarde les valeurs de l'année courante avant de changer
            if hasattr(self, "current_recettes_year") and self.recettes_table:
                self.save_recettes_table_to_memory(self.current_recettes_year)
            self.current_recettes_year = year

            # Efface l'ancien tableau si présent
            for i in reversed(range(self.recettes_layout.count())):
                widget = self.recettes_layout.itemAt(i).widget()
                if widget:
                    widget.setParent(None)

            # Utilise la même fonction pour obtenir les mois de l'année
            months = get_months_for_year(year)
            

            # Colonnes : uniquement Montant et Détail pour chaque mois
            colonnes = []
            for _, mois in months:
                colonnes.extend([f"{mois} - Montant", f"{mois} - Détail"])

            table = QTableWidget()
            # On commence avec le nombre de catégories prédéfinies, mais on pourra ajouter des lignes
            table.setRowCount(len(self.recettes_categories))
            table.setColumnCount(len(colonnes))
            table.setHorizontalHeaderLabels(colonnes)

            # Redimensionne les colonnes montant et détail
            for col in range(len(colonnes)):
                if "Montant" in colonnes[col]:
                    table.setColumnWidth(col, 120)
                elif "Détail" in colonnes[col]:
                    table.setColumnWidth(col, 200)

            # Remplit les lignes (une par catégorie, mais sans afficher la catégorie)
            for row in range(len(self.recettes_categories)):
                for col in range(len(colonnes)):
                    item_mois = QTableWidgetItem("")
                    item_mois.setFlags(item_mois.flags() | Qt.ItemFlag.ItemIsEditable)
                    table.setItem(row, col, item_mois)

            # Ajout du bouton pour ajouter une ligne
            btn_add_row = QPushButton("Ajouter une ligne")
            def add_row():
                table.insertRow(table.rowCount())
                for col in range(len(colonnes)):
                    item_mois = QTableWidgetItem("")
                    item_mois.setFlags(item_mois.flags() | Qt.ItemFlag.ItemIsEditable)
                    table.setItem(table.rowCount()-1, col, item_mois)
            btn_add_row.clicked.connect(add_row)

            # Charge systématiquement les données depuis la base et restaure le tableau
            self.load_recettes_data_from_db_for_year(year, table, colonnes, self.recettes_categories)

            # Ajout du contrôle de saisie sur les cellules Montant
            double_validator = QDoubleValidator(0.0, 9999999.99, 2)
            double_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
            def on_item_changed(item):
                col = item.column()
                if "Montant" in colonnes[col]:
                    value = item.text()
                    state = double_validator.validate(value, 0)[0]
                    if state != QDoubleValidator.State.Acceptable and value != "":
                        item.setText("")
                        return
                self.recettes_modified = True
            table.itemChanged.connect(on_item_changed)

            self.recettes_layout.addWidget(table)
            self.recettes_layout.addWidget(btn_add_row)
            aide_label = QLabel("Remplissez les montants et détails pour chaque recette.")
            self.recettes_layout.addWidget(aide_label)
            self.recettes_table = table
            self.recettes_colonnes = colonnes

            # Ajout du bouton Enregistrer
            btn_save = QPushButton("Enregistrer")
            def save_recettes_to_db():
                self.save_recettes_table_to_memory(year)
                conn = sqlite3.connect('gestion_budget.db')
                cursor = conn.cursor()
                # Efface les anciennes valeurs pour ce projet et cette année
                cursor.execute("DELETE FROM recettes WHERE projet_id=? AND annee=?", (self.projet_id, int(year)))
                # Insère toutes les valeurs de l'année
                for key, (montant, detail) in self.recettes_data.get(year, {}).items():
                    categorie, mois = key
                    if montant or detail:  # Ne sauvegarde que si au moins un champ est rempli
                        cursor.execute("""
                            INSERT INTO recettes (projet_id, annee, categorie, mois, montant, detail)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (self.projet_id, int(year), categorie, mois, montant or 0, detail or ""))
                conn.commit()
                conn.close()
                self.recettes_modified = False
                QMessageBox.information(self, "Enregistrement", "Les recettes ont été enregistrées.")
            btn_save.clicked.connect(save_recettes_to_db)
            self.recettes_layout.addWidget(btn_save)

        def save_recettes_table_to_memory(self, year):
            """Sauvegarde les valeurs du tableau recettes en mémoire pour l'année donnée"""
            if not hasattr(self, 'recettes_table') or not hasattr(self, 'recettes_colonnes'):
                return
            
            table = self.recettes_table
            colonnes = self.recettes_colonnes
            categories = self.recettes_categories
            data = {}
            
            for row in range(table.rowCount()):
                # Si la ligne correspond à une catégorie prédéfinie, on l'utilise, sinon on génère un nom générique
                if row < len(categories):
                    categorie = categories[row]
                else:
                    categorie = f"Catégorie {row+1}"
                col_index = 0
                while col_index < len(colonnes):
                    if col_index + 1 < len(colonnes):
                        col_montant_name = colonnes[col_index]
                        mois = col_montant_name.split(" - ")[0]
                        montant_item = table.item(row, col_index)
                        detail_item = table.item(row, col_index + 1)
                        montant = montant_item.text() if montant_item else ""
                        detail = detail_item.text() if detail_item else ""
                        try:
                            montant_val = float(montant) if montant else 0
                        except Exception:
                            montant_val = 0
                        if montant_val != 0 or detail:
                            key = (categorie, mois)
                            data[key] = (montant_val, detail)
                    col_index += 2
            
            self.recettes_data[year] = data

        def restore_recettes_table_from_memory(self, year, table, colonnes):
            """Restaure les valeurs du tableau recettes à partir de la mémoire pour l'année donnée"""
            data = self.recettes_data.get(year, {})
            categories = self.recettes_categories
            
            for row in range(table.rowCount()):
                if row < len(categories):
                    categorie = categories[row]
                else:
                    categorie = f"Catégorie {row+1}"
                col_index = 0
                while col_index < len(colonnes):
                    if col_index + 1 < len(colonnes):
                        col_montant_name = colonnes[col_index]
                        mois = col_montant_name.split(" - ")[0]
                        key = (categorie, mois)
                        montant, detail = data.get(key, (0, ""))
                        montant_item = table.item(row, col_index)
                        detail_item = table.item(row, col_index + 1)
                        if montant_item and detail_item:
                            table.blockSignals(True)
                            if montant != 0:
                                montant_item.setText(str(montant))
                            else:
                                montant_item.setText("")
                            detail_item.setText(detail)
                            table.blockSignals(False)
                    col_index += 2

        def load_recettes_data_from_db_for_year(self, year, table, colonnes, categories):
            """Charge les données des recettes depuis la base pour une année spécifique et les met dans le tableau"""
            conn = sqlite3.connect('gestion_budget.db')
            cursor = conn.cursor()
            cursor.execute("""
                SELECT categorie, mois, montant, detail 
                FROM recettes 
                WHERE projet_id=? AND annee=?
            """, (self.projet_id, int(year)))
            rows = cursor.fetchall()
            conn.close()

            if not rows:
                return

            # Met les données dans le tableau
            for categorie, mois, montant, detail in rows:
                found = False
                for row in range(table.rowCount()):
                    if row < len(categories):
                        cat_row = categories[row]
                    else:
                        cat_row = f"Catégorie {row+1}"
                    if cat_row == categorie:
                        col_index = 0
                        while col_index < len(colonnes):
                            if col_index + 1 < len(colonnes):
                                col_montant_name = colonnes[col_index]
                                mois_col = col_montant_name.split(" - ")[0]
                                if mois_col == mois:
                                    table.blockSignals(True)
                                    if montant:
                                        table.item(row, col_index).setText(str(montant))
                                    if detail:
                                        table.item(row, col_index + 1).setText(detail)
                                    table.blockSignals(False)
                                    break
                            col_index += 2
                        found = True
                        break

        # Affecte les méthodes à l'instance
        self.save_recettes_table_to_memory = save_recettes_table_to_memory.__get__(self)
        self.load_recettes_data_from_db_for_year = load_recettes_data_from_db_for_year.__get__(self)

        build_recettes_table_for_year(self.annee_combo.currentText())
        self.annee_combo.currentTextChanged.connect(lambda year: build_recettes_table_for_year(year))

        self.stacked.addWidget(recettes_widget)

        # --- Panneau Dépenses ---
        depenses_widget = QWidget()
        self.depenses_layout = QVBoxLayout(depenses_widget)
        self.depenses_table = None

        def build_depenses_table_for_year(year):
            if hasattr(self, "current_depenses_year") and self.depenses_table:
                self.save_depenses_table_to_memory(self.current_depenses_year)
            self.current_depenses_year = year

            for i in reversed(range(self.depenses_layout.count())):
                widget = self.depenses_layout.itemAt(i).widget()
                if widget:
                    widget.setParent(None)

            months = get_months_for_year(year)
            colonnes = []
            for _, mois in months:
                colonnes.extend([f"{mois} - Montant", f"{mois} - Détail"])

            table = QTableWidget()
            table.setRowCount(len(self.depenses_categories))
            table.setColumnCount(len(colonnes))
            table.setHorizontalHeaderLabels(colonnes)

            for col in range(len(colonnes)):
                if "Montant" in colonnes[col]:
                    table.setColumnWidth(col, 120)
                elif "Détail" in colonnes[col]:
                    table.setColumnWidth(col, 200)

            for row in range(len(self.depenses_categories)):
                for col in range(len(colonnes)):
                    item_mois = QTableWidgetItem("")
                    item_mois.setFlags(item_mois.flags() | Qt.ItemFlag.ItemIsEditable)
                    table.setItem(row, col, item_mois)

            btn_add_row = QPushButton("Ajouter une ligne")
            def add_row():
                table.insertRow(table.rowCount())
                for col in range(len(colonnes)):
                    item_mois = QTableWidgetItem("")
                    item_mois.setFlags(item_mois.flags() | Qt.ItemFlag.ItemIsEditable)
                    table.setItem(table.rowCount()-1, col, item_mois)
            btn_add_row.clicked.connect(add_row)

            # Charge systématiquement les données depuis la base et restaure le tableau
            self.load_depenses_data_from_db_for_year(year, table, colonnes, self.depenses_categories)

            # Ajout du contrôle de saisie sur les cellules Montant
            double_validator = QDoubleValidator(0.0, 9999999.99, 2)
            double_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
            def on_item_changed(item):
                col = item.column()
                if "Montant" in colonnes[col]:
                    value = item.text()
                    state = double_validator.validate(value, 0)[0]
                    if state != QDoubleValidator.State.Acceptable and value != "":
                        item.setText("")
                        return
                self.depenses_modified = True
            table.itemChanged.connect(on_item_changed)

            self.depenses_layout.addWidget(table)
            self.depenses_layout.addWidget(btn_add_row)
            aide_label = QLabel("Remplissez les montants et détails pour chaque dépense.")
            self.depenses_layout.addWidget(aide_label)
            self.depenses_table = table
            self.depenses_colonnes = colonnes

            btn_save = QPushButton("Enregistrer")
            def save_depenses_to_db():
                self.save_depenses_table_to_memory(year)
                conn = sqlite3.connect('gestion_budget.db')
                cursor = conn.cursor()
                cursor.execute("DELETE FROM depenses WHERE projet_id=? AND annee=?", (self.projet_id, int(year)))
                for key, (montant, detail) in self.depenses_data.get(year, {}).items():
                    categorie, mois = key
                    if montant or detail:
                        cursor.execute("""
                            INSERT INTO depenses (projet_id, annee, categorie, mois, montant, detail)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (self.projet_id, int(year), categorie, mois, montant or 0, detail or ""))
                conn.commit()
                conn.close()
                self.depenses_modified = False
                QMessageBox.information(self, "Enregistrement", "Les dépenses ont été enregistrées.")
            btn_save.clicked.connect(save_depenses_to_db)
            self.depenses_layout.addWidget(btn_save)

        def save_depenses_table_to_memory(self, year):
            if not hasattr(self, 'depenses_table') or not hasattr(self, 'depenses_colonnes'):
                return
            table = self.depenses_table
            colonnes = self.depenses_colonnes
            categories = self.depenses_categories
            data = {}
            for row in range(table.rowCount()):
                if row < len(categories):
                    categorie = categories[row]
                else:
                    categorie = f"Catégorie {row+1}"
                col_index = 0
                while col_index < len(colonnes):
                    if col_index + 1 < len(colonnes):
                        col_montant_name = colonnes[col_index]
                        mois = col_montant_name.split(" - ")[0]
                        montant_item = table.item(row, col_index)
                        detail_item = table.item(row, col_index + 1)
                        montant = montant_item.text() if montant_item else ""
                        detail = detail_item.text() if detail_item else ""
                        try:
                            montant_val = float(montant) if montant else 0
                        except Exception:
                            montant_val = 0
                        if montant_val != 0 or detail:
                            key = (categorie, mois)
                            data[key] = (montant_val, detail)
                    col_index += 2
            self.depenses_data[year] = data

        def load_depenses_data_from_db_for_year(self, year, table, colonnes, categories):
            conn = sqlite3.connect('gestion_budget.db')
            cursor = conn.cursor()
            cursor.execute("""
                SELECT categorie, mois, montant, detail 
                FROM depenses 
                WHERE projet_id=? AND annee=?
            """, (self.projet_id, int(year)))
            rows = cursor.fetchall()
            conn.close()
            if not rows:
                return

            # Remplit le tableau directement (comme pour temps de travail)
            for categorie, mois, montant, detail in rows:
                for row in range(table.rowCount()):
                    if row < len(categories):
                        cat_row = categories[row]
                    else:
                        cat_row = f"Catégorie {row+1}"
                    if cat_row == categorie:
                        for col in range(0, len(colonnes), 2):
                            mois_col = colonnes[col].split(" - ")[0]
                            if mois_col == mois:
                                table.blockSignals(True)
                                table.item(row, col).setText(str(montant) if montant else "")
                                table.item(row, col+1).setText(detail if detail else "")
                                table.blockSignals(False)
                                break
                        break

        self.save_depenses_table_to_memory = save_depenses_table_to_memory.__get__(self)
        self.load_depenses_data_from_db_for_year = load_depenses_data_from_db_for_year.__get__(self)

        build_depenses_table_for_year(self.annee_combo.currentText())
        self.annee_combo.currentTextChanged.connect(lambda year: build_depenses_table_for_year(year))
        self.stacked.addWidget(depenses_widget)

            # --- Panneau Autres Dépenses ---
        autres_depenses_widget = QWidget()
        self.autres_depenses_layout = QVBoxLayout(autres_depenses_widget)
        self.autres_depenses_table = None

        def build_autres_depenses_table_for_year(year):
            if hasattr(self, "current_autres_depenses_year") and self.autres_depenses_table:
                self.save_autres_depenses_table_to_memory(self.current_autres_depenses_year)
            self.current_autres_depenses_year = year

            for i in reversed(range(self.autres_depenses_layout.count())):
                widget = self.autres_depenses_layout.itemAt(i).widget()
                if widget:
                    widget.setParent(None)

            months = get_months_for_year(year)
            colonnes = []
            for _, mois in months:
                colonnes.extend([f"{mois} - Montant", f"{mois} - Détail"])

            table = QTableWidget()
            table.setRowCount(5)  # 5 lignes par défaut
            table.setColumnCount(len(colonnes))
            table.setHorizontalHeaderLabels(colonnes)

            for col in range(len(colonnes)):
                if "Montant" in colonnes[col]:
                    table.setColumnWidth(col, 120)
                elif "Détail" in colonnes[col]:
                    table.setColumnWidth(col, 200)

            for row in range(table.rowCount()):
                for col in range(len(colonnes)):
                    item_mois = QTableWidgetItem("")
                    item_mois.setFlags(item_mois.flags() | Qt.ItemFlag.ItemIsEditable)
                    table.setItem(row, col, item_mois)

            btn_add_row = QPushButton("Ajouter une ligne")
            def add_row():
                table.insertRow(table.rowCount())
                for col in range(len(colonnes)):
                    item_mois = QTableWidgetItem("")
                    item_mois.setFlags(item_mois.flags() | Qt.ItemFlag.ItemIsEditable)
                    table.setItem(table.rowCount()-1, col, item_mois)
            btn_add_row.clicked.connect(add_row)

            # Charge les données depuis la base de données
            self.load_autres_depenses_data_from_db_for_year(year, table, colonnes)

            double_validator = QDoubleValidator(0.0, 9999999.99, 2)
                # Si des modifications non enregistrées existent, restaure la mémoire
            if self.autres_depenses_modified:
                self.restore_autres_depenses_table_from_memory(year, table, colonnes)
            double_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
            def on_item_changed(item):
                col = item.column()
                if "Montant" in colonnes[col]:
                    value = item.text()
                    state = double_validator.validate(value, 0)[0]
                    if state != QDoubleValidator.State.Acceptable and value != "":
                        item.setText("")
                        return
                self.autres_depenses_modified = True
            table.itemChanged.connect(on_item_changed)

            self.autres_depenses_layout.addWidget(table)
            self.autres_depenses_layout.addWidget(btn_add_row)
            aide_label = QLabel("Remplissez les montants et détails pour chaque autre dépense.")
            self.autres_depenses_layout.addWidget(aide_label)
            self.autres_depenses_table = table
            self.autres_depenses_colonnes = colonnes

            btn_save = QPushButton("Enregistrer")
            def save_autres_depenses_to_db():
                self.save_autres_depenses_table_to_memory(year)
                conn = sqlite3.connect('gestion_budget.db')
                cursor = conn.cursor()
                cursor.execute("DELETE FROM autres_depenses WHERE projet_id=? AND annee=?", (self.projet_id, int(year)))
                for key, (montant, detail) in self.autres_depenses_data.get(year, {}).items():
                    ligne_index, mois = key
                    if montant or detail:
                        cursor.execute("""
                            INSERT INTO autres_depenses (projet_id, annee, ligne_index, mois, montant, detail)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (self.projet_id, int(year), ligne_index, mois, montant or 0, detail or ""))
                conn.commit()
                conn.close()
                self.autres_depenses_modified = False
                QMessageBox.information(self, "Enregistrement", "Les autres dépenses ont été enregistrées.")
            btn_save.clicked.connect(save_autres_depenses_to_db)
            self.autres_depenses_layout.addWidget(btn_save)

        def save_autres_depenses_table_to_memory(self, year):
            if not hasattr(self, 'autres_depenses_table') or not hasattr(self, 'autres_depenses_colonnes'):
                return
            table = self.autres_depenses_table
            colonnes = self.autres_depenses_colonnes
            data = {}
            for row in range(table.rowCount()):
                col_index = 0
                while col_index < len(colonnes):
                    if col_index + 1 < len(colonnes):
                        col_montant_name = colonnes[col_index]
                        mois = col_montant_name.split(" - ")[0]
                        montant_item = table.item(row, col_index)
                        detail_item = table.item(row, col_index + 1)
                        montant = montant_item.text() if montant_item else ""
                        detail = detail_item.text() if detail_item else ""
                        try:
                            montant_val = float(montant) if montant else 0
                        except Exception:
                            montant_val = 0
                        if montant_val != 0 or detail:
                            key = (row, mois)  # Ajout de l'index de ligne pour distinguer les lignes
                            data[key] = (montant_val, detail)
                    col_index += 2
            self.autres_depenses_data[year] = data

        def load_autres_depenses_data_from_db_for_year(self, year, table, colonnes):
            conn = sqlite3.connect('gestion_budget.db')
            cursor = conn.cursor()
            cursor.execute("""
                SELECT ligne_index, mois, montant, detail 
                FROM autres_depenses 
                WHERE projet_id=? AND annee=?
            """, (self.projet_id, int(year)))
            rows = cursor.fetchall()
            conn.close()
            if not rows:
                return
            
            # Met à jour aussi les données en mémoire
            data = {}
            for ligne_index, mois, montant, detail in rows:
                # Ajoute aux données mémoire
                key = (ligne_index, mois)
                data[key] = (montant or 0, detail or "")
                
                # Met dans le tableau
                if ligne_index < table.rowCount():
                    for col in range(0, len(colonnes), 2):
                        mois_col = colonnes[col].split(" - ")[0]
                        if mois_col == mois:
                            table.blockSignals(True)
                            table.item(ligne_index, col).setText(str(montant) if montant else "")
                            table.item(ligne_index, col+1).setText(detail if detail else "")
                            table.blockSignals(False)
                            break
            
            # Sauvegarde les données en mémoire
            self.autres_depenses_data[year] = data

        def restore_autres_depenses_table_from_memory(self, year, table, colonnes):
            """Restaure les valeurs du tableau autres dépenses à partir de la mémoire pour l'année donnée"""
            data = self.autres_depenses_data.get(year, {})
            
            for row in range(table.rowCount()):
                col_index = 0
                while col_index < len(colonnes):
                    if col_index + 1 < len(colonnes):
                        col_montant_name = colonnes[col_index]
                        mois = col_montant_name.split(" - ")[0]
                        key = (row, mois)
                        montant, detail = data.get(key, (0, ""))
                        montant_item = table.item(row, col_index)
                        detail_item = table.item(row, col_index + 1)
                        if montant_item and detail_item:
                            table.blockSignals(True)
                            if montant != 0:
                                montant_item.setText(str(montant))
                            else:
                                montant_item.setText("")
                            detail_item.setText(detail)
                            table.blockSignals(False)
                    col_index += 2

        self.save_autres_depenses_table_to_memory = save_autres_depenses_table_to_memory.__get__(self)
        self.load_autres_depenses_data_from_db_for_year = load_autres_depenses_data_from_db_for_year.__get__(self)
        self.restore_autres_depenses_table_from_memory = restore_autres_depenses_table_from_memory.__get__(self)

        build_autres_depenses_table_for_year(self.annee_combo.currentText())
        self.annee_combo.currentTextChanged.connect(lambda year: build_autres_depenses_table_for_year(year))
        self.stacked.addWidget(autres_depenses_widget)

        # --- Connexion des boutons ---
        self.btn_temps.clicked.connect(lambda: (self.stacked.setCurrentIndex(0), self.update_button_styles(0)))
        self.btn_recettes.clicked.connect(lambda: (self.stacked.setCurrentIndex(1), self.update_button_styles(1)))
        self.btn_depenses.clicked.connect(lambda: (self.stacked.setCurrentIndex(2), self.update_button_styles(2)))
        self.btn_autres_depenses.clicked.connect(lambda: (self.stacked.setCurrentIndex(3), self.update_button_styles(3)))
    # Surligne le bouton actif au démarrage
        self.update_button_styles(0)

        self.setLayout(main_layout)

        # --- Gestion fermeture : avertir si non enregistré ---
        def closeEvent(event):
            if self.modified or self.recettes_modified or self.depenses_modified or self.autres_depenses_modified:
                reply = QMessageBox.question(
                    self,
                    "Modifications non enregistrées",
                    "Des modifications n'ont pas été enregistrées. Voulez-vous vraiment quitter ?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.No:
                    event.ignore()
                    return
            event.accept()
        self.closeEvent = closeEvent

    def save_table_to_memory(self, year):
        """Sauvegarde les valeurs du tableau en mémoire pour l'année donnée"""
        if not hasattr(self, 'table_budget') or not hasattr(self, 'colonnes_budget'):
            return
        
        table = self.table_budget
        colonnes = self.colonnes_budget
        data = {}
        
        for row in range(table.rowCount()):
            # Si ligne direction, on passe
            if row in self.direction_rows:
                continue
            
            for col in range(1, table.columnCount()):
                mois = colonnes[col]
                item = table.item(row, col)
                jours = item.text() if item else ""
                try:
                    jours_val = float(jours) if jours else 0
                except Exception:
                    jours_val = 0
                # Sauvegarde par clé unique avec l'index de ligne
                key = (row, mois)
                if jours_val != 0:  # On ne sauvegarde que les valeurs non nulles
                    data[key] = jours_val
        
        self.budget_data[year] = data


    def restore_table_from_memory(self, year, table, colonnes, directions):
        """Restaure les valeurs du tableau à partir de la mémoire pour l'année donnée"""
        data = self.budget_data.get(year, {})
        
        for row in range(table.rowCount()):
            # Si ligne direction, on passe
            if row in self.direction_rows:
                continue
            
            for col in range(1, table.columnCount()):
                mois = colonnes[col]
                # Utilise la clé avec l'index de ligne
                key = (row, mois)
                val = data.get(key, 0)
                
                item = table.item(row, col)
                if item:
                    # Désactive temporairement les signaux pour éviter la détection de modification
                    table.blockSignals(True)
                    if val != 0:
                        item.setText(str(val))
                    else:
                        item.setText("")
                    table.blockSignals(False)

    def load_data_from_db(self):
        """Charge les données depuis la base de données"""
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("""
            SELECT annee, direction, categorie, mois, jours 
            FROM temps_travail 
            WHERE projet_id=?
            ORDER BY annee, direction, categorie, mois
        """, (self.projet_id,))
        rows = cursor.fetchall()
        conn.close()

    def load_data_from_db_for_year(self, year, table, colonnes):
        """Charge les données depuis la base pour une année spécifique et les met dans le tableau"""
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        cursor.execute("""
            SELECT direction, categorie, mois, jours 
            FROM temps_travail 
            WHERE projet_id=? AND annee=?
        """, (self.projet_id, int(year)))
        rows = cursor.fetchall()
        conn.close()

        if not rows:
            return

        # Trouve la ligne correspondante pour chaque donnée
        for direction, categorie, mois, jours in rows:
            for row in range(table.rowCount()):
                if row in self.direction_rows:
                    continue
                
                # Trouve la direction pour cette ligne
                current_direction = None
                for r in range(row + 1):
                    if r in self.direction_rows:
                        current_direction = table.item(r, 0).text()
                
                # Vérifie si c'est la bonne ligne
                if (current_direction == direction and 
                    table.item(row, 0).text() == categorie):
                    
                    # Trouve la colonne du mois
                    for col in range(1, table.columnCount()):
                        if colonnes[col] == mois:
                            table.blockSignals(True)
                            table.item(row, col).setText(str(jours))
                            table.blockSignals(False)
                            break
                    break