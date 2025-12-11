from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton, QTableWidget, QTableWidgetItem,
    QComboBox, QHBoxLayout, QStackedWidget, QWidget, QMessageBox, QFileDialog
)
from PyQt6.QtGui import QDoubleValidator
from PyQt6.QtCore import Qt
import sqlite3
import re
from calendar import month_name

from database import get_connection

class BudgetEditDialog(QDialog):
     
    def __init__(self, projet_id, parent=None):
        # --- Création table recettes si besoin ---
        conn = get_connection()
        cursor = conn.cursor()
        
        # Vérifier la structure actuelle de la table recettes
        cursor.execute("PRAGMA table_info(recettes)")
        columns = cursor.fetchall()
        column_names = [col[1] for col in columns]
        
        # Si la table a la structure ancienne (avec id et libelle), la migrer
        if 'id' in column_names and 'libelle' in column_names and 'ligne_index' not in column_names:
            # Créer la nouvelle table avec la structure correcte
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS recettes_new (
                    projet_id INTEGER,
                    annee INTEGER,
                    ligne_index INTEGER,
                    mois TEXT,
                    montant REAL,
                    detail TEXT,
                    PRIMARY KEY (projet_id, annee, ligne_index, mois)
                )
            """)
            
            # Copier les données en assignant des indices de ligne
            cursor.execute("""
                INSERT INTO recettes_new (projet_id, annee, ligne_index, mois, montant, detail)
                SELECT projet_id, annee, 
                       ROW_NUMBER() OVER (PARTITION BY projet_id, annee ORDER BY id) - 1 as ligne_index,
                       mois, montant, COALESCE(libelle, '') as detail
                FROM recettes
            """)
            
            # Supprimer l'ancienne table et renommer la nouvelle
            cursor.execute("DROP TABLE recettes")
            cursor.execute("ALTER TABLE recettes_new RENAME TO recettes")
        elif 'ligne_index' not in column_names:
            # Créer la table avec la nouvelle structure
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS recettes (
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
        # --- Création table depenses si besoin ---
        conn = get_connection()
        cursor = conn.cursor()
        
        # Vérifier la structure actuelle de la table depenses
        cursor.execute("PRAGMA table_info(depenses)")
        columns = cursor.fetchall()
        column_names = [col[1] for col in columns]
        
        # Si la table a la structure ancienne (avec id et libelle), la migrer
        if 'id' in column_names and 'libelle' in column_names and 'categorie' not in column_names:
            # Créer la nouvelle table avec la structure correcte
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS depenses_new (
                    projet_id INTEGER,
                    annee INTEGER,
                    categorie TEXT,
                    mois TEXT,
                    montant REAL,
                    detail TEXT,
                    PRIMARY KEY (projet_id, annee, categorie, mois)
                )
            """)
            
            # Copier les données en utilisant libelle comme categorie
            cursor.execute("""
                INSERT INTO depenses_new (projet_id, annee, categorie, mois, montant, detail)
                SELECT projet_id, annee, 
                       COALESCE(libelle, 'Autres') as categorie,
                       mois, montant, '' as detail
                FROM depenses
            """)
            
            # Supprimer l'ancienne table et renommer la nouvelle
            cursor.execute("DROP TABLE depenses")
            cursor.execute("ALTER TABLE depenses_new RENAME TO depenses")
        elif 'categorie' not in column_names:
            # Créer la table avec la nouvelle structure
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS depenses (
                    projet_id INTEGER,
                    annee INTEGER,
                    categorie TEXT,
                    mois TEXT,
                    montant REAL,
                    detail TEXT,
                    PRIMARY KEY (projet_id, annee, categorie, mois)
                )
            """)
        
        conn.commit()
        conn.close()
        # --- Création table autres_depenses si besoin ---
        conn = get_connection()
        cursor = conn.cursor()
        
        # Vérifier la structure actuelle de la table autres_depenses
        cursor.execute("PRAGMA table_info(autres_depenses)")
        columns = cursor.fetchall()
        column_names = [col[1] for col in columns]
        
        # Si la table a la structure ancienne (avec id et libelle), la migrer
        if 'id' in column_names and 'libelle' in column_names and 'ligne_index' not in column_names:
            # Créer la nouvelle table avec la structure correcte
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS autres_depenses_new (
                    projet_id INTEGER,
                    annee INTEGER,
                    ligne_index INTEGER,
                    mois TEXT,
                    montant REAL,
                    detail TEXT,
                    PRIMARY KEY (projet_id, annee, ligne_index, mois)
                )
            """)
            
            # Copier les données en assignant des indices de ligne
            cursor.execute("""
                INSERT INTO autres_depenses_new (projet_id, annee, ligne_index, mois, montant, detail)
                SELECT projet_id, annee, 
                       ROW_NUMBER() OVER (PARTITION BY projet_id, annee ORDER BY id) - 1 as ligne_index,
                       mois, montant, COALESCE(libelle, '') as detail
                FROM autres_depenses
            """)
            
            # Supprimer l'ancienne table et renommer la nouvelle
            cursor.execute("DROP TABLE autres_depenses")
            cursor.execute("ALTER TABLE autres_depenses_new RENAME TO autres_depenses")
        elif 'ligne_index' not in column_names:
            # Créer la table avec la nouvelle structure
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
        conn = get_connection()
        cursor = conn.cursor()
        
        # Vérifier la structure actuelle de la table temps_travail
        cursor.execute("PRAGMA table_info(temps_travail)")
        columns = cursor.fetchall()
        column_names = [col[1] for col in columns]
        
        # Si la table n'a pas la colonne membre_id, la migrer
        if 'membre_id' not in column_names:
            # Ajouter la colonne membre_id
            try:
                cursor.execute("ALTER TABLE temps_travail ADD COLUMN membre_id TEXT")
                
                # Mettre à jour les données existantes avec des membre_id générés
                cursor.execute("SELECT ROWID, projet_id, annee, direction, categorie FROM temps_travail")
                existing_rows = cursor.fetchall()
                
                # Générer des membre_id uniques pour les données existantes
                member_counters = {}
                for rowid, projet_id, annee, direction, categorie in existing_rows:
                    key = f"{direction}_{categorie}"
                    if key not in member_counters:
                        member_counters[key] = 0
                    
                    membre_id = f"{direction}_{categorie}_{member_counters[key]}"
                    member_counters[key] += 1
                    
                    cursor.execute("UPDATE temps_travail SET membre_id = ? WHERE ROWID = ?", 
                                 (membre_id, rowid))
                
                # Maintenant créer la nouvelle table avec la nouvelle structure
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS temps_travail_new (
                        projet_id INTEGER,
                        annee INTEGER,
                        direction TEXT,
                        categorie TEXT,
                        membre_id TEXT,
                        mois TEXT,
                        jours REAL,
                        PRIMARY KEY (projet_id, annee, membre_id, mois)
                    )
                """)
                
                # Copier les données vers la nouvelle table
                cursor.execute("""
                    INSERT INTO temps_travail_new (projet_id, annee, direction, categorie, membre_id, mois, jours)
                    SELECT projet_id, annee, direction, categorie, membre_id, mois, jours
                    FROM temps_travail
                """)
                
                # Supprimer l'ancienne table et renommer la nouvelle
                cursor.execute("DROP TABLE temps_travail")
                cursor.execute("ALTER TABLE temps_travail_new RENAME TO temps_travail")
                
            except sqlite3.Error as e:
                print(f"Erreur lors de la migration: {e}")
        else:
            # La table existe déjà avec la bonne structure, vérifier la clé primaire
            cursor.execute("PRAGMA table_info(temps_travail)")
            columns = cursor.fetchall()
            
            # Si la structure est différente, recréer la table
            expected_columns = ['projet_id', 'annee', 'direction', 'categorie', 'membre_id', 'mois', 'jours']
            actual_columns = [col[1] for col in columns]
            
            if actual_columns != expected_columns:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS temps_travail (
                        projet_id INTEGER,
                        annee INTEGER,
                        direction TEXT,
                        categorie TEXT,
                        membre_id TEXT,
                        mois TEXT,
                        jours REAL,
                        PRIMARY KEY (projet_id, annee, membre_id, mois)
                    )
                """)
        
        conn.commit()
        conn.close()

        # --- Stockage des valeurs par année ---
        self.budget_data = {}  # {annee: {(row_index, mois): {'jours': X, 'direction': Y, 'categorie': Z, 'membre_id': W}}}
        self.recettes_data = {}  # {annee: {(ligne_index, mois): (montant, detail)}}
        self.depenses_data = {}  # {annee: {(categorie, mois): (montant, detail)}}
        self.autres_depenses_data = {}  # {annee: {(ligne_index, mois): (montant, detail)}}
        self.modified_years = set()  # Années modifiées pour le temps de travail
        self.recettes_modified_years = set()  # Années modifiées pour les recettes
        self.depenses_modified_years = set()  # Années modifiées pour les dépenses
        self.autres_depenses_modified_years = set()  # Années modifiées pour les autres dépenses
        self.depenses_categories = [
            "Salaires",
            "Achats",
            "Prestations",
            "Frais de fonctionnement",
            "Autres dépenses"
        ]

        # --- Sélection de l'année en haut ---
        import re
        conn = get_connection()
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

        # --- Boutons Import/Export Excel ---
        excel_layout = QHBoxLayout()
        self.btn_generer_modele = QPushButton("📄 Générer Modèle Excel")
        self.btn_generer_modele.setToolTip(
            "Générer un fichier Excel modèle avec listes déroulantes\n"
            "pour saisir manuellement les données du projet"
        )
        self.btn_generer_modele.clicked.connect(self.generer_modele_excel)
        
        self.btn_importer_modele = QPushButton("📥 Importer Données Excel")
        self.btn_importer_modele.setToolTip(
            "Importer les données depuis un fichier Excel modèle\n"
            "contenant : Temps de travail, Dépenses, Recettes"
        )
        self.btn_importer_modele.clicked.connect(self.importer_modele_excel)
        
        excel_layout.addWidget(self.btn_generer_modele)
        excel_layout.addWidget(self.btn_importer_modele)
        excel_layout.addStretch()
        
        # --- Regroupe année + boutons ---
        top_layout = QHBoxLayout()
        top_layout.addLayout(annee_layout)
        top_layout.addLayout(btn_layout)
        main_layout.addLayout(top_layout)
        main_layout.addLayout(excel_layout)

        # --- QStackedWidget pour les panneaux ---
        self.stacked = QStackedWidget()
        main_layout.addWidget(self.stacked)

        # --- Panneau Temps de travail ---
        temps_widget = QWidget()
        self.temps_layout = QVBoxLayout(temps_widget)

        # --- Récupération des membres d'équipe et direction (table 'equipe') ---
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT type, direction, nombre FROM equipe WHERE projet_id=?", (self.projet_id,))
        equipe_rows = cursor.fetchall()
        conn.close()
        # --- Grouper les membres par direction ---
        directions = {}
        membre_idx = 1
        self.membre_mapping = {}  # Mappage row_index -> identifiant unique
        for type_, direction, nombre in equipe_rows:
            if direction not in directions:
                directions[direction] = []
            for i in range(int(nombre)):
                membre_id = f"{direction}_{type_}_{i}"  # Identifiant unique
                directions[direction].append((f"Membre {membre_idx}", type_, membre_id))
                membre_idx += 1

        # --- Récupération des investissements du projet ---
        conn = get_connection()
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
                # Essaie plusieurs formats de date
                date_str = str(date_str).strip()
                
                # Format MM/YYYY ou MM-YYYY
                match = re.search(r'(\d{1,2})[/-](\d{4})', date_str)
                if match:
                    month = int(match.group(1))
                    year = int(match.group(2))
                    return year, month
                
                # Format YYYY/MM ou YYYY-MM  
                match = re.search(r'(\d{4})[/-](\d{1,2})', date_str)
                if match:
                    year = int(match.group(1))
                    month = int(match.group(2))
                    return year, month
                
                # Format DD/MM/YYYY ou DD-MM-YYYY
                match = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', date_str)
                if match:
                    day = int(match.group(1))
                    month = int(match.group(2))
                    year = int(match.group(3))
                    return year, month
                
                # Format YYYY/MM/DD ou YYYY-MM-DD
                match = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', date_str)
                if match:
                    year = int(match.group(1))
                    month = int(match.group(2))
                    day = int(match.group(3))
                    return year, month
                
                return None, None
            
            y_start, m_start = extract_ym(self.date_debut)
            y_end, m_end = extract_ym(self.date_fin)
            year = int(year)
            
            
            if y_start is None or y_end is None:
                return [(m, mois_fr[m]) for m in range(1, 13)]
            
            if year == y_start and year == y_end:
                # Même année pour début et fin
                months = [(m, mois_fr[m]) for m in range(m_start, m_end + 1)]
            elif year == y_start:
                # Année de début du projet
                months = [(m, mois_fr[m]) for m in range(m_start, 13)]
            elif year == y_end:
                # Année de fin du projet
                months = [(m, mois_fr[m]) for m in range(1, m_end + 1)]
            elif y_start < year < y_end:
                # Année intermédiaire
                months = [(m, mois_fr[m]) for m in range(1, 13)]
            else:
                # Année hors période du projet
                months = []
            
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
                for _, categorie, membre_id in membres:
                    item_cat = QTableWidgetItem(categorie)
                    item_cat.setFlags(item_cat.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    table.setItem(row_idx, 0, item_cat)
                    # Stocke le mapping row -> membre_id
                    self.membre_mapping[row_idx] = membre_id
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
                col = item.column()
                if col > 0:
                    value = item.text()
                    state = double_validator.validate(value, 0)[0]
                    if state != QDoubleValidator.State.Acceptable and value != "":
                        item.setText("")
                        return
                self.modified_years.add(year)
            table.itemChanged.connect(on_item_changed)

            self.temps_layout.addWidget(table)
            aide_label = QLabel("Remplissez le nombre de jours pour chaque membres du projet.")
            self.temps_layout.addWidget(aide_label)
            self.table_budget = table
            self.colonnes_budget = colonnes
            self.directions_budget = directions



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

            # Nouveau tableau simplifié : Montant | Détail | Mois
            table = QTableWidget()
            table.setRowCount(0)  # Commence vide
            table.setColumnCount(3)
            table.setHorizontalHeaderLabels(["Montant", "Détail", "Mois"])
            
            # Largeurs des colonnes
            table.setColumnWidth(0, 100)  # Montant
            table.setColumnWidth(1, 300)  # Détail
            table.setColumnWidth(2, 120)  # Mois

            def add_row():
                row = table.rowCount()
                table.insertRow(row)
                
                # Colonne Montant
                montant_item = QTableWidgetItem("")
                montant_item.setFlags(montant_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 0, montant_item)
                
                # Colonne Détail  
                detail_item = QTableWidgetItem("")
                detail_item.setFlags(detail_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 1, detail_item)
                
                # Colonne Mois (ComboBox) - Seulement les mois du projet
                mois_combo = QComboBox()
                months = get_months_for_year(year)
                mois_items = [""]  # Option vide en premier
                for _, mois_nom in months:
                    mois_items.append(mois_nom)
                mois_combo.addItems(mois_items)
                mois_combo.currentTextChanged.connect(lambda: self.recettes_modified_years.add(year))
                table.setCellWidget(row, 2, mois_combo)

            btn_add_row = QPushButton("Ajouter une ligne")
            btn_add_row.clicked.connect(add_row)
            
            btn_delete_row = QPushButton("Supprimer la ligne sélectionnée")
            def delete_row():
                current_row = table.currentRow()
                if current_row >= 0:
                    reply = QMessageBox.question(
                        self,
                        "Confirmation de suppression",
                        "Êtes-vous sûr de vouloir supprimer cette ligne ?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.Yes:
                        # Suppression visuelle
                        table.removeRow(current_row)
                        self.recettes_modified_years.add(year)
                        
                        # Sauvegarde immédiate pour synchroniser la base de données
                        self.save_recettes_table_to_memory(year)
                        conn = get_connection()
                        cursor = conn.cursor()
                        
                        # Supprime toutes les recettes de cette année
                        cursor.execute("DELETE FROM recettes WHERE projet_id=? AND annee=?", 
                                     (self.projet_id, int(year)))
                        
                        # Recrée les données à partir du tableau actuel
                        data = self.recettes_data.get(year, {})
                        ligne_index = 0
                        for (_, mois), (montant, detail) in data.items():
                            if montant != 0 or detail.strip():
                                cursor.execute("""
                                    INSERT OR REPLACE INTO recettes 
                                    (projet_id, annee, ligne_index, mois, montant, detail) 
                                    VALUES (?, ?, ?, ?, ?, ?)
                                """, (self.projet_id, int(year), ligne_index, mois, montant, detail))
                                ligne_index += 1
                        
                        conn.commit()
                        conn.close()
            btn_delete_row.clicked.connect(delete_row)

            # Restaure les données depuis la mémoire si elles existent
            self.restore_recettes_table_from_memory(year, table)
            
            # Si pas de données en mémoire, charge depuis la base de données
            if year not in self.recettes_data or not self.recettes_data[year]:
                self.load_recettes_data_from_db_for_year(year, table)

            # Validation des montants
            double_validator = QDoubleValidator(0.0, 9999999.99, 2)
            double_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
            def on_item_changed(item):
                if item.column() == 0:  # Colonne Montant
                    value = item.text()
                    state = double_validator.validate(value, 0)[0]
                    if state != QDoubleValidator.State.Acceptable and value != "":
                        item.setText("")
                        return
                self.recettes_modified_years.add(year)
            table.itemChanged.connect(on_item_changed)

            # Ajout des widgets
            buttons_layout = QHBoxLayout()
            buttons_layout.addWidget(btn_add_row)
            buttons_layout.addWidget(btn_delete_row)
            buttons_layout.addStretch()
            
            buttons_widget = QWidget()
            buttons_widget.setLayout(buttons_layout)
            
            self.recettes_layout.addWidget(table)
            self.recettes_layout.addWidget(buttons_widget)
            aide_label = QLabel("Saisissez le montant, le détail et sélectionnez le mois pour chaque recette.")
            self.recettes_layout.addWidget(aide_label)
            self.recettes_table = table

        def save_recettes_table_to_memory(self, year):
            """Sauvegarde les valeurs du tableau recettes en mémoire pour l'année donnée"""
            if not hasattr(self, 'recettes_table'):
                return
            
            table = self.recettes_table
            data = {}
            
            for row in range(table.rowCount()):
                # Récupération des valeurs de chaque ligne
                montant_item = table.item(row, 0)  # Colonne Montant
                detail_item = table.item(row, 1)   # Colonne Détail
                mois_combo = table.cellWidget(row, 2)  # Colonne Mois (ComboBox)
                
                montant = montant_item.text() if montant_item else ""
                detail = detail_item.text() if detail_item else ""
                mois = mois_combo.currentText() if mois_combo else ""
                
                # Ne sauvegarde que si au moins un champ est rempli
                try:
                    montant_val = float(montant) if montant else 0
                except Exception:
                    montant_val = 0
                    
                if montant_val != 0 or detail.strip() or mois:
                    # Utilise le numéro de ligne comme identifiant unique
                    key = (row, mois)
                    data[key] = (montant_val, detail)
            
            self.recettes_data[year] = data
            # Marque toujours l'année comme modifiée, même si data est vide (suppression)
            self.recettes_modified_years.add(year)

        def load_recettes_data_from_db_for_year(self, year, table):
            """Charge les données des recettes depuis la base pour une année spécifique et les met dans le tableau"""
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT ligne_index, mois, montant, detail 
                FROM recettes 
                WHERE projet_id=? AND annee=?
                ORDER BY ligne_index
            """, (self.projet_id, int(year)))
            rows = cursor.fetchall()
            conn.close()

            if not rows:
                # Ajoute une ligne vide par défaut
                table.insertRow(0)
                
                # Colonne Montant
                montant_item = QTableWidgetItem("")
                montant_item.setFlags(montant_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(0, 0, montant_item)
                
                # Colonne Détail  
                detail_item = QTableWidgetItem("")
                detail_item.setFlags(detail_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(0, 1, detail_item)
                
                # Colonne Mois (ComboBox) - Seulement les mois du projet
                mois_combo = QComboBox()
                months = get_months_for_year(year)
                mois_items = [""]  # Option vide en premier
                for _, mois_nom in months:
                    mois_items.append(mois_nom)
                mois_combo.addItems(mois_items)
                mois_combo.currentTextChanged.connect(lambda: self.recettes_modified_years.add(year))
                table.setCellWidget(0, 2, mois_combo)
                return

            # Charge les données existantes
            for ligne_index, mois, montant, detail in rows:
                row = table.rowCount()
                table.insertRow(row)
                
                # Colonne Montant
                montant_item = QTableWidgetItem(str(montant) if montant else "")
                montant_item.setFlags(montant_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 0, montant_item)
                
                # Colonne Détail  
                detail_item = QTableWidgetItem(detail if detail else "")
                detail_item.setFlags(detail_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 1, detail_item)
                
                # Colonne Mois (ComboBox) - Seulement les mois du projet
                mois_combo = QComboBox()
                months = get_months_for_year(year)
                mois_items = [""]  # Option vide en premier
                for _, mois_nom in months:
                    mois_items.append(mois_nom)
                mois_combo.addItems(mois_items)
                if mois:
                    index = mois_combo.findText(mois)
                    if index >= 0:
                        mois_combo.setCurrentIndex(index)
                mois_combo.currentTextChanged.connect(lambda: self.recettes_modified_years.add(year))
                table.setCellWidget(row, 2, mois_combo)

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

            # Nouveau tableau simplifié : Montant | Détail | Mois
            table = QTableWidget()
            table.setRowCount(0)  # Commence vide
            table.setColumnCount(3)
            table.setHorizontalHeaderLabels(["Montant", "Détail", "Mois"])
            
            # Largeurs des colonnes
            table.setColumnWidth(0, 100)  # Montant
            table.setColumnWidth(1, 300)  # Détail
            table.setColumnWidth(2, 120)  # Mois

            def add_row():
                row = table.rowCount()
                table.insertRow(row)
                
                # Colonne Montant
                montant_item = QTableWidgetItem("")
                montant_item.setFlags(montant_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 0, montant_item)
                
                # Colonne Détail  
                detail_item = QTableWidgetItem("")
                detail_item.setFlags(detail_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 1, detail_item)
                
                # Colonne Mois (ComboBox) - Seulement les mois du projet
                mois_combo = QComboBox()
                months = get_months_for_year(year)
                mois_items = [""]  # Option vide en premier
                for _, mois_nom in months:
                    mois_items.append(mois_nom)
                mois_combo.addItems(mois_items)
                mois_combo.currentTextChanged.connect(lambda: self.depenses_modified_years.add(year))
                table.setCellWidget(row, 2, mois_combo)

            btn_add_row = QPushButton("Ajouter une ligne")
            btn_add_row.clicked.connect(add_row)
            
            btn_delete_row = QPushButton("Supprimer la ligne sélectionnée")
            def delete_row():
                current_row = table.currentRow()
                if current_row >= 0:
                    reply = QMessageBox.question(
                        self,
                        "Confirmation de suppression",
                        "Êtes-vous sûr de vouloir supprimer cette ligne ?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.Yes:
                        # Suppression visuelle
                        table.removeRow(current_row)
                        self.depenses_modified_years.add(year)
                        
                        # Sauvegarde immédiate pour synchroniser la base de données
                        self.save_depenses_table_to_memory(year)
                        conn = get_connection()
                        cursor = conn.cursor()
                        
                        # Supprime toutes les dépenses de cette année
                        cursor.execute("DELETE FROM depenses WHERE projet_id=? AND annee=?", 
                                     (self.projet_id, int(year)))
                        
                        # Recrée les données à partir du tableau actuel
                        data = self.depenses_data.get(year, {})
                        for (ligne_nom, mois), (montant, detail) in data.items():
                            if montant != 0 or detail.strip():
                                cursor.execute("""
                                    INSERT OR REPLACE INTO depenses 
                                    (projet_id, annee, categorie, mois, montant, detail) 
                                    VALUES (?, ?, ?, ?, ?, ?)
                                """, (self.projet_id, int(year), ligne_nom, mois, montant, detail))
                        
                        conn.commit()
                        conn.close()
            btn_delete_row.clicked.connect(delete_row)

            # Restaure les données depuis la mémoire si elles existent
            self.restore_depenses_table_from_memory(year, table)
            
            # Si pas de données en mémoire, charge depuis la base de données
            if year not in self.depenses_data or not self.depenses_data[year]:
                self.load_depenses_data_from_db_for_year(year, table)

            # Validation des montants
            double_validator = QDoubleValidator(0.0, 9999999.99, 2)
            double_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
            def on_item_changed(item):
                if item.column() == 0:  # Colonne Montant
                    value = item.text()
                    state = double_validator.validate(value, 0)[0]
                    if state != QDoubleValidator.State.Acceptable and value != "":
                        item.setText("")
                        return
                self.depenses_modified_years.add(year)
            table.itemChanged.connect(on_item_changed)

            # Ajout des widgets
            buttons_layout = QHBoxLayout()
            buttons_layout.addWidget(btn_add_row)
            buttons_layout.addWidget(btn_delete_row)
            buttons_layout.addStretch()
            
            buttons_widget = QWidget()
            buttons_widget.setLayout(buttons_layout)
            
            self.depenses_layout.addWidget(table)
            self.depenses_layout.addWidget(buttons_widget)
            aide_label = QLabel("Saisissez le montant, le détail et sélectionnez le mois pour chaque dépense externe.")
            self.depenses_layout.addWidget(aide_label)
            self.depenses_table = table

        def save_depenses_table_to_memory(self, year):
            if not hasattr(self, 'depenses_table'):
                return
            table = self.depenses_table
            data = {}
            for row in range(table.rowCount()):
                # Récupération des valeurs de chaque ligne
                montant_item = table.item(row, 0)  # Colonne Montant
                detail_item = table.item(row, 1)   # Colonne Détail
                mois_combo = table.cellWidget(row, 2)  # Colonne Mois (ComboBox)
                
                montant = montant_item.text() if montant_item else ""
                detail = detail_item.text() if detail_item else ""
                mois = mois_combo.currentText() if mois_combo else ""
                
                # Ne sauvegarde que si au moins un champ est rempli
                try:
                    montant_val = float(montant) if montant else 0
                except Exception:
                    montant_val = 0
                    
                if montant_val != 0 or detail.strip() or mois:
                    # Utilise le numéro de ligne et le mois pour créer une clé unique
                    key = (f"Ligne {row+1}", mois)
                    data[key] = (montant_val, detail)
            self.depenses_data[year] = data
            # Marque toujours l'année comme modifiée, même si data est vide (suppression)
            self.depenses_modified_years.add(year)

        def load_depenses_data_from_db_for_year(self, year, table):
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT categorie, mois, montant, detail 
                FROM depenses 
                WHERE projet_id=? AND annee=?
                ORDER BY categorie
            """, (self.projet_id, int(year)))
            rows = cursor.fetchall()
            conn.close()
            
            if not rows:
                # Ajoute une ligne vide par défaut
                table.insertRow(0)
                
                # Colonne Montant
                montant_item = QTableWidgetItem("")
                montant_item.setFlags(montant_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(0, 0, montant_item)
                
                # Colonne Détail  
                detail_item = QTableWidgetItem("")
                detail_item.setFlags(detail_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(0, 1, detail_item)
                
                # Colonne Mois (ComboBox) - Seulement les mois du projet
                mois_combo = QComboBox()
                months = get_months_for_year(year)
                mois_items = [""]  # Option vide en premier
                for _, mois_nom in months:
                    mois_items.append(mois_nom)
                mois_combo.addItems(mois_items)
                mois_combo.currentTextChanged.connect(lambda: self.depenses_modified_years.add(year))
                table.setCellWidget(0, 2, mois_combo)
                return

            # Charge les données existantes
            for categorie, mois, montant, detail in rows:
                row = table.rowCount()
                table.insertRow(row)
                
                # Colonne Montant
                montant_item = QTableWidgetItem(str(montant) if montant else "")
                montant_item.setFlags(montant_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 0, montant_item)
                
                # Colonne Détail  
                detail_item = QTableWidgetItem(detail if detail else "")
                detail_item.setFlags(detail_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 1, detail_item)
                
                # Colonne Mois (ComboBox) - Seulement les mois du projet
                mois_combo = QComboBox()
                months = get_months_for_year(year)
                mois_items = [""]  # Option vide en premier
                for _, mois_nom in months:
                    mois_items.append(mois_nom)
                mois_combo.addItems(mois_items)
                if mois:
                    index = mois_combo.findText(mois)
                    if index >= 0:
                        mois_combo.setCurrentIndex(index)
                mois_combo.currentTextChanged.connect(lambda: self.depenses_modified_years.add(year))
                table.setCellWidget(row, 2, mois_combo)

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

            # Nouveau tableau simplifié : Montant | Détail | Mois
            table = QTableWidget()
            table.setRowCount(0)  # Commence vide
            table.setColumnCount(3)
            table.setHorizontalHeaderLabels(["Montant", "Détail", "Mois"])
            
            # Largeurs des colonnes
            table.setColumnWidth(0, 100)  # Montant
            table.setColumnWidth(1, 300)  # Détail
            table.setColumnWidth(2, 120)  # Mois

            def add_row():
                row = table.rowCount()
                table.insertRow(row)
                
                # Colonne Montant
                montant_item = QTableWidgetItem("")
                montant_item.setFlags(montant_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 0, montant_item)
                
                # Colonne Détail  
                detail_item = QTableWidgetItem("")
                detail_item.setFlags(detail_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 1, detail_item)
                
                # Colonne Mois (ComboBox) - Seulement les mois du projet
                mois_combo = QComboBox()
                months = get_months_for_year(year)
                mois_items = [""]  # Option vide en premier
                for _, mois_nom in months:
                    mois_items.append(mois_nom)
                mois_combo.addItems(mois_items)
                mois_combo.currentTextChanged.connect(lambda: self.autres_depenses_modified_years.add(year))
                table.setCellWidget(row, 2, mois_combo)

            btn_add_row = QPushButton("Ajouter une ligne")
            btn_add_row.clicked.connect(add_row)
            
            btn_delete_row = QPushButton("Supprimer la ligne sélectionnée")
            def delete_row():
                current_row = table.currentRow()
                if current_row >= 0:
                    reply = QMessageBox.question(
                        self,
                        "Confirmation de suppression",
                        "Êtes-vous sûr de vouloir supprimer cette ligne ?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No
                    )
                    if reply == QMessageBox.StandardButton.Yes:
                        # Suppression visuelle
                        table.removeRow(current_row)
                        self.autres_depenses_modified_years.add(year)
                        
                        # Sauvegarde immédiate pour synchroniser la base de données
                        self.save_autres_depenses_table_to_memory(year)
                        conn = get_connection()
                        cursor = conn.cursor()
                        
                        # Supprime toutes les autres dépenses de cette année
                        cursor.execute("DELETE FROM autres_depenses WHERE projet_id=? AND annee=?", 
                                     (self.projet_id, int(year)))
                        
                        # Recrée les données à partir du tableau actuel
                        data = self.autres_depenses_data.get(year, {})
                        ligne_index = 0
                        for (_, mois), (montant, detail) in data.items():
                            if montant != 0 or detail.strip():
                                cursor.execute("""
                                    INSERT OR REPLACE INTO autres_depenses 
                                    (projet_id, annee, ligne_index, mois, montant, detail) 
                                    VALUES (?, ?, ?, ?, ?, ?)
                                """, (self.projet_id, int(year), ligne_index, mois, montant, detail))
                                ligne_index += 1
                        
                        conn.commit()
                        conn.close()
            btn_delete_row.clicked.connect(delete_row)

            # Restaure les données depuis la mémoire si elles existent
            self.restore_autres_depenses_table_from_memory(year, table)
            
            # Si pas de données en mémoire, charge depuis la base de données
            if year not in self.autres_depenses_data or not self.autres_depenses_data[year]:
                self.load_autres_depenses_data_from_db_for_year(year, table)

            # Validation des montants
            double_validator = QDoubleValidator(0.0, 9999999.99, 2)
            double_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
            def on_item_changed(item):
                if item.column() == 0:  # Colonne Montant
                    value = item.text()
                    state = double_validator.validate(value, 0)[0]
                    if state != QDoubleValidator.State.Acceptable and value != "":
                        item.setText("")
                        return
                self.autres_depenses_modified_years.add(year)
            table.itemChanged.connect(on_item_changed)

            # Ajout des widgets
            buttons_layout = QHBoxLayout()
            buttons_layout.addWidget(btn_add_row)
            buttons_layout.addWidget(btn_delete_row)
            buttons_layout.addStretch()
            
            buttons_widget = QWidget()
            buttons_widget.setLayout(buttons_layout)
            
            self.autres_depenses_layout.addWidget(table)
            self.autres_depenses_layout.addWidget(buttons_widget)
            aide_label = QLabel("Saisissez le montant, le détail et sélectionnez le mois pour chaque autre dépense.")
            self.autres_depenses_layout.addWidget(aide_label)
            self.autres_depenses_table = table

        def save_autres_depenses_table_to_memory(self, year):
            if not hasattr(self, 'autres_depenses_table'):
                return
            table = self.autres_depenses_table
            data = {}
            for row in range(table.rowCount()):
                # Récupération des valeurs de chaque ligne
                montant_item = table.item(row, 0)  # Colonne Montant
                detail_item = table.item(row, 1)   # Colonne Détail
                mois_combo = table.cellWidget(row, 2)  # Colonne Mois (ComboBox)
                
                montant = montant_item.text() if montant_item else ""
                detail = detail_item.text() if detail_item else ""
                mois = mois_combo.currentText() if mois_combo else ""
                
                # Ne sauvegarde que si au moins un champ est rempli
                try:
                    montant_val = float(montant) if montant else 0
                except Exception:
                    montant_val = 0
                    
                if montant_val != 0 or detail.strip() or mois:
                    # Utilise le numéro de ligne et le mois pour créer une clé unique
                    key = (row, mois)
                    data[key] = (montant_val, detail)
            self.autres_depenses_data[year] = data
            # Marque toujours l'année comme modifiée, même si data est vide (suppression)
            self.autres_depenses_modified_years.add(year)

        def load_autres_depenses_data_from_db_for_year(self, year, table):
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT ligne_index, mois, montant, detail 
                FROM autres_depenses 
                WHERE projet_id=? AND annee=?
                ORDER BY ligne_index
            """, (self.projet_id, int(year)))
            rows = cursor.fetchall()
            conn.close()
            
            if not rows:
                # Ajoute une ligne vide par défaut
                table.insertRow(0)
                
                # Colonne Montant
                montant_item = QTableWidgetItem("")
                montant_item.setFlags(montant_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(0, 0, montant_item)
                
                # Colonne Détail  
                detail_item = QTableWidgetItem("")
                detail_item.setFlags(detail_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(0, 1, detail_item)
                
                # Colonne Mois (ComboBox) - Seulement les mois du projet
                mois_combo = QComboBox()
                months = get_months_for_year(year)
                mois_items = [""]  # Option vide en premier
                for _, mois_nom in months:
                    mois_items.append(mois_nom)
                mois_combo.addItems(mois_items)
                mois_combo.currentTextChanged.connect(lambda: self.autres_depenses_modified_years.add(year))
                table.setCellWidget(0, 2, mois_combo)
                return

            # Charge les données existantes
            for ligne_index, mois, montant, detail in rows:
                row = table.rowCount()
                table.insertRow(row)
                
                # Colonne Montant
                montant_item = QTableWidgetItem(str(montant) if montant else "")
                montant_item.setFlags(montant_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 0, montant_item)
                
                # Colonne Détail  
                detail_item = QTableWidgetItem(detail if detail else "")
                detail_item.setFlags(detail_item.flags() | Qt.ItemFlag.ItemIsEditable)
                table.setItem(row, 1, detail_item)
                
                # Colonne Mois (ComboBox) - Seulement les mois du projet
                mois_combo = QComboBox()
                months = get_months_for_year(year)
                mois_items = [""]  # Option vide en premier
                for _, mois_nom in months:
                    mois_items.append(mois_nom)
                mois_combo.addItems(mois_items)
                if mois:
                    index = mois_combo.findText(mois)
                    if index >= 0:
                        mois_combo.setCurrentIndex(index)
                mois_combo.currentTextChanged.connect(lambda: self.autres_depenses_modified_years.add(year))
                table.setCellWidget(row, 2, mois_combo)

        self.save_autres_depenses_table_to_memory = save_autres_depenses_table_to_memory.__get__(self)
        self.load_autres_depenses_data_from_db_for_year = load_autres_depenses_data_from_db_for_year.__get__(self)

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

        # --- Ajout d'un bouton Enregistrer global ---
        btn_save_all = QPushButton("Enregistrer tout")
        main_layout.addWidget(btn_save_all)

        def save_all_to_db():
            # Sauvegarde d'abord les données de l'année actuellement affichée
            if hasattr(self, 'current_year'):
                self.save_table_to_memory(self.current_year)
            if hasattr(self, 'current_recettes_year'):
                self.save_recettes_table_to_memory(self.current_recettes_year)
            if hasattr(self, 'current_depenses_year'):
                self.save_depenses_table_to_memory(self.current_depenses_year)
            if hasattr(self, 'current_autres_depenses_year'):
                self.save_autres_depenses_table_to_memory(self.current_autres_depenses_year)

            conn = get_connection()
            cursor = conn.cursor()

            # Sauvegarde seulement les années modifiées pour le temps de travail
            for annee in self.modified_years:
                # Toujours supprimer les anciennes données, même si pas de nouvelles données
                cursor.execute("DELETE FROM temps_travail WHERE projet_id=? AND annee=?", (self.projet_id, int(annee)))
                
                data = self.budget_data.get(annee, {})
                if data:  # Seulement si il y a des données à insérer
                    for key, val_data in data.items():
                        row_index, mois = key
                        if isinstance(val_data, dict):
                            jours = val_data['jours']
                            direction = val_data['direction']
                            categorie = val_data['categorie']
                            membre_id = val_data.get('membre_id', f"membre_{row_index}")
                        else:
                            # Rétrocompatibilité si anciennes données
                            jours = val_data
                            direction = "Direction 1"
                            categorie = "Membre 1"
                            membre_id = f"membre_{row_index}"
                        
                        cursor.execute("""
                            INSERT OR REPLACE INTO temps_travail (projet_id, annee, direction, categorie, membre_id, mois, jours)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        """, (self.projet_id, int(annee), direction, categorie, membre_id, mois, jours))

            # Sauvegarde seulement les années modifiées pour les recettes
            for annee in self.recettes_modified_years:
                # Toujours supprimer les anciennes données, même si pas de nouvelles données
                cursor.execute("DELETE FROM recettes WHERE projet_id=? AND annee=?", (self.projet_id, int(annee)))
                
                data = self.recettes_data.get(annee, {})
                if data:  # Seulement si il y a des données à insérer
                    ligne_index = 0
                    for key, (montant, detail) in data.items():
                        if montant or detail:  # Ne sauvegarde que si au moins un champ est rempli
                            _, mois = key
                            cursor.execute("INSERT INTO recettes (projet_id, annee, ligne_index, mois, montant, detail) VALUES (?, ?, ?, ?, ?, ?)", 
                                         (self.projet_id, int(annee), ligne_index, mois, montant or 0, detail or ""))
                            ligne_index += 1

            # Sauvegarde seulement les années modifiées pour les dépenses
            for annee in self.depenses_modified_years:
                # Toujours supprimer les anciennes données, même si pas de nouvelles données
                cursor.execute("DELETE FROM depenses WHERE projet_id=? AND annee=?", (self.projet_id, int(annee)))
                
                data = self.depenses_data.get(annee, {})
                if data:  # Seulement si il y a des données à insérer
                    for key, (montant, detail) in data.items():
                        if montant or detail:  # Ne sauvegarde que si au moins un champ est rempli
                            ligne_nom, mois = key  # La clé est (f"Ligne {row+1}", mois)
                            # Utilise le nom de ligne comme catégorie générique pour ce système simplifié
                            cursor.execute("INSERT INTO depenses (projet_id, annee, categorie, mois, montant, detail) VALUES (?, ?, ?, ?, ?, ?)", 
                                         (self.projet_id, int(annee), ligne_nom, mois, montant or 0, detail or ""))

            # Sauvegarde seulement les années modifiées pour les autres dépenses
            for annee in self.autres_depenses_modified_years:
                # Toujours supprimer les anciennes données, même si pas de nouvelles données
                cursor.execute("DELETE FROM autres_depenses WHERE projet_id=? AND annee=?", (self.projet_id, int(annee)))
                
                data = self.autres_depenses_data.get(annee, {})
                if data:  # Seulement si il y a des données à insérer
                    ligne_index = 0
                    for key, (montant, detail) in data.items():
                        if montant or detail:  # Ne sauvegarde que si au moins un champ est rempli
                            _, mois = key
                            cursor.execute("INSERT INTO autres_depenses (projet_id, annee, ligne_index, mois, montant, detail) VALUES (?, ?, ?, ?, ?, ?)", 
                                         (self.projet_id, int(annee), ligne_index, mois, montant or 0, detail or ""))
                            ligne_index += 1

            conn.commit()
            conn.close()
            
            # Réinitialise tous les flags de modification
            self.modified_years.clear()
            self.recettes_modified_years.clear()
            self.depenses_modified_years.clear()
            self.autres_depenses_modified_years.clear()

            QMessageBox.information(self, "Enregistrement", "Toutes les données ont été enregistrées.")
            
            # Ferme la fenêtre après enregistrement
            self.accept()

        btn_save_all.clicked.connect(save_all_to_db)

        # --- Gestion fermeture : avertir si non enregistré ---
        def closeEvent(event):
            if (self.modified_years or self.recettes_modified_years or 
                self.depenses_modified_years or self.autres_depenses_modified_years):
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

    def generer_modele_excel(self):
        """Génère un fichier Excel modèle pour saisie manuelle"""
        from generer_modele_excel import creer_modele_excel
        
        # Récupérer le code du projet pour le nom de fichier
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT code FROM projets WHERE id = ?", (self.projet_id,))
        result = cursor.fetchone()
        conn.close()
        
        code_projet = result[0] if result else "Projet"
        nom_defaut = f"Modele_{code_projet}.xlsx"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Enregistrer le modèle Excel",
            nom_defaut,
            "Fichiers Excel (*.xlsx)"
        )
        
        if file_path:
            try:
                creer_modele_excel(file_path, projet_id=self.projet_id)
                
                message = f"Le modèle Excel a été créé avec succès :\n{file_path}\n\n"
                message += "✓ Le modèle contient des listes déroulantes pour :\n"
                message += "  • Directions de l'équipe\n"
                message += "  • Catégories de l'équipe\n"
                message += "  • Membres de l'équipe\n"
                message += "  • Mois de la période du projet\n\n"
                message += "Vous pouvez maintenant :\n"
                message += "1. Ouvrir le fichier Excel\n"
                message += "2. Remplir les données avec les listes déroulantes\n"
                message += "3. Utiliser 'Importer Données Excel' pour l'importer"
                
                QMessageBox.information(self, "Modèle créé", message)
            except Exception as e:
                QMessageBox.critical(self, "Erreur", f"Erreur lors de la création du modèle :\n{e}")
    
    def importer_modele_excel(self):
        """Ouvre le dialogue d'import depuis un modèle Excel"""
        from import_modele_excel_dialog import ImportExcelModeleDialog
        
        dialog = ImportExcelModeleDialog(self, projet_id=self.projet_id)
        dialog.exec()
        
        # Rafraîchir l'affichage après fermeture du dialog
        self.refresh_all_data()
    
    def refresh_all_data(self):
        """Rafraîchit toutes les données affichées"""
        current_year = self.annee_combo.currentText()
        
        # Forcer le rechargement en émettant le signal de changement d'année
        if current_year:
            # Émettre le signal manuellement pour forcer le rechargement
            self.annee_combo.currentTextChanged.emit(current_year)

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
            
            # Trouve la direction pour cette ligne
            current_direction = None
            for r in range(row + 1):
                if r in self.direction_rows:
                    current_direction = table.item(r, 0).text()
            
            categorie = table.item(row, 0).text()
            membre_id = self.membre_mapping.get(row, f"membre_{row}")  # Identifiant unique
            
            for col in range(1, table.columnCount()):
                mois = colonnes[col]
                item = table.item(row, col)
                jours = item.text() if item else ""
                try:
                    jours_val = float(jours) if jours else 0
                except Exception:
                    jours_val = 0
                # Sauvegarde avec row_index comme clé principale mais stocke aussi toutes les infos
                key = (row, mois)
                if jours_val != 0:  # On ne sauvegarde que les valeurs non nulles
                    data[key] = {
                        'jours': jours_val,
                        'direction': current_direction,
                        'categorie': categorie,
                        'membre_id': membre_id
                    }
        
        self.budget_data[year] = data
        # Marque toujours l'année comme modifiée, même si data est vide (suppression)
        self.modified_years.add(year)


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
                val_data = data.get(key, {})
                val = val_data.get('jours', 0) if isinstance(val_data, dict) else 0
                
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
        conn = get_connection()
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
        conn = get_connection()
        cursor = conn.cursor()
        
        # Vérifier si la colonne membre_id existe
        cursor.execute("PRAGMA table_info(temps_travail)")
        columns = cursor.fetchall()
        column_names = [col[1] for col in columns]
        
        if 'membre_id' in column_names:
            cursor.execute("""
                SELECT direction, categorie, membre_id, mois, jours 
                FROM temps_travail 
                WHERE projet_id=? AND annee=?
            """, (self.projet_id, int(year)))
            rows = cursor.fetchall()
            
            if rows:
                # Met aussi les données en mémoire pour cet année
                data = {}

                # Trouve la ligne correspondante pour chaque donnée
                for direction, categorie, membre_id, mois, jours in rows:
                    # Cherche la ligne correspondante par membre_id exact d'abord
                    target_row = None
                    for row, stored_membre_id in self.membre_mapping.items():
                        if stored_membre_id == membre_id:
                            target_row = row
                            break
                    
                    # Si pas de correspondance exacte, cherche par direction/catégorie
                    if target_row is None:
                        for row in range(table.rowCount()):
                            if row not in self.direction_rows:
                                # Vérifier si cette ligne correspond à la direction/catégorie
                                categorie_item = table.item(row, 0)
                                if categorie_item and categorie_item.text() == categorie:
                                    # Vérifier la direction (ligne précédente de type direction)
                                    direction_row = None
                                    for r in range(row - 1, -1, -1):
                                        if r in self.direction_rows:
                                            direction_item = table.item(r, 0)
                                            if direction_item and direction_item.text() == direction:
                                                target_row = row
                                                break
                                            break
                                    if target_row is not None:
                                        break
                    
                    if target_row is not None and target_row not in self.direction_rows:
                        # Trouve la colonne du mois
                        for col in range(1, table.columnCount()):
                            if colonnes[col] == mois:
                                table.blockSignals(True)
                                table.item(target_row, col).setText(str(jours))
                                table.blockSignals(False)
                                
                                # Stocke aussi en mémoire
                                key = (target_row, mois)
                                data[key] = {
                                    'jours': jours,
                                    'direction': direction,
                                    'categorie': categorie,
                                    'membre_id': membre_id
                                }
                                break
                
                # Sauvegarde les données en mémoire
                if data:
                    self.budget_data[year] = data
        else:
            # Ancienne structure sans membre_id - méthode de fallback
            cursor.execute("""
                SELECT direction, categorie, mois, jours 
                FROM temps_travail 
                WHERE projet_id=? AND annee=?
            """, (self.projet_id, int(year)))
            rows = cursor.fetchall()
            
            if rows:
                data = {}
                # Méthode de fallback - cherche par direction et catégorie
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
                                    
                                    # Stocke aussi en mémoire avec le nouveau format
                                    membre_id = self.membre_mapping.get(row, f"membre_{row}")
                                    key = (row, mois)
                                    data[key] = {
                                        'jours': jours,
                                        'direction': direction,
                                        'categorie': categorie,
                                        'membre_id': membre_id
                                    }
                                    found = True
                                    break
                        
                        if found:
                            break
                
                # Sauvegarde les données mises à jour
                if data:
                    self.budget_data[year] = data
        
        conn.close()

    def restore_recettes_table_from_memory(self, year, table):
        """Restaure les valeurs du tableau recettes à partir de la mémoire pour l'année donnée"""
        data = self.recettes_data.get(year, {})
        if not data:
            return
        
        # Obtenir la fonction get_months_for_year depuis la portée locale
        mois_fr = ["", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                   "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        
        def get_months_for_year_local(year):
            if not self.date_debut or not self.date_fin:
                return [(m, mois_fr[m]) for m in range(1, 13)]
            
            import re
            def extract_ym(date_str):
                date_str = str(date_str).strip()
                match = re.search(r'(\d{1,2})[/-](\d{4})', date_str)
                if match:
                    return int(match.group(2)), int(match.group(1))
                match = re.search(r'(\d{4})[/-](\d{1,2})', date_str)
                if match:
                    return int(match.group(1)), int(match.group(2))
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
            else:
                months = []
            
            return months
        
        # Restaurer les données
        for (ligne_index, mois), (montant, detail) in data.items():
            row = table.rowCount()
            table.insertRow(row)
            
            # Colonne Montant
            montant_item = QTableWidgetItem(str(montant) if montant else "")
            table.setItem(row, 0, montant_item)
            
            # Colonne Détail  
            detail_item = QTableWidgetItem(detail if detail else "")
            table.setItem(row, 1, detail_item)
            
            # Colonne Mois (ComboBox)
            from PyQt6.QtWidgets import QComboBox
            mois_combo = QComboBox()
            months = get_months_for_year_local(year)
            mois_items = [""]
            for _, mois_nom in months:
                mois_items.append(mois_nom)
            mois_combo.addItems(mois_items)
            if mois:
                index = mois_combo.findText(mois)
                if index >= 0:
                    mois_combo.setCurrentIndex(index)
            mois_combo.currentTextChanged.connect(lambda: self.recettes_modified_years.add(year))
            table.setCellWidget(row, 2, mois_combo)

    def restore_depenses_table_from_memory(self, year, table):
        """Restaure les valeurs du tableau dépenses à partir de la mémoire pour l'année donnée"""
        data = self.depenses_data.get(year, {})
        if not data:
            return
        
        # Obtenir la fonction get_months_for_year depuis la portée locale
        mois_fr = ["", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                   "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        
        def get_months_for_year_local(year):
            if not self.date_debut or not self.date_fin:
                return [(m, mois_fr[m]) for m in range(1, 13)]
            
            import re
            def extract_ym(date_str):
                date_str = str(date_str).strip()
                match = re.search(r'(\d{1,2})[/-](\d{4})', date_str)
                if match:
                    return int(match.group(2)), int(match.group(1))
                match = re.search(r'(\d{4})[/-](\d{1,2})', date_str)
                if match:
                    return int(match.group(1)), int(match.group(2))
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
            else:
                months = []
            
            return months
        
        # Restaurer les données
        for (categorie, mois), (montant, detail) in data.items():
            row = table.rowCount()
            table.insertRow(row)
            
            # Colonne Montant
            montant_item = QTableWidgetItem(str(montant) if montant else "")
            table.setItem(row, 0, montant_item)
            
            # Colonne Détail  
            detail_item = QTableWidgetItem(detail if detail else "")
            table.setItem(row, 1, detail_item)
            
            # Colonne Mois (ComboBox)
            from PyQt6.QtWidgets import QComboBox
            mois_combo = QComboBox()
            months = get_months_for_year_local(year)
            mois_items = [""]
            for _, mois_nom in months:
                mois_items.append(mois_nom)
            mois_combo.addItems(mois_items)
            if mois:
                index = mois_combo.findText(mois)
                if index >= 0:
                    mois_combo.setCurrentIndex(index)
            mois_combo.currentTextChanged.connect(lambda: self.depenses_modified_years.add(year))
            table.setCellWidget(row, 2, mois_combo)

    def restore_autres_depenses_table_from_memory(self, year, table):
        """Restaure les valeurs du tableau autres dépenses à partir de la mémoire pour l'année donnée"""
        data = self.autres_depenses_data.get(year, {})
        if not data:
            return
        
        # Obtenir la fonction get_months_for_year depuis la portée locale
        mois_fr = ["", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                   "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        
        def get_months_for_year_local(year):
            if not self.date_debut or not self.date_fin:
                return [(m, mois_fr[m]) for m in range(1, 13)]
            
            import re
            def extract_ym(date_str):
                date_str = str(date_str).strip()
                match = re.search(r'(\d{1,2})[/-](\d{4})', date_str)
                if match:
                    return int(match.group(2)), int(match.group(1))
                match = re.search(r'(\d{4})[/-](\d{1,2})', date_str)
                if match:
                    return int(match.group(1)), int(match.group(2))
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
            else:
                months = []
            
            return months
        
        # Restaurer les données
        for (ligne_index, mois), (montant, detail) in data.items():
            row = table.rowCount()
            table.insertRow(row)
            
            # Colonne Montant
            montant_item = QTableWidgetItem(str(montant) if montant else "")
            table.setItem(row, 0, montant_item)
            
            # Colonne Détail  
            detail_item = QTableWidgetItem(detail if detail else "")
            table.setItem(row, 1, detail_item)
            
            # Colonne Mois (ComboBox)
            from PyQt6.QtWidgets import QComboBox
            mois_combo = QComboBox()
            months = get_months_for_year_local(year)
            mois_items = [""]
            for _, mois_nom in months:
                mois_items.append(mois_nom)
            mois_combo.addItems(mois_items)
            if mois:
                index = mois_combo.findText(mois)
                if index >= 0:
                    mois_combo.setCurrentIndex(index)
            mois_combo.currentTextChanged.connect(lambda: self.autres_depenses_modified_years.add(year))
            table.setCellWidget(row, 2, mois_combo)
