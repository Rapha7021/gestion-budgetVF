import sqlite3
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                            QComboBox, QGroupBox, QListWidget, QListWidgetItem,
                            QRadioButton, QButtonGroup, QMessageBox)
from PyQt6.QtCore import Qt

DB_PATH = 'gestion_budget.db'

class PrintConfigDialog(QDialog):
    def __init__(self, parent, projet_id=None):
        super().__init__(parent)
        self.parent = parent
        self.projet_id = projet_id
        self.setWindowTitle("Configuration du Compte de Résultat")
        self.setMinimumSize(600, 600)

        layout = QVBoxLayout()

        # Titre
        title = QLabel("Configuration du Compte de Résultat")
        title.setStyleSheet("font-size: 16pt; font-weight: bold; margin-bottom: 20px;")
        layout.addWidget(title)

        # Sélection du projet
        project_group = QGroupBox("Sélection du projet")
        project_layout = QVBoxLayout()

        # Options de sélection
        self.selection_type_group = QButtonGroup(self)

        self.radio_single_project = QRadioButton("Un projet spécifique")
        self.radio_all_projects = QRadioButton("Tous les projets")
        self.radio_multiple_projects = QRadioButton("Sélection de certains projets")
        self.radio_by_theme = QRadioButton("Sélection par thèmes")
        self.radio_by_main_theme = QRadioButton("Sélection par le thème principal")

        self.selection_type_group.addButton(self.radio_single_project)
        self.selection_type_group.addButton(self.radio_all_projects)
        self.selection_type_group.addButton(self.radio_multiple_projects)
        self.selection_type_group.addButton(self.radio_by_theme)
        self.selection_type_group.addButton(self.radio_by_main_theme)

        self.radio_single_project.setChecked(True)

        project_layout.addWidget(self.radio_single_project)
        project_layout.addWidget(self.radio_all_projects)
        project_layout.addWidget(self.radio_multiple_projects)
        project_layout.addWidget(self.radio_by_theme)
        project_layout.addWidget(self.radio_by_main_theme)

        # Widgets dynamiques pour les options
        self.project_combo = QComboBox()
        self.project_list_widget = QListWidget()
        self.theme_list_widget = QListWidget()
        self.main_theme_list_widget = QListWidget()

        self.load_projects()
        self.load_themes()
        self.load_main_themes()

        project_layout.addWidget(self.project_combo)
        project_layout.addWidget(self.project_list_widget)
        project_layout.addWidget(self.theme_list_widget)
        project_layout.addWidget(self.main_theme_list_widget)

        self.project_combo.hide()
        self.project_list_widget.hide()
        self.theme_list_widget.hide()
        self.main_theme_list_widget.hide()

        self.radio_single_project.toggled.connect(self.update_selection_widgets)
        self.radio_all_projects.toggled.connect(self.update_selection_widgets)
        self.radio_multiple_projects.toggled.connect(self.update_selection_widgets)
        self.radio_by_theme.toggled.connect(self.update_selection_widgets)
        self.radio_by_main_theme.toggled.connect(self.update_selection_widgets)
        
        # Connecter le changement de projet pour mettre à jour les années
        self.project_combo.currentIndexChanged.connect(self.on_project_changed)

        project_group.setLayout(project_layout)
        layout.addWidget(project_group)

        # Sélection de la période
        period_group = QGroupBox("Période")
        period_layout = QVBoxLayout()

        # Type de période
        self.period_type = QComboBox()
        self.period_type.addItems(["Année spécifique", "Plusieurs années", "Période globale du projet"])
        self.period_type.currentIndexChanged.connect(self.update_period_widgets)
        period_layout.addWidget(QLabel("Type de période:"))
        period_layout.addWidget(self.period_type)

        # Sélection d'une année spécifique
        self.year_combo = QComboBox()
        period_layout.addWidget(self.year_combo)

        # Sélection de plusieurs années
        self.years_list_widget = QListWidget()
        self.years_list_widget.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        for year in range(2000, 2101):
            item = QListWidgetItem(str(year))
            item.setCheckState(Qt.CheckState.Unchecked)
            self.years_list_widget.addItem(item)
        period_layout.addWidget(self.years_list_widget)

        # Granularité (par mois ou par an)
        granularity_group = QGroupBox("Granularité")
        granularity_layout = QHBoxLayout()

        self.radio_monthly = QRadioButton("Par mois")
        self.radio_yearly = QRadioButton("Par an")
        self.radio_yearly.setChecked(True)

        granularity_layout.addWidget(self.radio_monthly)
        granularity_layout.addWidget(self.radio_yearly)
        granularity_group.setLayout(granularity_layout)
        period_layout.addWidget(granularity_group)

        period_group.setLayout(period_layout)
        layout.addWidget(period_group)

        # Type de coût
        cost_type_group = QGroupBox("Type de coût")
        cost_type_layout = QVBoxLayout()

        self.cost_type_group = QButtonGroup(self)

        self.radio_montant_charge = QRadioButton("Montant chargé")
        self.radio_cout_production = QRadioButton("Coût de production")
        self.radio_cout_complet = QRadioButton("Coût complet")

        self.cost_type_group.addButton(self.radio_montant_charge)
        self.cost_type_group.addButton(self.radio_cout_production)
        self.cost_type_group.addButton(self.radio_cout_complet)

        # Coût de production par défaut
        self.radio_cout_production.setChecked(True)

        cost_type_layout.addWidget(self.radio_montant_charge)
        cost_type_layout.addWidget(self.radio_cout_production)
        cost_type_layout.addWidget(self.radio_cout_complet)

        cost_type_group.setLayout(cost_type_layout)
        layout.addWidget(cost_type_group)

        # Boutons
        buttons_layout = QHBoxLayout()

        generate_btn = QPushButton("Générer le Compte de Résultat")
        generate_btn.clicked.connect(self.generate_compte_resultat)
        generate_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 10px; }")

        cancel_btn = QPushButton("Annuler")
        cancel_btn.clicked.connect(self.reject)

        buttons_layout.addStretch()
        buttons_layout.addWidget(generate_btn)
        buttons_layout.addWidget(cancel_btn)

        layout.addLayout(buttons_layout)
        self.setLayout(layout)

        # Initialiser l'affichage
        self.update_period_widgets()
        self.update_selection_widgets()

    def load_projects(self):
        """Charge les projets depuis la base de données"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT id, code, nom FROM projets ORDER BY code")
        projects = cursor.fetchall()
        conn.close()

        self.project_combo.clear()
        self.project_list_widget.clear()

        for project_id, code, name in projects:
            self.project_combo.addItem(f"{code} - {name}", project_id)
            item = QListWidgetItem(f"{code} - {name}")
            item.setData(Qt.ItemDataRole.UserRole, project_id)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.project_list_widget.addItem(item)
        
        # Connecter le signal pour recharger les années quand on coche/décoche des projets
        self.project_list_widget.itemChanged.connect(self.on_project_selection_changed)

    def load_themes(self):
        """Charge les thèmes depuis la base de données"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT id, nom FROM themes ORDER BY nom")
        themes = cursor.fetchall()
        conn.close()

        self.theme_list_widget.clear()

        for theme_id, name in themes:
            item = QListWidgetItem(name)
            item.setData(Qt.ItemDataRole.UserRole, theme_id)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.theme_list_widget.addItem(item)
        
        # Connecter le signal pour recharger les années quand on coche/décoche des thèmes
        self.theme_list_widget.itemChanged.connect(self.on_project_selection_changed)
        
        # Connecter le signal pour recharger les années quand on coche/décoche des thèmes principaux
        self.main_theme_list_widget.itemChanged.connect(self.on_project_selection_changed)

    def load_main_themes(self):
        """Charge les thèmes principaux depuis la base de données"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT theme_principal FROM projets WHERE theme_principal IS NOT NULL AND theme_principal != '' ORDER BY theme_principal")
        main_themes = cursor.fetchall()
        conn.close()

        self.main_theme_list_widget.clear()

        for (theme_name,) in main_themes:
            item = QListWidgetItem(theme_name)
            item.setData(Qt.ItemDataRole.UserRole, theme_name)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.main_theme_list_widget.addItem(item)

    def on_project_changed(self):
        """Appelée quand le projet sélectionné change dans le ComboBox."""
        # Mettre à jour les années si on est en mode "Projet spécifique"
        if self.radio_single_project.isChecked():
            if self.period_type.currentIndex() == 0:  # Année spécifique
                self.update_years_for_project()
            elif self.period_type.currentIndex() == 1:  # Plusieurs années
                self.update_years_for_single_project_multiple_years()

    def update_years_for_single_project_multiple_years(self):
        """Met à jour la liste des années cochables pour un projet spécifique en mode 'plusieurs années'."""
        project_id = self.project_combo.currentData()
        if not project_id:
            QMessageBox.warning(self, "Erreur", "Aucun projet sélectionné.")
            self.years_list_widget.clear()
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id = ?", (project_id,))
            result = cursor.fetchone()
            if not result:
                QMessageBox.warning(self, "Erreur", "Impossible de récupérer les dates pour le projet sélectionné.")
                self.years_list_widget.clear()
                return

            date_debut, date_fin = result
            if not date_debut or not date_fin:
                QMessageBox.warning(self, "Erreur", "Les dates de début ou de fin sont manquantes pour ce projet.")
                self.years_list_widget.clear()
                return

            # Extraire les années à partir des dates au format MM/AAAA
            try:
                debut_annee = int(date_debut.split("/")[1])
                fin_annee = int(date_fin.split("/")[1])
            except (IndexError, ValueError):
                QMessageBox.warning(self, "Erreur", "Les dates de début ou de fin sont mal formatées (attendu: MM/AAAA).")
                self.years_list_widget.clear()
                return

            # Remplir la liste des années cochables
            self.years_list_widget.clear()
            for year in range(debut_annee, fin_annee + 1):
                item = QListWidgetItem(str(year))
                item.setCheckState(Qt.CheckState.Unchecked)
                self.years_list_widget.addItem(item)

        finally:
            conn.close()

    def update_years_for_all_projects(self):
        """Met à jour la liste déroulante des années avec toutes les années où il y a au moins un projet."""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE date_debut IS NOT NULL AND date_fin IS NOT NULL")
            results = cursor.fetchall()
            
            if not results:
                QMessageBox.warning(self, "Aucun projet trouvé", "Aucun projet avec des dates valides n'a été trouvé.")
                self.year_combo.clear()
                return

            years_set = set()
            
            for date_debut, date_fin in results:
                try:
                    # Extraire les années à partir des dates au format MM/AAAA
                    debut_annee = int(date_debut.split("/")[1])
                    fin_annee = int(date_fin.split("/")[1])
                    
                    # Ajouter toutes les années entre début et fin (incluses)
                    for year in range(debut_annee, fin_annee + 1):
                        years_set.add(year)
                        
                except (IndexError, ValueError):
                    # Si la conversion échoue, on ignore cette entrée
                    continue

            years = sorted(list(years_set))
            
            if not years:
                QMessageBox.warning(self, "Aucune année trouvée", "Aucune année valide n'a été trouvée dans les projets.")
                self.year_combo.clear()
                return

            # Remplir la liste déroulante des années
            self.year_combo.clear()
            for year in years:
                self.year_combo.addItem(str(year))

        finally:
            conn.close()

    def update_years_for_selected_projects(self):
        """Met à jour la liste déroulante des années avec les années des projets sélectionnés."""
        # Récupérer les IDs des projets cochés
        selected_project_ids = []
        for i in range(self.project_list_widget.count()):
            item = self.project_list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_project_ids.append(item.data(Qt.ItemDataRole.UserRole))
        
        if not selected_project_ids:
            # Aucun projet sélectionné, vider la liste
            self.year_combo.clear()
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            # Récupérer les dates des projets sélectionnés
            placeholders = ','.join('?' * len(selected_project_ids))
            cursor.execute(f"""
                SELECT date_debut, date_fin FROM projets 
                WHERE id IN ({placeholders}) 
                AND date_debut IS NOT NULL AND date_fin IS NOT NULL
            """, selected_project_ids)
            
            results = cursor.fetchall()
            
            if not results:
                QMessageBox.warning(self, "Aucun projet trouvé", "Aucun projet sélectionné avec des dates valides n'a été trouvé.")
                self.year_combo.clear()
                return

            years_set = set()
            
            for date_debut, date_fin in results:
                try:
                    debut_annee = int(date_debut.split("/")[1])
                    fin_annee = int(date_fin.split("/")[1])
                    
                    for year in range(debut_annee, fin_annee + 1):
                        years_set.add(year)
                        
                except (IndexError, ValueError):
                    continue

            years = sorted(list(years_set))
            
            if not years:
                QMessageBox.warning(self, "Aucune année trouvée", "Aucune année valide n'a été trouvée dans les projets sélectionnés.")
                self.year_combo.clear()
                return

            # Remplir la liste déroulante des années
            self.year_combo.clear()
            for year in years:
                self.year_combo.addItem(str(year))

        finally:
            conn.close()

    def update_years_for_selected_themes(self):
        """Met à jour la liste déroulante des années avec les années des projets liés aux thèmes sélectionnés."""
        # Récupérer les IDs des thèmes cochés
        selected_theme_ids = []
        for i in range(self.theme_list_widget.count()):
            item = self.theme_list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_theme_ids.append(item.data(Qt.ItemDataRole.UserRole))
        
        if not selected_theme_ids:
            # Aucun thème sélectionné, vider la liste
            self.year_combo.clear()
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            # Récupérer les projets liés aux thèmes sélectionnés
            placeholders = ','.join('?' * len(selected_theme_ids))
            cursor.execute(f"""
                SELECT DISTINCT p.date_debut, p.date_fin 
                FROM projets p
                JOIN projet_themes pt ON p.id = pt.projet_id
                WHERE pt.theme_id IN ({placeholders})
                AND p.date_debut IS NOT NULL AND p.date_fin IS NOT NULL
            """, selected_theme_ids)
            
            results = cursor.fetchall()
            
            if not results:
                QMessageBox.warning(self, "Aucun projet trouvé", "Aucun projet lié aux thèmes sélectionnés avec des dates valides n'a été trouvé.")
                self.year_combo.clear()
                return

            years_set = set()
            
            for date_debut, date_fin in results:
                try:
                    debut_annee = int(date_debut.split("/")[1])
                    fin_annee = int(date_fin.split("/")[1])
                    
                    for year in range(debut_annee, fin_annee + 1):
                        years_set.add(year)
                        
                except (IndexError, ValueError):
                    continue

            years = sorted(list(years_set))
            
            if not years:
                QMessageBox.warning(self, "Aucune année trouvée", "Aucune année valide n'a été trouvée dans les projets liés aux thèmes sélectionnés.")
                self.year_combo.clear()
                return

            # Remplir la liste déroulante des années
            self.year_combo.clear()
            for year in years:
                self.year_combo.addItem(str(year))

        finally:
            conn.close()

    def update_years_for_selected_themes_multiple(self):
        """Met à jour la liste des années cochables avec les années des projets liés aux thèmes sélectionnés."""
        # Récupérer les IDs des thèmes cochés
        selected_theme_ids = []
        for i in range(self.theme_list_widget.count()):
            item = self.theme_list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_theme_ids.append(item.data(Qt.ItemDataRole.UserRole))
        
        if not selected_theme_ids:
            # Aucun thème sélectionné, vider la liste
            self.years_list_widget.clear()
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            # Récupérer les projets liés aux thèmes sélectionnés
            placeholders = ','.join('?' * len(selected_theme_ids))
            cursor.execute(f"""
                SELECT DISTINCT p.date_debut, p.date_fin 
                FROM projets p
                JOIN projet_themes pt ON p.id = pt.projet_id
                WHERE pt.theme_id IN ({placeholders})
                AND p.date_debut IS NOT NULL AND p.date_fin IS NOT NULL
            """, selected_theme_ids)
            
            results = cursor.fetchall()
            
            if not results:
                QMessageBox.warning(self, "Aucun projet trouvé", "Aucun projet lié aux thèmes sélectionnés avec des dates valides n'a été trouvé.")
                self.years_list_widget.clear()
                return

            years_set = set()
            
            for date_debut, date_fin in results:
                try:
                    debut_annee = int(date_debut.split("/")[1])
                    fin_annee = int(date_fin.split("/")[1])
                    
                    for year in range(debut_annee, fin_annee + 1):
                        years_set.add(year)
                        
                except (IndexError, ValueError):
                    continue

            years = sorted(list(years_set))
            
            if not years:
                QMessageBox.warning(self, "Aucune année trouvée", "Aucune année valide n'a été trouvée dans les projets liés aux thèmes sélectionnés.")
                self.years_list_widget.clear()
                return

            # Remplir la liste des années cochables
            self.years_list_widget.clear()
            for year in years:
                item = QListWidgetItem(str(year))
                item.setCheckState(Qt.CheckState.Unchecked)
                self.years_list_widget.addItem(item)

        finally:
            conn.close()

    def update_years_for_selected_main_themes(self):
        """Met à jour la liste déroulante des années avec les années des projets liés aux thèmes principaux sélectionnés."""
        # Récupérer les thèmes principaux cochés
        selected_main_themes = []
        for i in range(self.main_theme_list_widget.count()):
            item = self.main_theme_list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_main_themes.append(item.data(Qt.ItemDataRole.UserRole))
        
        if not selected_main_themes:
            # Aucun thème principal sélectionné, vider la liste
            self.year_combo.clear()
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            # Récupérer les projets avec les thèmes principaux sélectionnés
            placeholders = ','.join('?' * len(selected_main_themes))
            cursor.execute(f"""
                SELECT date_debut, date_fin 
                FROM projets 
                WHERE theme_principal IN ({placeholders})
                AND date_debut IS NOT NULL AND date_fin IS NOT NULL
            """, selected_main_themes)
            
            results = cursor.fetchall()
            
            if not results:
                QMessageBox.warning(self, "Aucun projet trouvé", "Aucun projet avec les thèmes principaux sélectionnés et des dates valides n'a été trouvé.")
                self.year_combo.clear()
                return

            years_set = set()
            
            for date_debut, date_fin in results:
                try:
                    debut_annee = int(date_debut.split("/")[1])
                    fin_annee = int(date_fin.split("/")[1])
                    
                    for year in range(debut_annee, fin_annee + 1):
                        years_set.add(year)
                        
                except (IndexError, ValueError):
                    continue

            years = sorted(list(years_set))
            
            if not years:
                QMessageBox.warning(self, "Aucune année trouvée", "Aucune année valide n'a été trouvée dans les projets avec les thèmes principaux sélectionnés.")
                self.year_combo.clear()
                return

            # Remplir la liste déroulante des années
            self.year_combo.clear()
            for year in years:
                self.year_combo.addItem(str(year))

        finally:
            conn.close()

    def update_years_for_selected_main_themes_multiple(self):
        """Met à jour la liste des années cochables avec les années des projets liés aux thèmes principaux sélectionnés."""
        # Récupérer les thèmes principaux cochés
        selected_main_themes = []
        for i in range(self.main_theme_list_widget.count()):
            item = self.main_theme_list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_main_themes.append(item.data(Qt.ItemDataRole.UserRole))
        
        if not selected_main_themes:
            # Aucun thème principal sélectionné, vider la liste
            self.years_list_widget.clear()
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            # Récupérer les projets avec les thèmes principaux sélectionnés
            placeholders = ','.join('?' * len(selected_main_themes))
            cursor.execute(f"""
                SELECT date_debut, date_fin 
                FROM projets 
                WHERE theme_principal IN ({placeholders})
                AND date_debut IS NOT NULL AND date_fin IS NOT NULL
            """, selected_main_themes)
            
            results = cursor.fetchall()
            
            if not results:
                QMessageBox.warning(self, "Aucun projet trouvé", "Aucun projet avec les thèmes principaux sélectionnés et des dates valides n'a été trouvé.")
                self.years_list_widget.clear()
                return

            years_set = set()
            
            for date_debut, date_fin in results:
                try:
                    debut_annee = int(date_debut.split("/")[1])
                    fin_annee = int(date_fin.split("/")[1])
                    
                    for year in range(debut_annee, fin_annee + 1):
                        years_set.add(year)
                        
                except (IndexError, ValueError):
                    continue

            years = sorted(list(years_set))
            
            if not years:
                QMessageBox.warning(self, "Aucune année trouvée", "Aucune année valide n'a été trouvée dans les projets avec les thèmes principaux sélectionnés.")
                self.years_list_widget.clear()
                return

            # Remplir la liste des années cochables
            self.years_list_widget.clear()
            for year in years:
                item = QListWidgetItem(str(year))
                item.setCheckState(Qt.CheckState.Unchecked)
                self.years_list_widget.addItem(item)

        finally:
            conn.close()

    def on_project_selection_changed(self):
        """Appelée quand la sélection de projets ou thèmes change"""
        # Mettre à jour les années selon le mode et le type de période
        if self.radio_multiple_projects.isChecked() and self.period_type.currentIndex() == 0:
            # Mode "Sélection de certains projets" + "Année spécifique"
            self.update_years_for_selected_projects()
        elif self.radio_by_theme.isChecked() and self.period_type.currentIndex() == 0:
            # Mode "Sélection par thèmes" + "Année spécifique"
            self.update_years_for_selected_themes()
        elif self.radio_by_theme.isChecked() and self.period_type.currentIndex() == 1:
            # Mode "Sélection par thèmes" + "Plusieurs années"
            self.update_years_for_selected_themes_multiple()
        elif self.radio_by_main_theme.isChecked() and self.period_type.currentIndex() == 0:
            # Mode "Sélection par le thème principal" + "Année spécifique"
            self.update_years_for_selected_main_themes()
        elif self.radio_by_main_theme.isChecked() and self.period_type.currentIndex() == 1:
            # Mode "Sélection par le thème principal" + "Plusieurs années"
            self.update_years_for_selected_main_themes_multiple()
        elif self.period_type.currentIndex() == 1:
            # Mode "Plusieurs années" pour les autres cas
            self.load_years_with_projects()

    def get_selected_project_ids(self):
        """Retourne la liste des IDs des projets sélectionnés selon l'option choisie"""
        project_ids = []
        
        if self.radio_single_project.isChecked():
            # Un projet spécifique
            project_id = self.project_combo.currentData()
            if project_id:
                project_ids = [project_id]
                
        elif self.radio_all_projects.isChecked():
            # Tous les projets
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM projets")
            project_ids = [row[0] for row in cursor.fetchall()]
            conn.close()
            
        elif self.radio_multiple_projects.isChecked():
            # Projets sélectionnés dans la liste
            for i in range(self.project_list_widget.count()):
                item = self.project_list_widget.item(i)
                if item.checkState() == Qt.CheckState.Checked:
                    project_ids.append(item.data(Qt.ItemDataRole.UserRole))
                    
        elif self.radio_by_theme.isChecked():
            # Projets par thèmes sélectionnés
            selected_theme_ids = []
            for i in range(self.theme_list_widget.count()):
                item = self.theme_list_widget.item(i)
                if item.checkState() == Qt.CheckState.Checked:
                    selected_theme_ids.append(item.data(Qt.ItemDataRole.UserRole))
            
            if selected_theme_ids:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                placeholders = ','.join('?' * len(selected_theme_ids))
                cursor.execute(f"""
                    SELECT DISTINCT projet_id FROM projet_themes 
                    WHERE theme_id IN ({placeholders})
                """, selected_theme_ids)
                project_ids = [row[0] for row in cursor.fetchall()]
                conn.close()
        elif self.radio_by_main_theme.isChecked():
            # Projets par thème principal sélectionné
            selected_main_themes = []
            for i in range(self.main_theme_list_widget.count()):
                item = self.main_theme_list_widget.item(i)
                if item.checkState() == Qt.CheckState.Checked:
                    selected_main_themes.append(item.data(Qt.ItemDataRole.UserRole))
            
            if selected_main_themes:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                placeholders = ','.join('?' * len(selected_main_themes))
                cursor.execute(f"""
                    SELECT id FROM projets 
                    WHERE theme_principal IN ({placeholders})
                """, selected_main_themes)
                project_ids = [row[0] for row in cursor.fetchall()]
                conn.close()
        
        return project_ids

    def load_years_with_projects(self):
        """Charge les années où des projets sélectionnés existent depuis la base de données"""
        project_ids = self.get_selected_project_ids()
        
        if not project_ids:
            # Ne pas afficher de message d'erreur, juste vider la liste
            self.years_list_widget.clear()
            return
            
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            # Récupérer les projets sélectionnés avec leurs dates
            placeholders = ','.join('?' * len(project_ids))
            cursor.execute(f"""
                SELECT date_debut, date_fin FROM projets 
                WHERE id IN ({placeholders})
                AND date_debut IS NOT NULL AND date_debut LIKE '__/____'
            """, project_ids)
            
            years_set = set()
            
            for date_debut, date_fin in cursor.fetchall():
                # Extraire l'année de début
                if date_debut and len(date_debut) >= 7:
                    year = int(date_debut[-4:])
                    years_set.add(year)

            years = sorted(list(years_set))

            if not years:
                QMessageBox.warning(self, "Aucune année trouvée", "Aucune année avec des projets sélectionnés n'a été trouvée.")

            self.years_list_widget.clear()
            for year in years:
                item = QListWidgetItem(str(year))
                item.setCheckState(Qt.CheckState.Unchecked)
                self.years_list_widget.addItem(item)
                
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors du chargement des années : {str(e)}")
        finally:
            conn.close()

    def update_years_for_project(self):
        """Met à jour la liste des années disponibles pour le projet sélectionné."""
        project_id = self.project_combo.currentData()
        if not project_id:
            QMessageBox.warning(self, "Erreur", "Aucun projet sélectionné.")
            self.year_combo.clear()
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id = ?", (project_id,))
            result = cursor.fetchone()
            if not result:
                QMessageBox.warning(self, "Erreur", "Impossible de récupérer les dates pour le projet sélectionné.")
                self.year_combo.clear()
                return

            date_debut, date_fin = result
            if not date_debut or not date_fin:
                QMessageBox.warning(self, "Erreur", "Les dates de début ou de fin sont manquantes pour ce projet.")
                self.year_combo.clear()
                return

            # Extraire les années à partir des dates au format MM/AAAA
            try:
                debut_annee = int(date_debut.split("/")[1])
                fin_annee = int(date_fin.split("/")[1])
            except (IndexError, ValueError):
                QMessageBox.warning(self, "Erreur", "Les dates de début ou de fin sont mal formatées (attendu: MM/AAAA).")
                self.year_combo.clear()
                return

            # Remplir la liste déroulante des années
            self.year_combo.clear()
            for year in range(debut_annee, fin_annee + 1):
                self.year_combo.addItem(str(year))

        finally:
            conn.close()

    def update_period_widgets(self):
        """Met à jour l'affichage selon le type de période sélectionné"""
        period_type = self.period_type.currentIndex()

        if period_type == 0:  # Année spécifique
            self.year_combo.show()
            self.years_list_widget.hide()
            # Mettre à jour les années selon le type de sélection
            if self.radio_single_project.isChecked():
                self.update_years_for_project()
            elif self.radio_all_projects.isChecked():
                self.update_years_for_all_projects()
            elif self.radio_multiple_projects.isChecked():
                self.update_years_for_selected_projects()
            elif self.radio_by_theme.isChecked():
                self.update_years_for_selected_themes()
            elif self.radio_by_main_theme.isChecked():
                self.update_years_for_selected_main_themes()
        elif period_type == 1:  # Plusieurs années
            self.year_combo.hide()
            self.years_list_widget.show()
            # Recharger les années selon le type de sélection
            if self.radio_single_project.isChecked():
                self.update_years_for_single_project_multiple_years()
            elif self.radio_by_theme.isChecked():
                self.update_years_for_selected_themes_multiple()
            elif self.radio_by_main_theme.isChecked():
                self.update_years_for_selected_main_themes_multiple()
            else:
                self.load_years_with_projects()
        else:  # Période globale du projet
            self.year_combo.hide()
            self.years_list_widget.hide()

    def update_selection_widgets(self):
        """Met à jour les widgets de sélection en fonction de l'option choisie"""
        if self.radio_single_project.isChecked():
            self.project_combo.show()
            self.project_list_widget.hide()
            self.theme_list_widget.hide()
            self.main_theme_list_widget.hide()

            # Pré-sélectionner le projet courant si défini
            if self.projet_id is not None:
                index = self.project_combo.findData(self.projet_id)
                if index != -1:
                    self.project_combo.setCurrentIndex(index)

            # Mettre à jour les années si "Année spécifique" est sélectionné
            if self.period_type.currentIndex() == 0:  # Année spécifique
                self.update_years_for_project()

        elif self.radio_all_projects.isChecked():
            self.project_combo.hide()
            self.project_list_widget.hide()
            self.theme_list_widget.hide()
            self.main_theme_list_widget.hide()
        elif self.radio_multiple_projects.isChecked():
            self.project_combo.hide()
            self.project_list_widget.show()
            self.theme_list_widget.hide()
            self.main_theme_list_widget.hide()
        elif self.radio_by_theme.isChecked():
            self.project_combo.hide()
            self.project_list_widget.hide()
            self.theme_list_widget.show()
            self.main_theme_list_widget.hide()
        elif self.radio_by_main_theme.isChecked():
            self.project_combo.hide()
            self.project_list_widget.hide()
            self.theme_list_widget.hide()
            self.main_theme_list_widget.show()
        
        # Mettre à jour les années selon le type de sélection si "Tous les projets" est sélectionné
        if self.radio_all_projects.isChecked():
            # Mettre à jour les années si on est en mode "Année spécifique"
            if self.period_type.currentIndex() == 0:  # Année spécifique
                self.update_years_for_all_projects()

        # Recharger les années si on est en mode "Plusieurs années" et seulement pour "Tous les projets"
        if self.period_type.currentIndex() == 1 and self.radio_all_projects.isChecked():
            self.load_years_with_projects()

    def get_selected_project_id(self):
        """Retourne l'ID du projet sélectionné"""
        if self.radio_single_project.isChecked():
            return self.project_combo.currentData()
        elif self.radio_multiple_projects.isChecked():
            # Récupérer les projets sélectionnés dans la liste
            selected_items = self.project_list_widget.selectedItems()
            return [item.data(Qt.ItemDataRole.UserRole) for item in selected_items]
        else:
            return None

    def get_selected_theme_ids(self):
        """Retourne les IDs des thèmes sélectionnés"""
        if self.radio_by_theme.isChecked():
            # Récupérer les thèmes sélectionnés dans la liste
            selected_items = self.theme_list_widget.selectedItems()
            return [item.data(Qt.ItemDataRole.UserRole) for item in selected_items]
        else:
            return None

    def get_selected_period(self):
        """Retourne les informations de période sélectionnée"""
        project_id = self.get_selected_project_id()
        year = int(self.year_combo.currentText()) if self.year_combo.currentText() else None
        
        if self.period_type.currentIndex() == 0:  # Année spécifique
            return project_id, year, None
        elif self.period_type.currentIndex() == 1:  # Plusieurs années
            # Récupérer les années sélectionnées dans la liste
            selected_items = self.years_list_widget.selectedItems()
            years = [int(item.text()) for item in selected_items]
            return project_id, years, None
        else:  # Période globale du projet
            return project_id, None, "complete_project"
    
    def generate_compte_resultat(self):
        """Génère le compte de résultat"""
        # Récupérer les IDs des projets sélectionnés
        project_ids = self.get_selected_project_ids()
        
        if not project_ids:
            QMessageBox.warning(self, "Avertissement", "Veuillez sélectionner au moins un projet.")
            return
        
        # Récupérer les années sélectionnées
        years = self.get_selected_years()
        
        if not years:
            QMessageBox.warning(self, "Avertissement", "Veuillez sélectionner au moins une année.")
            return
        
        # Déterminer la granularité
        granularity = "monthly" if self.radio_monthly.isChecked() else "yearly"
        
        # Préparer la configuration pour le compte de résultat
        config_data = {
            'project_ids': project_ids,
            'years': years,
            'granularity': granularity,
            'period_type': self.get_period_type(),
            'cost_type': self.get_cost_type()
        }
        
        # Importer et afficher le compte de résultat
        try:
            from compte_resultat_display import show_compte_resultat
            show_compte_resultat(self.parent, config_data)
            self.accept()
        except ImportError as e:
            QMessageBox.critical(self, "Erreur", f"Impossible d'importer le module compte_resultat_display: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'affichage du compte de résultat: {str(e)}")
    
    def get_selected_years(self):
        """Retourne la liste des années sélectionnées"""
        if self.period_type.currentIndex() == 0:  # Année spécifique
            year_text = self.year_combo.currentText()
            if year_text:
                return [int(year_text)]
        elif self.period_type.currentIndex() == 1:  # Plusieurs années
            years = []
            for i in range(self.years_list_widget.count()):
                item = self.years_list_widget.item(i)
                if item.checkState() == Qt.CheckState.Checked:
                    years.append(int(item.text()))
            return years
        else:  # Période globale du projet
            # Récupérer toutes les années du projet
            return self.get_all_project_years()
        
        return []
    
    def get_all_project_years(self):
        """Récupère toutes les années des projets sélectionnés"""
        project_ids = self.get_selected_project_ids()
        if not project_ids:
            return []
        
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            placeholders = ','.join('?' * len(project_ids))
            cursor.execute(f"""
                SELECT date_debut, date_fin FROM projets 
                WHERE id IN ({placeholders})
                AND date_debut IS NOT NULL AND date_fin IS NOT NULL
            """, project_ids)
            
            years_set = set()
            for date_debut, date_fin in cursor.fetchall():
                try:
                    debut_annee = int(date_debut.split("/")[1])
                    fin_annee = int(date_fin.split("/")[1])
                    
                    for year in range(debut_annee, fin_annee + 1):
                        years_set.add(year)
                        
                except (IndexError, ValueError):
                    continue
            
            return sorted(list(years_set))
            
        finally:
            conn.close()
    
    def get_period_type(self):
        """Retourne le type de période sélectionné"""
        if self.period_type.currentIndex() == 0:
            return "specific_year"
        elif self.period_type.currentIndex() == 1:
            return "multiple_years"
        else:
            return "complete_project"
    
    def get_cost_type(self):
        """Retourne le type de coût sélectionné"""
        if self.radio_montant_charge.isChecked():
            return "montant_charge"
        elif self.radio_cout_production.isChecked():
            return "cout_production"
        elif self.radio_cout_complet.isChecked():
            return "cout_complet"
        else:
            return "cout_production"  # Par défaut

def show_print_config_dialog(parent, projet_id=None):
    """Fonction d'entrée pour afficher la configuration d'impression"""
    dialog = PrintConfigDialog(parent, projet_id)
    return dialog.exec()
