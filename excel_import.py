from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QTableWidget, QTableWidgetItem, QComboBox, QFileDialog,
    QTabWidget, QWidget, QLineEdit, QSpinBox, QCheckBox,
    QMessageBox, QTextEdit, QListWidget, QSplitter, QGroupBox,
    QScrollArea, QProgressBar, QTreeWidget, QTreeWidgetItem, QInputDialog,
    QFormLayout, QGridLayout
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
import pandas as pd
import json
import os
import sqlite3
from datetime import datetime
import re

class ExcelImportDialog(QDialog):
    def __init__(self, projet_id, parent=None):
        super().__init__(parent)
        self.projet_id = projet_id
        self.setWindowTitle("Import Excel - Configuration universelle")
        self.setMinimumSize(1400, 900)
        
        # Données actuelles
        self.excel_file = None
        self.excel_sheets = {}
        self.current_config = {}
        self.configs_dir = "import_configs"
        
        # Créer le dossier de configs s'il n'existe pas
        if not os.path.exists(self.configs_dir):
            os.makedirs(self.configs_dir)
        
        layout = QVBoxLayout()
        
        # === TITRE ===
        title = QLabel("Configuration d'import Excel")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # === ÉTAPE 1: Gestion des profils ===
        profile_group = QGroupBox("1. Profils de configuration")
        profile_layout = QHBoxLayout()
        
        profile_layout.addWidget(QLabel("Profil actuel:"))
        self.profile_combo = QComboBox()
        self.load_available_profiles()
        profile_layout.addWidget(self.profile_combo)
        
        btn_new_profile = QPushButton("Nouveau")
        btn_load_profile = QPushButton("Charger")
        btn_save_profile = QPushButton("Sauvegarder")
        btn_delete_profile = QPushButton("Supprimer")
        
        profile_layout.addWidget(btn_new_profile)
        profile_layout.addWidget(btn_load_profile)
        profile_layout.addWidget(btn_save_profile)
        profile_layout.addWidget(btn_delete_profile)
        
        profile_group.setLayout(profile_layout)
        layout.addWidget(profile_group)
        
        # === ÉTAPE 2: Sélection fichier ===
        file_group = QGroupBox("2. Sélection du fichier Excel")
        file_layout = QVBoxLayout()
        
        file_select_layout = QHBoxLayout()
        self.file_label = QLabel("Aucun fichier sélectionné")
        btn_select = QPushButton("Parcourir...")
        btn_select.clicked.connect(self.select_excel_file)
        file_select_layout.addWidget(self.file_label)
        file_select_layout.addWidget(btn_select)
        file_layout.addLayout(file_select_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # === ÉTAPE 3: Configuration par type de données ===
        config_group = QGroupBox("3. Configuration du mapping")
        config_layout = QVBoxLayout()
        
        # Onglets pour chaque type
        self.tabs = QTabWidget()
        
        # Chaque onglet = un type de données à importer
        self.mapping_widgets = {
            'temps_travail': MatrixMappingWidget('temps_travail', self, self.projet_id),
            'recettes': UniversalMappingWidget('recettes', self),
            'depenses': UniversalMappingWidget('depenses', self),
            'autres_depenses': UniversalMappingWidget('autres_depenses', self)
        }
        
        self.tabs.addTab(self.mapping_widgets['temps_travail'], "Temps de travail (Matrice)")
        self.tabs.addTab(self.mapping_widgets['recettes'], "Recettes")
        self.tabs.addTab(self.mapping_widgets['depenses'], "Dépenses")
        self.tabs.addTab(self.mapping_widgets['autres_depenses'], "Autres dépenses")
        
        config_layout.addWidget(self.tabs)
        config_group.setLayout(config_layout)
        layout.addWidget(config_group)
        
        # === BOUTONS ===
        btn_layout = QHBoxLayout()
        
        btn_preview = QPushButton("Aperçu données")
        btn_import = QPushButton("IMPORTER")
        btn_close = QPushButton("Fermer")
        
        btn_preview.clicked.connect(self.preview_data)
        btn_import.clicked.connect(self.import_data)
        btn_close.clicked.connect(self.close)
        
        # Style pour le bouton d'import
        btn_import.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 8px; }")
        btn_preview.setStyleSheet("QPushButton { background-color: #2196F3; color: white; padding: 8px; }")
        
        btn_layout.addWidget(btn_preview)
        btn_layout.addWidget(btn_import)
        btn_layout.addWidget(btn_close)
        
        layout.addLayout(btn_layout)
        
        # === Barre de progression ===
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        self.setLayout(layout)
        
        # Connexions
        btn_new_profile.clicked.connect(self.new_profile)
        btn_load_profile.clicked.connect(self.load_profile)
        btn_save_profile.clicked.connect(self.save_profile)
        btn_delete_profile.clicked.connect(self.delete_profile)
    
    def load_available_profiles(self):
        """Charge la liste des profils disponibles"""
        self.profile_combo.clear()
        self.profile_combo.addItem("-- Nouveau profil --")
        
        if os.path.exists(self.configs_dir):
            for file in os.listdir(self.configs_dir):
                if file.endswith('.json'):
                    profile_name = file[:-5]  # Enlève .json
                    self.profile_combo.addItem(profile_name)
    
    def new_profile(self):
        """Crée un nouveau profil"""
        self.profile_combo.setCurrentText("-- Nouveau profil --")
        self.current_config = {}
        
        # Reset tous les widgets de mapping
        for widget in self.mapping_widgets.values():
            widget.reset_configuration()
    
    def load_profile(self):
        """Charge un profil existant"""
        profile_name = self.profile_combo.currentText()
        if profile_name == "-- Nouveau profil --":
            return
        
        config_file = os.path.join(self.configs_dir, f"{profile_name}.json")
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    self.current_config = json.load(f)
                
                # Applique la config aux widgets
                for data_type, widget in self.mapping_widgets.items():
                    if data_type in self.current_config:
                        widget.load_configuration(self.current_config[data_type])
                
                QMessageBox.information(self, "Succès", f"Profil '{profile_name}' chargé avec succès !")
                
            except Exception as e:
                QMessageBox.warning(self, "Erreur", f"Impossible de charger le profil: {e}")
    
    def save_profile(self):
        """Sauvegarde le profil actuel"""
        profile_name = self.profile_combo.currentText()
        if profile_name == "-- Nouveau profil --":
            profile_name, ok = QInputDialog.getText(self, "Nouveau profil", "Nom du profil:")
            if not ok or not profile_name.strip():
                return
            profile_name = profile_name.strip()
        
        # Collecte la config de tous les widgets
        config = {}
        for data_type, widget in self.mapping_widgets.items():
            config[data_type] = widget.get_configuration()
        
        config_file = os.path.join(self.configs_dir, f"{profile_name}.json")
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            self.load_available_profiles()
            self.profile_combo.setCurrentText(profile_name)
            QMessageBox.information(self, "Succès", f"Profil '{profile_name}' sauvegardé !")
            
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Impossible de sauvegarder: {e}")
    
    def delete_profile(self):
        """Supprime un profil"""
        profile_name = self.profile_combo.currentText()
        if profile_name == "-- Nouveau profil --":
            QMessageBox.warning(self, "Erreur", "Sélectionnez un profil à supprimer")
            return
        
        reply = QMessageBox.question(self, "Confirmation", 
                                   f"Supprimer le profil '{profile_name}' ?",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            config_file = os.path.join(self.configs_dir, f"{profile_name}.json")
            try:
                os.remove(config_file)
                self.load_available_profiles()
                QMessageBox.information(self, "Succès", "Profil supprimé !")
            except Exception as e:
                QMessageBox.warning(self, "Erreur", f"Impossible de supprimer: {e}")
    
    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Sélectionner fichier Excel", "", "Fichiers Excel (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_file = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.load_excel_sheets()
    
    def load_excel_sheets(self):
        """Charge toutes les feuilles du fichier Excel"""
        try:
            # Lit toutes les feuilles (juste les 10 premières lignes pour l'aperçu)
            self.excel_sheets = pd.read_excel(self.excel_file, sheet_name=None, nrows=10)
            
            # Met à jour les widgets de mapping
            for widget in self.mapping_widgets.values():
                widget.update_excel_data(self.excel_sheets)
                
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Impossible de lire le fichier Excel: {e}")
    
    def preview_data(self):
        """Affiche un aperçu des données qui seront importées"""
        if not self.excel_file:
            QMessageBox.warning(self, "Erreur", "Sélectionnez d'abord un fichier Excel")
            return
        
        # Ouvre la dialog d'aperçu
        preview_dialog = PreviewDialog(self, self.excel_file, self.mapping_widgets, self.projet_id)
        preview_dialog.exec()
    
    def import_data(self):
        """Lance l'import des données"""
        if not self.excel_file:
            QMessageBox.warning(self, "Erreur", "Sélectionnez d'abord un fichier Excel")
            return
        
        # Vérifie qu'au moins un mapping est configuré
        has_mapping = False
        for widget in self.mapping_widgets.values():
            if widget.has_valid_mapping():
                has_mapping = True
                break
        
        if not has_mapping:
            QMessageBox.warning(self, "Erreur", "Configurez au moins un mapping valide")
            return
        
        reply = QMessageBox.question(self, "Confirmation", 
                                   "Lancer l'import des données ?\n\nAttention: cela peut écraser des données existantes.",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self.run_import()
    
    def run_import(self):
        """Exécute l'import des données"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        try:
            total_imports = 0
            successful_imports = 0
            
            for data_type, widget in self.mapping_widgets.items():
                if widget.has_valid_mapping():
                    self.progress_bar.setLabelText(f"Import {data_type}...")
                    
                    success = widget.import_data_to_database(self.excel_file, self.projet_id)
                    total_imports += 1
                    if success:
                        successful_imports += 1
                    
                    self.progress_bar.setValue(int((total_imports / len(self.mapping_widgets)) * 100))
            
            self.progress_bar.setValue(100)
            QMessageBox.information(self, "Import terminé", 
                                  f"Import terminé !\n{successful_imports}/{total_imports} types importés avec succès.")
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'import", f"Erreur lors de l'import: {e}")
        
        finally:
            self.progress_bar.setVisible(False)


class UniversalMappingWidget(QWidget):
    def __init__(self, data_type, parent_dialog):
        super().__init__()
        self.data_type = data_type
        self.parent_dialog = parent_dialog
        self.excel_sheets = {}
        
        layout = QVBoxLayout()
        
        # === Configuration de base ===
        basic_config = QGroupBox("Configuration de base")
        basic_layout = QVBoxLayout()
        
        # Sélection feuille
        sheet_layout = QHBoxLayout()
        sheet_layout.addWidget(QLabel("Feuille Excel:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        sheet_layout.addWidget(self.sheet_combo)
        basic_layout.addLayout(sheet_layout)
        
        # Ligne d'en-tête
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("Ligne d'en-têtes:"))
        self.header_row = QSpinBox()
        self.header_row.setMinimum(0)
        self.header_row.setValue(0)
        self.header_row.valueChanged.connect(self.update_columns_preview)
        header_layout.addWidget(self.header_row)
        
        # Première ligne de données
        data_start_layout = QHBoxLayout()
        data_start_layout.addWidget(QLabel("Première ligne de données:"))
        self.data_start_row = QSpinBox()
        self.data_start_row.setMinimum(1)
        self.data_start_row.setValue(1)
        self.data_start_row.valueChanged.connect(self.update_columns_preview)  # Connexion ajoutée
        data_start_layout.addWidget(self.data_start_row)
        header_layout.addLayout(data_start_layout)
        
        basic_layout.addLayout(header_layout)
        basic_config.setLayout(basic_layout)
        layout.addWidget(basic_config)
        
        # === Mapping des colonnes ===
        mapping_config = QGroupBox("Mapping des colonnes")
        mapping_layout = QHBoxLayout()
        
        # Colonne gauche: Colonnes Excel disponibles
        left_panel = QVBoxLayout()
        left_panel.addWidget(QLabel("Colonnes Excel disponibles:"))
        self.excel_columns_list = QListWidget()
        left_panel.addWidget(self.excel_columns_list)
        
        # Preview des données
        left_panel.addWidget(QLabel("Aperçu des données:"))
        self.data_preview = QTableWidget()
        self.data_preview.setMaximumHeight(150)
        left_panel.addWidget(self.data_preview)
        
        # Colonne droite: Champs requis + mapping
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel("Mapping vers vos champs:"))
        
        self.mapping_table = QTableWidget()
        self.setup_mapping_table()
        right_panel.addWidget(self.mapping_table)
        
        # Ajout du splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        right_widget = QWidget()
        right_widget.setLayout(right_panel)
        
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([400, 600])
        
        mapping_layout.addWidget(splitter)
        mapping_config.setLayout(mapping_layout)
        layout.addWidget(mapping_config)
        
        # === Options avancées ===
        advanced_config = QGroupBox("Options avancées")
        advanced_layout = QVBoxLayout()
        
        # Filtres
        filter_layout = QVBoxLayout()
        filter_layout.addWidget(QLabel("Filtre des données (optionnel):"))
        self.filter_text = QTextEdit()
        self.filter_text.setMaximumHeight(60)
        self.filter_text.setPlaceholderText("Ex: Colonne_A == 'MonProjet' and Colonne_B >= 2024")
        filter_layout.addWidget(self.filter_text)
        advanced_layout.addLayout(filter_layout)
        
        advanced_config.setLayout(advanced_layout)
        layout.addWidget(advanced_config)
        
        self.setLayout(layout)
    
    def setup_mapping_table(self):
        """Configure le tableau de mapping selon le type de données"""
        # Définit les champs requis selon le type
        field_definitions = {
            'temps_travail': [
                ('Projet', True, 'Nom ou ID du projet'),
                ('Année', True, 'Année (format YYYY)'),
                ('Direction', False, 'Direction/Service'),
                ('Catégorie', False, 'Catégorie de personnel'),
                ('Mois', True, 'Mois (texte ou numéro)'),
                ('Jours', True, 'Nombre de jours travaillés')
            ],
            'recettes': [
                ('Projet', True, 'Nom ou ID du projet'),
                ('Année', True, 'Année (format YYYY)'),
                ('Catégorie', False, 'Type de recette'),
                ('Mois', True, 'Mois (texte ou numéro)'),
                ('Montant', True, 'Montant en euros'),
                ('Détail', False, 'Description/commentaire')
            ],
            'depenses': [
                ('Projet', True, 'Nom ou ID du projet'),
                ('Année', True, 'Année (format YYYY)'),
                ('Catégorie', False, 'Type de dépense'),
                ('Mois', True, 'Mois (texte ou numéro)'),
                ('Montant', True, 'Montant en euros'),
                ('Détail', False, 'Description/commentaire')
            ],
            'autres_depenses': [
                ('Projet', True, 'Nom ou ID du projet'),
                ('Année', True, 'Année (format YYYY)'),
                ('Ligne', False, 'Numéro de ligne'),
                ('Mois', True, 'Mois (texte ou numéro)'),
                ('Montant', True, 'Montant en euros'),
                ('Détail', False, 'Description/commentaire')
            ]
        }
        
        fields = field_definitions.get(self.data_type, [])
        
        self.mapping_table.setRowCount(len(fields))
        self.mapping_table.setColumnCount(4)
        self.mapping_table.setHorizontalHeaderLabels([
            "Champ requis", "Colonne Excel", "Obligatoire", "Description"
        ])
        
        for i, (field_name, required, description) in enumerate(fields):
            # Nom du champ (non éditable)
            field_item = QTableWidgetItem(field_name)
            field_item.setFlags(field_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.mapping_table.setItem(i, 0, field_item)
            
            # Combobox pour la colonne Excel
            combo = QComboBox()
            combo.addItem("-- Non mappé --")
            self.mapping_table.setCellWidget(i, 1, combo)
            
            # Checkbox obligatoire (non éditable pour les champs requis)
            checkbox = QCheckBox()
            checkbox.setChecked(required)
            if required:
                checkbox.setEnabled(False)
            self.mapping_table.setCellWidget(i, 2, checkbox)
            
            # Description
            desc_item = QTableWidgetItem(description)
            desc_item.setFlags(desc_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.mapping_table.setItem(i, 3, desc_item)
        
        # Redimensionne les colonnes
        self.mapping_table.setColumnWidth(0, 120)
        self.mapping_table.setColumnWidth(1, 150)
        self.mapping_table.setColumnWidth(2, 80)
        self.mapping_table.setColumnWidth(3, 200)
    
    def update_excel_data(self, excel_sheets):
        """Met à jour les données Excel disponibles"""
        self.excel_sheets = excel_sheets
        
        # Met à jour la liste des feuilles
        self.sheet_combo.clear()
        self.sheet_combo.addItems(list(excel_sheets.keys()))
        
        if excel_sheets:
            self.on_sheet_changed()
    
    def on_sheet_changed(self):
        """Appelé quand la feuille Excel change"""
        self.update_columns_preview()
    
    def update_columns_preview(self):
        """Met à jour l'aperçu des colonnes et données"""
        sheet_name = self.sheet_combo.currentText()
        if not sheet_name or sheet_name not in self.excel_sheets:
            return
        
        try:
            # Calcule les lignes à ignorer entre l'en-tête et le début des données
            header_row = self.header_row.value()
            data_start_row = self.data_start_row.value()
            
            # Si la ligne de données commence après l'en-tête, on ignore les lignes intermédiaires
            if data_start_row > header_row + 1:
                skip_rows = list(range(header_row + 1, data_start_row))
            else:
                skip_rows = None
            
            # Relit avec la bonne ligne d'en-tête et en ignorant les bonnes lignes
            df = pd.read_excel(
                self.parent_dialog.excel_file,
                sheet_name=sheet_name,
                header=header_row,
                skiprows=skip_rows,
                nrows=10  # Limite pour l'aperçu
            )            # Met à jour la liste des colonnes
            self.excel_columns_list.clear()
            columns = list(df.columns)
            self.excel_columns_list.addItems([str(col) for col in columns])
            
            # Met à jour les combobox de mapping
            for row in range(self.mapping_table.rowCount()):
                combo = self.mapping_table.cellWidget(row, 1)
                if combo:
                    current_value = combo.currentText()
                    combo.clear()
                    combo.addItem("-- Non mappé --")
                    combo.addItems([str(col) for col in columns])
                    
                    # Restaure la sélection si possible
                    if current_value in [str(col) for col in columns]:
                        combo.setCurrentText(current_value)
            
            # Met à jour l'aperçu des données
            self.update_data_preview(df)
            
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Erreur lors de la lecture: {e}")
    
    def update_data_preview(self, df):
        """Met à jour le tableau d'aperçu des données"""
        self.data_preview.setRowCount(min(5, len(df)))
        self.data_preview.setColumnCount(len(df.columns))
        self.data_preview.setHorizontalHeaderLabels([str(col) for col in df.columns])
        
        for row in range(min(5, len(df))):
            for col, column_name in enumerate(df.columns):
                value = str(df.iloc[row, col])
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.data_preview.setItem(row, col, item)
    
    def reset_configuration(self):
        """Remet à zéro la configuration"""
        self.sheet_combo.clear()
        self.header_row.setValue(0)
        self.data_start_row.setValue(1)
        self.filter_text.clear()
        
        # Reset mapping table
        for row in range(self.mapping_table.rowCount()):
            combo = self.mapping_table.cellWidget(row, 1)
            if combo:
                combo.setCurrentText("-- Non mappé --")
    
    def get_configuration(self):
        """Retourne la configuration actuelle"""
        config = {
            'sheet': self.sheet_combo.currentText(),
            'header_row': self.header_row.value(),
            'data_start_row': self.data_start_row.value(),
            'filter': self.filter_text.toPlainText(),
            'mappings': {}
        }
        
        for row in range(self.mapping_table.rowCount()):
            field_name = self.mapping_table.item(row, 0).text()
            combo = self.mapping_table.cellWidget(row, 1)
            checkbox = self.mapping_table.cellWidget(row, 2)
            
            if combo and checkbox:
                config['mappings'][field_name] = {
                    'excel_column': combo.currentText(),
                    'required': checkbox.isChecked()
                }
        
        return config
    
    def load_configuration(self, config):
        """Charge une configuration"""
        if 'sheet' in config:
            self.sheet_combo.setCurrentText(config['sheet'])
        if 'header_row' in config:
            self.header_row.setValue(config['header_row'])
        if 'data_start_row' in config:
            self.data_start_row.setValue(config['data_start_row'])
        if 'filter' in config:
            self.filter_text.setPlainText(config['filter'])
        
        if 'mappings' in config:
            for row in range(self.mapping_table.rowCount()):
                field_name = self.mapping_table.item(row, 0).text()
                if field_name in config['mappings']:
                    mapping = config['mappings'][field_name]
                    
                    combo = self.mapping_table.cellWidget(row, 1)
                    checkbox = self.mapping_table.cellWidget(row, 2)
                    
                    if combo and 'excel_column' in mapping:
                        combo.setCurrentText(mapping['excel_column'])
                    if checkbox and 'required' in mapping:
                        checkbox.setChecked(mapping['required'])
    
    def has_valid_mapping(self):
        """Vérifie si le mapping est valide"""
        if not self.sheet_combo.currentText():
            return False
        
        # Vérifie qu'au moins un champ obligatoire est mappé
        for row in range(self.mapping_table.rowCount()):
            checkbox = self.mapping_table.cellWidget(row, 2)
            combo = self.mapping_table.cellWidget(row, 1)
            
            if checkbox and checkbox.isChecked():
                if not combo or combo.currentText() == "-- Non mappé --":
                    return False
        
        return True
    
    def import_data_to_database(self, excel_file, projet_id):
        """Importe les données dans la base de données"""
        try:
            config = self.get_configuration()
            
            # Calcule les lignes à ignorer entre l'en-tête et le début des données
            header_row = config['header_row']
            data_start_row = config['data_start_row']
            
            if data_start_row > header_row + 1:
                skip_rows = list(range(header_row + 1, data_start_row))
            else:
                skip_rows = None
            
            # Lit toutes les données de la feuille
            df = pd.read_excel(
                excel_file,
                sheet_name=config['sheet'],
                header=header_row,
                skiprows=skip_rows
            )
            
            # Applique le filtre si spécifié
            if config['filter'].strip():
                try:
                    df = df.query(config['filter'])
                except Exception as e:
                    QMessageBox.warning(self, "Erreur de filtre", f"Filtre invalide: {e}")
                    return False
            
            # Prépare le mapping des colonnes
            column_mapping = {}
            for field, mapping in config['mappings'].items():
                if mapping['excel_column'] != "-- Non mappé --":
                    column_mapping[field] = mapping['excel_column']
            
            # Import selon le type de données
            return self._import_specific_data(df, column_mapping, projet_id)
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'import", f"Erreur lors de l'import {self.data_type}: {e}")
            return False
    
    def _import_specific_data(self, df, column_mapping, projet_id):
        """Importe les données spécifiques selon le type"""
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        
        try:
            imported_count = 0
            
            for _, row in df.iterrows():
                try:
                    if self.data_type == 'temps_travail':
                        self._import_temps_travail_row(cursor, row, column_mapping, projet_id)
                    elif self.data_type == 'recettes':
                        self._import_recettes_row(cursor, row, column_mapping, projet_id)
                    elif self.data_type == 'depenses':
                        self._import_depenses_row(cursor, row, column_mapping, projet_id)
                    elif self.data_type == 'autres_depenses':
                        self._import_autres_depenses_row(cursor, row, column_mapping, projet_id)
                    
                    imported_count += 1
                    
                except Exception as e:
                    print(f"Erreur ligne {_}: {e}")  # Log mais continue
            
            conn.commit()
            QMessageBox.information(self, "Succès", f"{imported_count} lignes importées pour {self.data_type}")
            return True
            
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()
    
    def _import_temps_travail_row(self, cursor, row, mapping, projet_id):
        """Importe une ligne de temps de travail"""
        # Extrait les valeurs selon le mapping
        annee = int(row[mapping['Année']]) if 'Année' in mapping else None
        direction = str(row[mapping['Direction']]) if 'Direction' in mapping else ""
        categorie = str(row[mapping['Catégorie']]) if 'Catégorie' in mapping else ""
        mois = str(row[mapping['Mois']]) if 'Mois' in mapping else ""
        jours = float(row[mapping['Jours']]) if 'Jours' in mapping else 0.0
        
        # Insert ou update
        cursor.execute("""
            INSERT OR REPLACE INTO temps_travail 
            (projet_id, annee, direction, categorie, mois, jours)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (projet_id, annee, direction, categorie, mois, jours))
    
    def _import_recettes_row(self, cursor, row, mapping, projet_id):
        """Importe une ligne de recettes"""
        annee = int(row[mapping['Année']]) if 'Année' in mapping else None
        mois = str(row[mapping['Mois']]) if 'Mois' in mapping else ""
        montant = float(row[mapping['Montant']]) if 'Montant' in mapping else 0.0
        detail = str(row[mapping['Détail']]) if 'Détail' in mapping else ""
        
        cursor.execute("""
            INSERT OR REPLACE INTO recettes 
            (projet_id, annee, mois, montant, detail)
            VALUES (?, ?, ?, ?, ?)
        """, (projet_id, annee, mois, montant, detail))
    
    def _import_depenses_row(self, cursor, row, mapping, projet_id):
        """Importe une ligne de dépenses"""
        annee = int(row[mapping['Année']]) if 'Année' in mapping else None
        mois = str(row[mapping['Mois']]) if 'Mois' in mapping else ""
        montant = float(row[mapping['Montant']]) if 'Montant' in mapping else 0.0
        detail = str(row[mapping['Détail']]) if 'Détail' in mapping else ""
        
        cursor.execute("""
            INSERT OR REPLACE INTO depenses 
            (projet_id, annee, mois, montant, detail)
            VALUES (?, ?, ?, ?, ?)
        """, (projet_id, annee, mois, montant, detail))
    
    def _import_autres_depenses_row(self, cursor, row, mapping, projet_id):
        """Importe une ligne d'autres dépenses"""
        annee = int(row[mapping['Année']]) if 'Année' in mapping else None
        ligne_index = int(row[mapping['Ligne']]) if 'Ligne' in mapping else 0
        mois = str(row[mapping['Mois']]) if 'Mois' in mapping else ""
        montant = float(row[mapping['Montant']]) if 'Montant' in mapping else 0.0
        detail = str(row[mapping['Détail']]) if 'Détail' in mapping else ""
        
        cursor.execute("""
            INSERT OR REPLACE INTO autres_depenses 
            (projet_id, annee, ligne_index, mois, montant, detail)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (projet_id, annee, ligne_index, mois, montant, detail))


class MatrixMappingWidget(QWidget):
    """Widget spécialisé pour l'import de temps de travail en format matriciel"""
    
    def __init__(self, data_type, parent_dialog, projet_id):
        super().__init__()
        self.data_type = data_type
        self.parent_dialog = parent_dialog
        self.projet_id = projet_id
        self.excel_sheets = {}
        self.team_members = []
        self.project_periods = []
        
        layout = QVBoxLayout()
        
        # === Configuration de base ===
        basic_config = QGroupBox("Configuration de base")
        basic_layout = QVBoxLayout()
        
        # Sélection feuille
        sheet_layout = QHBoxLayout()
        sheet_layout.addWidget(QLabel("Feuille Excel:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        sheet_layout.addWidget(self.sheet_combo)
        basic_layout.addLayout(sheet_layout)
        
        # Aperçu Excel
        excel_preview_group = QGroupBox("Aperçu Excel")
        excel_preview_layout = QVBoxLayout()
        
        self.excel_preview_table = QTableWidget()
        self.excel_preview_table.setMaximumHeight(200)
        excel_preview_layout.addWidget(self.excel_preview_table)
        
        excel_preview_group.setLayout(excel_preview_layout)
        basic_layout.addWidget(excel_preview_group)
        
        basic_config.setLayout(basic_layout)
        layout.addWidget(basic_config)
        
        # === Périodes du projet ===
        periods_group = QGroupBox("Périodes du projet")
        periods_layout = QVBoxLayout()
        
        self.periods_label = QLabel("Chargement des périodes...")
        periods_layout.addWidget(self.periods_label)
        
        # Liste des périodes détectées
        self.periods_list = QListWidget()
        self.periods_list.setMaximumHeight(100)
        periods_layout.addWidget(self.periods_list)
        
        periods_group.setLayout(periods_layout)
        layout.addWidget(periods_group)
        
        # === Mapping Équipe → Excel ===
        mapping_group = QGroupBox("Mapping Équipe → Excel")
        mapping_layout = QVBoxLayout()
        
        # Scroll area pour les mappings
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.mapping_content = QWidget()
        self.mapping_content_layout = QVBoxLayout()
        self.mapping_content.setLayout(self.mapping_content_layout)
        scroll.setWidget(self.mapping_content)
        scroll.setMinimumHeight(300)
        
        mapping_layout.addWidget(scroll)
        
        # Boutons d'action
        action_layout = QHBoxLayout()
        self.btn_detect_structure = QPushButton("🔍 Détecter structure auto")
        self.btn_reset_mapping = QPushButton("🔄 Reset mapping")
        self.btn_test_extraction = QPushButton("👁️ Test extraction")
        
        self.btn_detect_structure.clicked.connect(self.detect_structure)
        self.btn_reset_mapping.clicked.connect(self.reset_mapping)
        self.btn_test_extraction.clicked.connect(self.test_extraction)
        
        action_layout.addWidget(self.btn_detect_structure)
        action_layout.addWidget(self.btn_reset_mapping)
        action_layout.addWidget(self.btn_test_extraction)
        mapping_layout.addLayout(action_layout)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        self.setLayout(layout)
        
        # Variables pour stocker les widgets de mapping
        self.member_mapping_widgets = {}  # {member_id: {'position_combo': combo, 'period_mappings': {period: cellEdit}}}
        
        # Chargement initial
        self.load_team_members()
        self.load_project_periods()
    
    def load_team_members(self):
        """Charge les membres de l'équipe du projet"""
        try:
            conn = sqlite3.connect('gestion_budget.db')
            cursor = conn.cursor()
            cursor.execute("""
                SELECT type, nombre, direction 
                FROM equipe 
                WHERE projet_id = ? AND nombre > 0
            """, (self.projet_id,))
            
            self.team_members = []
            member_id = 0
            for type_person, nombre, direction in cursor.fetchall():
                for i in range(nombre):
                    member_id += 1
                    self.team_members.append({
                        'id': member_id,
                        'type': type_person,
                        'direction': direction,
                        'index': i + 1,
                        'display_name': f"{type_person} {i+1} ({direction})"
                    })
            
            conn.close()
            
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Impossible de charger l'équipe: {e}")
    
    def load_project_periods(self):
        """Génère les périodes basées sur les dates du projet"""
        try:
            conn = sqlite3.connect('gestion_budget.db')
            cursor = conn.cursor()
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id = ?", (self.projet_id,))
            result = cursor.fetchone()
            conn.close()
            
            if not result:
                self.periods_label.setText("❌ Projet non trouvé")
                return
            
            date_debut, date_fin = result
            
            # Parse des dates MM/yyyy
            try:
                debut_month, debut_year = map(int, date_debut.split('/'))
                fin_month, fin_year = map(int, date_fin.split('/'))
                
                # Génère toutes les périodes
                self.project_periods = []
                current_year = debut_year
                current_month = debut_month
                
                while (current_year < fin_year) or (current_year == fin_year and current_month <= fin_month):
                    # Ajoute l'année si pas déjà présente
                    year_period = f"{current_year}"
                    if year_period not in [p['period'] for p in self.project_periods]:
                        self.project_periods.append({
                            'period': year_period,
                            'type': 'année',
                            'display': f"Année {current_year}"
                        })
                    
                    # Ajoute le mois
                    month_names = ['', 'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
                                  'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
                    month_period = f"{current_year}-{current_month:02d}"
                    self.project_periods.append({
                        'period': month_period,
                        'type': 'mois',
                        'display': f"{month_names[current_month]} {current_year}"
                    })
                    
                    # Mois suivant
                    current_month += 1
                    if current_month > 12:
                        current_month = 1
                        current_year += 1
                
                # Met à jour l'affichage
                self.periods_label.setText(f"✅ {len(self.project_periods)} périodes générées")
                self.periods_list.clear()
                for period in self.project_periods:
                    self.periods_list.addItem(f"{period['display']} ({period['type']})")
                
                # Crée les widgets de mapping
                self.create_mapping_widgets()
                
            except ValueError:
                self.periods_label.setText("❌ Format de date invalide (attendu: MM/yyyy)")
                
        except Exception as e:
            self.periods_label.setText(f"❌ Erreur: {e}")
    
    def create_mapping_widgets(self):
        """Crée les widgets de mapping pour chaque membre d'équipe"""
        # Nettoie les anciens widgets
        while self.mapping_content_layout.count():
            child = self.mapping_content_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        self.member_mapping_widgets = {}
        
        for member in self.team_members:
            member_id = member['id']
            
            # Groupe pour ce membre
            member_group = QGroupBox(f"👤 {member['display_name']}")
            member_layout = QVBoxLayout()
            
            # Position de base (ligne/colonne/cellule)
            position_layout = QHBoxLayout()
            position_layout.addWidget(QLabel("Position:"))
            
            # Type de position
            position_type_combo = QComboBox()
            position_type_combo.addItems(["Ligne", "Colonne", "Cellule fixe"])
            position_layout.addWidget(position_type_combo)
            
            # Valeur de position
            position_value_edit = QLineEdit()
            position_value_edit.setPlaceholderText("Ex: 5, C, A5...")
            position_layout.addWidget(position_value_edit)
            
            member_layout.addLayout(position_layout)
            
            # Mapping des périodes
            periods_scroll = QScrollArea()
            periods_scroll.setMaximumHeight(150)
            periods_widget = QWidget()
            periods_layout = QGridLayout()
            periods_widget.setLayout(periods_layout)
            periods_scroll.setWidget(periods_widget)
            periods_scroll.setWidgetResizable(True)
            
            # Headers
            periods_layout.addWidget(QLabel("Période"), 0, 0)
            periods_layout.addWidget(QLabel("Cellule Excel"), 0, 1)
            periods_layout.addWidget(QLabel("Auto"), 0, 2)
            
            # Widgets pour chaque période
            period_mappings = {}
            for row, period in enumerate(self.project_periods, 1):
                # Nom de la période
                periods_layout.addWidget(QLabel(period['display']), row, 0)
                
                # Cellule Excel
                cell_edit = QLineEdit()
                cell_edit.setPlaceholderText("Ex: B5, C10...")
                periods_layout.addWidget(cell_edit, row, 1)
                
                # Bouton auto-calcul
                auto_btn = QPushButton("🎯")
                auto_btn.setMaximumWidth(30)
                auto_btn.setToolTip("Calcul automatique basé sur la position")
                auto_btn.clicked.connect(lambda checked, m=member_id, p=period['period']: self.auto_calculate_cell(m, p))
                periods_layout.addWidget(auto_btn, row, 2)
                
                period_mappings[period['period']] = cell_edit
            
            member_layout.addWidget(periods_scroll)
            
            # Stocke les widgets
            self.member_mapping_widgets[member_id] = {
                'position_type_combo': position_type_combo,
                'position_value_edit': position_value_edit,
                'period_mappings': period_mappings,
                'group_widget': member_group
            }
            
            # Connexions pour auto-calcul
            position_type_combo.currentTextChanged.connect(lambda: self.update_auto_calculations(member_id))
            position_value_edit.textChanged.connect(lambda: self.update_auto_calculations(member_id))
            
            member_group.setLayout(member_layout)
            self.mapping_content_layout.addWidget(member_group)
    
    def update_excel_data(self, excel_sheets):
        """Met à jour les données Excel disponibles"""
        self.excel_sheets = excel_sheets
        
        # Met à jour la liste des feuilles
        self.sheet_combo.clear()
        self.sheet_combo.addItems(list(excel_sheets.keys()))
        
        if excel_sheets:
            self.on_sheet_changed()
    
    def on_sheet_changed(self):
        """Appelé quand la feuille Excel change"""
        self.update_excel_preview()
    
    def update_excel_preview(self):
        """Met à jour l'aperçu Excel"""
        sheet_name = self.sheet_combo.currentText()
        if not sheet_name or not self.parent_dialog.excel_file:
            return
        
        try:
            # Lit un aperçu du fichier
            df = pd.read_excel(
                self.parent_dialog.excel_file,
                sheet_name=sheet_name,
                header=None,  # Pas d'en-tête auto
                nrows=20  # Limite pour aperçu
            )
            
            # Met à jour le tableau d'aperçu
            self.excel_preview_table.setRowCount(min(20, len(df)))
            self.excel_preview_table.setColumnCount(min(20, len(df.columns)))
            
            # Headers avec coordonnées
            col_headers = [f"{chr(65 + i)}" for i in range(min(20, len(df.columns)))]
            row_headers = [str(i + 1) for i in range(min(20, len(df)))]
            
            self.excel_preview_table.setHorizontalHeaderLabels(col_headers)
            self.excel_preview_table.setVerticalHeaderLabels(row_headers)
            
            # Remplit les données
            for row in range(min(20, len(df))):
                for col in range(min(20, len(df.columns))):
                    value = str(df.iloc[row, col])
                    if value == 'nan':
                        value = ''
                    
                    item = QTableWidgetItem(value)
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    
                    # Coloration pour faciliter la lecture
                    if row % 2 == 0:
                        item.setBackground(Qt.GlobalColor.lightGray)
                    
                    self.excel_preview_table.setItem(row, col, item)
            
            # Ajuste les tailles
            self.excel_preview_table.resizeColumnsToContents()
            
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Erreur lecture Excel: {e}")
    
    def auto_calculate_cell(self, member_id, period):
        """Calcule automatiquement la cellule pour un membre et une période"""
        if member_id not in self.member_mapping_widgets:
            return
        
        widgets = self.member_mapping_widgets[member_id]
        position_type = widgets['position_type_combo'].currentText()
        position_value = widgets['position_value_edit'].text().strip()
        
        if not position_value:
            QMessageBox.warning(self, "Erreur", "Définissez d'abord la position de base")
            return
        
        try:
            if position_type == "Ligne":
                # Position = numéro de ligne, les colonnes varient par période
                row_num = int(position_value)
                # Trouve l'index de la période
                period_index = next((i for i, p in enumerate(self.project_periods) if p['period'] == period), 0)
                col_letter = chr(65 + period_index)  # A, B, C...
                calculated_cell = f"{col_letter}{row_num}"
                
            elif position_type == "Colonne":
                # Position = lettre de colonne, les lignes varient par période
                col_letter = position_value.upper()
                period_index = next((i for i, p in enumerate(self.project_periods) if p['period'] == period), 0)
                row_num = period_index + 1  # 1, 2, 3...
                calculated_cell = f"{col_letter}{row_num}"
                
            else:  # Cellule fixe
                calculated_cell = position_value
            
            # Met à jour le champ
            if period in widgets['period_mappings']:
                widgets['period_mappings'][period].setText(calculated_cell)
                
        except Exception as e:
            QMessageBox.warning(self, "Erreur", f"Calcul impossible: {e}")
    
    def update_auto_calculations(self, member_id):
        """Met à jour tous les calculs automatiques pour un membre"""
        # Déclenche le recalcul pour toutes les périodes
        for period in [p['period'] for p in self.project_periods]:
            self.auto_calculate_cell(member_id, period)
    
    def detect_structure(self):
        """Détection automatique de la structure Excel"""
        QMessageBox.information(self, "Détection auto", 
                              "🚧 Fonctionnalité en développement\n\n"
                              "Pour l'instant, configurez manuellement:\n"
                              "1. Définissez la position de chaque membre\n"
                              "2. Utilisez les boutons 🎯 pour le calcul auto")
    
    def reset_mapping(self):
        """Remet à zéro tous les mappings"""
        reply = QMessageBox.question(self, "Reset", 
                                   "Effacer tous les mappings ?",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            for member_id, widgets in self.member_mapping_widgets.items():
                widgets['position_value_edit'].clear()
                for cell_edit in widgets['period_mappings'].values():
                    cell_edit.clear()
    
    def test_extraction(self):
        """Test l'extraction des données avec la configuration actuelle"""
        if not self.has_valid_mapping():
            QMessageBox.warning(self, "Erreur", "Configurez d'abord le mapping")
            return
        
        # Ouvre une dialog de prévisualisation
        dialog = MatrixPreviewDialog(self, self.parent_dialog.excel_file, self.sheet_combo.currentText())
        dialog.exec()
    
    def reset_configuration(self):
        """Remet à zéro la configuration"""
        self.sheet_combo.clear()
        self.member_mapping_widgets = {}
    
    def get_configuration(self):
        """Retourne la configuration actuelle"""
        config = {
            'sheet': self.sheet_combo.currentText(),
            'team_mappings': {}
        }
        
        for member_id, widgets in self.member_mapping_widgets.items():
            member = next((m for m in self.team_members if m['id'] == member_id), None)
            if member:
                config['team_mappings'][member_id] = {
                    'member_info': member,
                    'position_type': widgets['position_type_combo'].currentText(),
                    'position_value': widgets['position_value_edit'].text(),
                    'period_mappings': {period: edit.text() for period, edit in widgets['period_mappings'].items()}
                }
        
        return config
    
    def load_configuration(self, config):
        """Charge une configuration"""
        if 'sheet' in config:
            self.sheet_combo.setCurrentText(config['sheet'])
        
        # TODO: Charger les mappings
    
    def has_valid_mapping(self):
        """Vérifie si le mapping est valide"""
        if not self.sheet_combo.currentText():
            return False
        
        # Vérifie qu'au moins un membre a un mapping complet
        for member_id, widgets in self.member_mapping_widgets.items():
            if widgets['position_value_edit'].text().strip():
                # Vérifie qu'au moins une période est mappée
                for cell_edit in widgets['period_mappings'].values():
                    if cell_edit.text().strip():
                        return True
        
        return False
    
    def import_data_to_database(self, excel_file, projet_id):
        """Importe les données dans la base de données"""
        if not self.has_valid_mapping():
            return False
        
        try:
            # Lit le fichier Excel complet
            df = pd.read_excel(
                excel_file,
                sheet_name=self.get_configuration()['sheet'],
                header=None
            )
            
            conn = sqlite3.connect('gestion_budget.db')
            cursor = conn.cursor()
            
            imported_count = 0
            
            # Pour chaque membre de l'équipe
            for member_id, widgets in self.member_mapping_widgets.items():
                member = next((m for m in self.team_members if m['id'] == member_id), None)
                if not member:
                    continue
                
                # Pour chaque période mappée
                for period, cell_edit in widgets['period_mappings'].items():
                    cell_ref = cell_edit.text().strip()
                    if not cell_ref:
                        continue
                    
                    try:
                        # Parse la référence de cellule (ex: B5)
                        col_letter = ''.join(filter(str.isalpha, cell_ref))
                        row_num = int(''.join(filter(str.isdigit, cell_ref)))
                        
                        # Convertit la lettre en index (A=0, B=1, etc.)
                        col_index = ord(col_letter.upper()) - ord('A')
                        row_index = row_num - 1  # Excel commence à 1, pandas à 0
                        
                        # Extrait la valeur
                        if row_index < len(df) and col_index < len(df.columns):
                            value = df.iloc[row_index, col_index]
                            
                            # Convertit en float si possible
                            try:
                                jours = float(value) if pd.notna(value) else 0.0
                            except:
                                jours = 0.0
                            
                            if jours > 0:
                                # Détermine l'année et le mois
                                if '-' in period:  # Format YYYY-MM
                                    annee, mois = period.split('-')
                                    annee = int(annee)
                                    mois = int(mois)
                                else:  # Format YYYY (année seule)
                                    annee = int(period)
                                    mois = 1  # Par défaut janvier
                                
                                # Insert dans la base
                                cursor.execute("""
                                    INSERT OR REPLACE INTO temps_travail 
                                    (projet_id, annee, direction, categorie, mois, jours)
                                    VALUES (?, ?, ?, ?, ?, ?)
                                """, (projet_id, annee, member['direction'], member['type'], str(mois), jours))
                                
                                imported_count += 1
                    
                    except Exception as e:
                        print(f"Erreur extraction {member['display_name']} - {period}: {e}")
                        continue
            
            conn.commit()
            conn.close()
            
            QMessageBox.information(self, "Import réussi", 
                                  f"✅ {imported_count} données importées avec succès !")
            return True
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'import", f"❌ Erreur: {e}")
            return False


class MatrixPreviewDialog(QDialog):
    """Dialog de prévisualisation des données matricielles"""
    
    def __init__(self, parent, excel_file, sheet_name):
        super().__init__(parent)
        self.parent_widget = parent
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        
        self.setWindowTitle("👁️ Aperçu extraction matricielle")
        self.setMinimumSize(1000, 600)
        
        layout = QVBoxLayout()
        
        # Info
        info_label = QLabel(f"📊 Aperçu extraction - Feuille: {sheet_name}")
        info_label.setStyleSheet("font-weight: bold; padding: 5px;")
        layout.addWidget(info_label)
        
        # Table de prévisualisation
        self.preview_table = QTableWidget()
        layout.addWidget(self.preview_table)
        
        # Boutons
        btn_layout = QHBoxLayout()
        btn_close = QPushButton("Fermer")
        btn_close.clicked.connect(self.close)
        btn_layout.addWidget(btn_close)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
        
        # Charge l'aperçu
        self.load_preview()
    
    def load_preview(self):
        """Charge l'aperçu des données qui seraient extraites"""
        try:
            # Récupère la configuration du parent
            config = self.parent_widget.get_configuration()
            
            # Lit le fichier Excel
            df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name, header=None)
            
            # Prépare les données d'aperçu
            preview_data = []
            
            for member_id, mapping in config['team_mappings'].items():
                member_info = mapping['member_info']
                
                for period, cell_ref in mapping['period_mappings'].items():
                    if not cell_ref.strip():
                        continue
                    
                    try:
                        # Parse la cellule
                        col_letter = ''.join(filter(str.isalpha, cell_ref))
                        row_num = int(''.join(filter(str.isdigit, cell_ref)))
                        col_index = ord(col_letter.upper()) - ord('A')
                        row_index = row_num - 1
                        
                        # Extrait la valeur
                        value = "N/A"
                        if row_index < len(df) and col_index < len(df.columns):
                            raw_value = df.iloc[row_index, col_index]
                            value = str(raw_value) if pd.notna(raw_value) else "0"
                        
                        preview_data.append([
                            member_info['display_name'],
                            period,
                            cell_ref,
                            value
                        ])
                    
                    except Exception as e:
                        preview_data.append([
                            member_info['display_name'],
                            period,
                            cell_ref,
                            f"Erreur: {e}"
                        ])
            
            # Remplit le tableau
            self.preview_table.setRowCount(len(preview_data))
            self.preview_table.setColumnCount(4)
            self.preview_table.setHorizontalHeaderLabels([
                "Membre équipe", "Période", "Cellule Excel", "Valeur extraite"
            ])
            
            for row, data in enumerate(preview_data):
                for col, value in enumerate(data):
                    item = QTableWidgetItem(str(value))
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    self.preview_table.setItem(row, col, item)
            
            self.preview_table.resizeColumnsToContents()
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de générer l'aperçu: {e}")


class ManualMatchDialog(QDialog):
    """Dialog pour le matching manuel entre équipe et personnes Excel"""
    
    def __init__(self, parent, team_members, found_persons):
        super().__init__(parent)
        self.team_members = team_members
        self.found_persons = found_persons
        self.mappings = {}
        
        self.setWindowTitle("Matching manuel des personnes")
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout()
        
        # Instructions
        instructions = QLabel(
            "Associez chaque personne de votre équipe avec une ligne de l'Excel.\n"
            "Une personne de l'équipe peut rester non assignée si elle n'apparaît pas dans l'Excel."
        )
        layout.addWidget(instructions)
        
        # Table de matching
        self.match_table = QTableWidget()
        self.match_table.setColumnCount(3)
        self.match_table.setHorizontalHeaderLabels([
            "Membre de l'équipe", "Personne Excel", "Statut"
        ])
        self.match_table.setRowCount(len(team_members))
        
        # Remplit la table
        for i, member in enumerate(team_members):
            # Colonne 1: Membre équipe
            item = QTableWidgetItem(member['display_name'])
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.match_table.setItem(i, 0, item)
            
            # Colonne 2: Combo pour sélectionner la personne Excel
            combo = QComboBox()
            combo.addItem("-- Non assigné --")
            for person in found_persons:
                combo.addItem(f"Ligne {person['row']}: {person['nom']}")
            self.match_table.setCellWidget(i, 1, combo)
            
            # Colonne 3: Statut
            status_item = QTableWidgetItem("Non assigné")
            status_item.setFlags(status_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.match_table.setItem(i, 2, status_item)
            
            # Connexion pour mise à jour du statut
            combo.currentTextChanged.connect(lambda text, row=i: self.update_status(row, text))
        
        layout.addWidget(self.match_table)
        
        # Boutons
        btn_layout = QHBoxLayout()
        btn_auto = QPushButton("Auto-matching par nom")
        btn_ok = QPushButton("Valider")
        btn_cancel = QPushButton("Annuler")
        
        btn_auto.clicked.connect(self.auto_match_by_name)
        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addWidget(btn_auto)
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
    
    def update_status(self, row, text):
        """Met à jour le statut d'une ligne"""
        if text == "-- Non assigné --":
            status = "Non assigné"
        else:
            status = "Assigné"
        
        self.match_table.setItem(row, 2, QTableWidgetItem(status))
    
    def auto_match_by_name(self):
        """Tentative de matching automatique par similarité de nom"""
        for i, member in enumerate(self.team_members):
            combo = self.match_table.cellWidget(i, 1)
            member_type = member['type'].lower()
            
            # Cherche une correspondance partielle dans les noms Excel
            best_match = None
            for j, person in enumerate(self.found_persons):
                person_nom = person['nom'].lower()
                # Logique simple: cherche le type dans le nom
                if member_type in person_nom:
                    best_match = j + 1  # +1 car index 0 = "Non assigné"
                    break
            
            if best_match:
                combo.setCurrentIndex(best_match)
    
    def get_mappings(self):
        """Retourne les mappings configurés"""
        mappings = {}
        for i, member in enumerate(self.team_members):
            combo = self.match_table.cellWidget(i, 1)
            if combo.currentIndex() > 0:  # Pas "Non assigné"
                person_idx = combo.currentIndex() - 1
                person = self.found_persons[person_idx]
                mappings[person['excel_row']] = {
                    'team_member': member,
                    'excel_person': person
                }
        return mappings


class DataPreviewDialog(QDialog):
    """Dialog d'aperçu des données extraites"""
    
    def __init__(self, parent, excel_file, sheet_name):
        super().__init__(parent)
        self.setWindowTitle("Aperçu des données extraites")
        self.setMinimumSize(1000, 600)
        
        layout = QVBoxLayout()
        
        # TODO: Implémenter l'aperçu des données matricielles
        info = QLabel("Aperçu des données matricielles - En cours de développement")
        layout.addWidget(info)
        
        btn_close = QPushButton("Fermer")
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)
        
        self.setLayout(layout)


class PreviewDialog(QDialog):
    def __init__(self, parent, excel_file, mapping_widgets, projet_id):
        super().__init__(parent)
        self.setWindowTitle("Aperçu des données à importer")
        self.setMinimumSize(1000, 600)
        
        layout = QVBoxLayout()
        
        # Onglets pour chaque type de données
        tabs = QTabWidget()
        
        for data_type, widget in mapping_widgets.items():
            if widget.has_valid_mapping():
                preview_widget = self._create_preview_widget(excel_file, widget, projet_id)
                tabs.addTab(preview_widget, data_type.replace('_', ' ').title())
        
        layout.addWidget(tabs)
        
        # Bouton fermer
        btn_close = QPushButton("Fermer")
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)
        
        self.setLayout(layout)
    
    def _create_preview_widget(self, excel_file, mapping_widget, projet_id):
        """Crée un widget d'aperçu pour un type de données"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        try:
            config = mapping_widget.get_configuration()
            
            # Calcule les lignes à ignorer entre l'en-tête et le début des données
            header_row = config['header_row']
            data_start_row = config['data_start_row']
            
            if data_start_row > header_row + 1:
                skip_rows = list(range(header_row + 1, data_start_row))
            else:
                skip_rows = None
            
            # Lit les données
            df = pd.read_excel(
                excel_file,
                sheet_name=config['sheet'],
                header=header_row,
                skiprows=skip_rows,
                nrows=20  # Limite pour l'aperçu
            )
            
            # Applique le filtre si spécifié
            if config['filter'].strip():
                df = df.query(config['filter'])
            
            # Affiche les informations
            info_label = QLabel(f"Feuille: {config['sheet']} | Lignes trouvées: {len(df)}")
            layout.addWidget(info_label)
            
            # Tableau d'aperçu
            table = QTableWidget()
            table.setRowCount(min(10, len(df)))
            
            # Colonnes mappées seulement
            mapped_columns = []
            for field, mapping in config['mappings'].items():
                if mapping['excel_column'] != "-- Non mappé --":
                    mapped_columns.append((field, mapping['excel_column']))
            
            table.setColumnCount(len(mapped_columns))
            table.setHorizontalHeaderLabels([f"{field}\n({col})" for field, col in mapped_columns])
            
            # Remplit le tableau
            for row in range(min(10, len(df))):
                for col, (field, excel_col) in enumerate(mapped_columns):
                    if excel_col in df.columns:
                        value = str(df.iloc[row][excel_col])
                        item = QTableWidgetItem(value)
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                        table.setItem(row, col, item)
            
            layout.addWidget(table)
            
        except Exception as e:
            error_label = QLabel(f"Erreur: {e}")
            layout.addWidget(error_label)
        
        widget.setLayout(layout)
        return widget
