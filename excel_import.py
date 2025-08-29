import pandas as pd
import json
import sqlite3
import os
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLabel, QComboBox, QTableWidget, QTableWidgetItem,
                             QFileDialog, QMessageBox, QLineEdit, QSpinBox,
                             QTextEdit, QTabWidget, QWidget, QFormLayout,
                             QCheckBox, QGroupBox, QScrollArea, QFrame)
from PyQt6.QtCore import Qt
import traceback

class ExcelImportConfigManager:
    """Gestionnaire des configurations d'import Excel"""
    
    def __init__(self):
        self.config_dir = "import_configs"
        self.ensure_config_dir()
    
    def ensure_config_dir(self):
        """Crée le répertoire de configuration s'il n'existe pas"""
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
    
    def save_config(self, name: str, config: Dict) -> bool:
        """Sauvegarde une configuration"""
        try:
            filepath = os.path.join(self.config_dir, f"{name}.json")
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Erreur sauvegarde config: {e}")
            return False
    
    def load_config(self, name: str) -> Optional[Dict]:
        """Charge une configuration"""
        try:
            filepath = os.path.join(self.config_dir, f"{name}.json")
            if os.path.exists(filepath):
                with open(filepath, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Erreur chargement config: {e}")
        return None
    
    def list_configs(self) -> List[str]:
        """Liste toutes les configurations disponibles"""
        configs = []
        if os.path.exists(self.config_dir):
            for file in os.listdir(self.config_dir):
                if file.endswith('.json'):
                    configs.append(file[:-5])  # Enlever .json
        return configs
    
    def delete_config(self, name: str) -> bool:
        """Supprime une configuration"""
        try:
            filepath = os.path.join(self.config_dir, f"{name}.json")
            if os.path.exists(filepath):
                os.remove(filepath)
                return True
        except Exception as e:
            print(f"Erreur suppression config: {e}")
        return False

class TempsTravailvailMapper:
    """Classe pour mapper les données Excel vers la table temps_travail"""
    
    def __init__(self):
        self.db_path = "gestion_budget.db"
        
    def get_available_projects(self) -> List[Tuple[int, str, str]]:
        """Récupère la liste des projets disponibles"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, code, nom FROM projets ORDER BY code")
        projects = cursor.fetchall()
        conn.close()
        return projects
    
    def get_project_team_data(self, projet_id: int) -> List[Tuple[str, str]]:
        """Récupère les données d'équipe pour un projet (direction, catégorie)"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT direction, type FROM equipe WHERE projet_id = ?", (projet_id,))
        team_data = cursor.fetchall()
        conn.close()
        return team_data
    
    def get_project_dates(self, projet_id: int) -> Optional[Tuple[str, str]]:
        """Récupère les dates de début et fin d'un projet"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id = ?", (projet_id,))
        dates = cursor.fetchone()
        conn.close()
        return dates
    
    def generate_months_list(self, date_debut: str, date_fin: str) -> List[str]:
        """Génère la liste des mois entre deux dates au format MM/yyyy"""
        try:
            debut = datetime.strptime(date_debut, '%m/%Y')
            fin = datetime.strptime(date_fin, '%m/%Y')
            
            months = []
            current = debut
            while current <= fin:
                months.append(current.strftime('%m/%Y'))
                # Passer au mois suivant
                if current.month == 12:
                    current = current.replace(year=current.year + 1, month=1)
                else:
                    current = current.replace(month=current.month + 1)
            
            return months
        except Exception as e:
            print(f"Erreur génération mois: {e}")
            return []
    
    def process_excel_data(self, df: pd.DataFrame, config: Dict, projet_id: int) -> List[Dict]:
        """Traite les données Excel selon la configuration"""
        results = []
        
        try:
            # Récupérer les données du projet
            project_dates = self.get_project_dates(projet_id)
            if not project_dates or not project_dates[0] or not project_dates[1]:
                raise ValueError("Dates du projet non définies")
            
            # Générer la liste des mois du projet
            months_list = self.generate_months_list(project_dates[0], project_dates[1])
            if not months_list:
                raise ValueError("Impossible de générer la liste des mois")
            
            # Récupérer les données d'équipe
            team_data = self.get_project_team_data(projet_id)
            
            # Configuration des colonnes
            column_mapping = config.get('column_mapping', {})
            data_structure = config.get('data_structure', 'rows')  # 'rows' ou 'columns'
            
            if data_structure == 'rows':
                # Chaque ligne = une entrée temps_travail
                results = self._process_row_based_data(df, config, projet_id, months_list, team_data)
            else:
                # Structure en colonnes (mois en colonnes par exemple)
                results = self._process_column_based_data(df, config, projet_id, months_list, team_data)
                
        except Exception as e:
            print(f"Erreur traitement données: {e}")
            traceback.print_exc()
        
        return results
    
    def _process_row_based_data(self, df: pd.DataFrame, config: Dict, projet_id: int, 
                               months_list: List[str], team_data: List[Tuple[str, str]]) -> List[Dict]:
        """Traite les données organisées en lignes"""
        results = []
        column_mapping = config.get('column_mapping', {})
        
        for index, row in df.iterrows():
            try:
                # Extraction des valeurs selon le mapping
                direction = self._extract_value(row, column_mapping.get('direction'))
                categorie = self._extract_value(row, column_mapping.get('categorie'))
                annee = self._extract_value(row, column_mapping.get('annee'))
                mois = self._extract_value(row, column_mapping.get('mois'))
                jours = self._extract_value(row, column_mapping.get('jours'))
                
                # Validation et nettoyage
                if not all([direction, categorie, annee, mois]) or jours is None:
                    continue
                
                # Normalisation du mois
                mois_normalise = self._normalize_month(mois, annee)
                if mois_normalise not in months_list:
                    continue  # Mois hors période du projet
                
                results.append({
                    'projet_id': projet_id,
                    'annee': int(annee),
                    'direction': str(direction),
                    'categorie': str(categorie),
                    'mois': mois_normalise,
                    'jours': float(jours)
                })
                
            except Exception as e:
                print(f"Erreur ligne {index}: {e}")
                continue
        
        return results
    
    def _process_column_based_data(self, df: pd.DataFrame, config: Dict, projet_id: int,
                                  months_list: List[str], team_data: List[Tuple[str, str]]) -> List[Dict]:
        """Traite les données organisées en colonnes (ex: mois en colonnes)"""
        results = []
        column_mapping = config.get('column_mapping', {})
        month_columns = config.get('month_columns', [])
        
        for index, row in df.iterrows():
            try:
                # Extraction des données fixes
                direction = self._extract_value(row, column_mapping.get('direction'))
                categorie = self._extract_value(row, column_mapping.get('categorie'))
                annee = self._extract_value(row, column_mapping.get('annee'))
                
                if not all([direction, categorie, annee]):
                    continue
                
                # Traitement des colonnes de mois
                for month_col_config in month_columns:
                    col_name = month_col_config.get('column')
                    mois_value = month_col_config.get('month')
                    
                    if col_name in df.columns:
                        jours = self._extract_value(row, {'column': col_name})
                        if jours is not None and jours > 0:
                            mois_normalise = self._normalize_month(mois_value, annee)
                            if mois_normalise in months_list:
                                results.append({
                                    'projet_id': projet_id,
                                    'annee': int(annee),
                                    'direction': str(direction),
                                    'categorie': str(categorie),
                                    'mois': mois_normalise,
                                    'jours': float(jours)
                                })
                
            except Exception as e:
                print(f"Erreur ligne {index}: {e}")
                continue
        
        return results
    
    def _extract_value(self, row, column_config):
        """Extrait une valeur selon la configuration de colonne"""
        if not column_config:
            return None
        
        col_name = column_config.get('column')
        default_value = column_config.get('default')
        transform = column_config.get('transform')
        
        if col_name and col_name in row.index:
            value = row[col_name]
            
            # Gestion des valeurs vides
            if pd.isna(value) or value == '':
                return default_value
            
            # Transformations
            if transform == 'upper':
                value = str(value).upper()
            elif transform == 'lower':
                value = str(value).lower()
            elif transform == 'strip':
                value = str(value).strip()
            
            return value
        
        return default_value
    
    def _normalize_month(self, mois, annee):
        """Normalise le format du mois vers MM/yyyy"""
        try:
            if isinstance(mois, str) and '/' in mois:
                return mois  # Déjà au bon format
            
            # Si c'est un numéro de mois
            if isinstance(mois, (int, float)):
                return f"{int(mois):02d}/{int(annee)}"
            
            # Si c'est un nom de mois
            month_names = {
                'janvier': 1, 'février': 2, 'mars': 3, 'avril': 4,
                'mai': 5, 'juin': 6, 'juillet': 7, 'août': 8,
                'septembre': 9, 'octobre': 10, 'novembre': 11, 'décembre': 12
            }
            
            mois_lower = str(mois).lower()
            if mois_lower in month_names:
                return f"{month_names[mois_lower]:02d}/{int(annee)}"
            
        except Exception as e:
            print(f"Erreur normalisation mois: {e}")
        
        return None
    
    def save_to_database(self, data: List[Dict]) -> Tuple[bool, str]:
        """Sauvegarde les données dans la base"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Vérifier que la table existe
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
            
            # Insertion des données
            inserted_count = 0
            for entry in data:
                try:
                    cursor.execute("""
                        INSERT OR REPLACE INTO temps_travail 
                        (projet_id, annee, direction, categorie, mois, jours)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (
                        entry['projet_id'],
                        entry['annee'],
                        entry['direction'],
                        entry['categorie'],
                        entry['mois'],
                        entry['jours']
                    ))
                    inserted_count += 1
                except Exception as e:
                    print(f"Erreur insertion: {e}")
                    continue
            
            conn.commit()
            conn.close()
            
            return True, f"Import réussi: {inserted_count} entrées sauvegardées"
            
        except Exception as e:
            return False, f"Erreur sauvegarde: {str(e)}"

class ExcelImportDialog(QDialog):
    """Interface graphique pour l'import Excel configurable"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Import Excel - Temps de Travail")
        self.setMinimumSize(1200, 800)
        
        self.config_manager = ExcelImportConfigManager()
        self.mapper = TempsTravailvailMapper()
        self.current_df = None
        self.current_config = None
        
        self.setup_ui()
        self.load_projects()
        self.load_configs()
    
    def setup_ui(self):
        """Configure l'interface utilisateur"""
        layout = QVBoxLayout()
        
        # Onglets
        self.tab_widget = QTabWidget()
        
        # Onglet 1: Configuration
        self.setup_config_tab()
        
        # Onglet 2: Import
        self.setup_import_tab()
        
        # Onglet 3: Preview
        self.setup_preview_tab()
        
        layout.addWidget(self.tab_widget)
        
        # Boutons
        button_layout = QHBoxLayout()
        
        self.test_btn = QPushButton("Test Configuration")
        self.test_btn.clicked.connect(self.test_configuration)
        button_layout.addWidget(self.test_btn)
        
        self.import_btn = QPushButton("Importer")
        self.import_btn.clicked.connect(self.perform_import)
        self.import_btn.setEnabled(False)
        button_layout.addWidget(self.import_btn)
        
        button_layout.addStretch()
        
        close_btn = QPushButton("Fermer")
        close_btn.clicked.connect(self.close)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def setup_config_tab(self):
        """Configure l'onglet de configuration"""
        config_tab = QWidget()
        layout = QVBoxLayout()
        
        # Gestion des configurations
        config_group = QGroupBox("Gestion des Configurations")
        config_layout = QHBoxLayout()
        
        config_layout.addWidget(QLabel("Configuration:"))
        self.config_combo = QComboBox()
        self.config_combo.currentTextChanged.connect(self.load_selected_config)
        config_layout.addWidget(self.config_combo)
        
        new_btn = QPushButton("Nouvelle")
        new_btn.clicked.connect(self.new_config)
        config_layout.addWidget(new_btn)
        
        save_btn = QPushButton("Sauvegarder")
        save_btn.clicked.connect(self.save_config)
        config_layout.addWidget(save_btn)
        
        delete_btn = QPushButton("Supprimer")
        delete_btn.clicked.connect(self.delete_config)
        config_layout.addWidget(delete_btn)
        
        config_group.setLayout(config_layout)
        layout.addWidget(config_group)
        
        # Configuration détaillée
        detail_group = QGroupBox("Configuration Détaillée")
        detail_layout = QFormLayout()
        
        self.config_name_edit = QLineEdit()
        detail_layout.addRow("Nom:", self.config_name_edit)
        
        self.description_edit = QTextEdit()
        self.description_edit.setMaximumHeight(60)
        detail_layout.addRow("Description:", self.description_edit)
        
        self.structure_combo = QComboBox()
        self.structure_combo.addItems(["rows", "columns"])
        self.structure_combo.currentTextChanged.connect(self.on_structure_changed)
        detail_layout.addRow("Structure données:", self.structure_combo)
        
        # Mapping des colonnes
        self.setup_column_mapping_ui(detail_layout)
        
        detail_group.setLayout(detail_layout)
        layout.addWidget(detail_group)
        
        config_tab.setLayout(layout)
        self.tab_widget.addTab(config_tab, "Configuration")
    
    def setup_column_mapping_ui(self, parent_layout):
        """Configure l'interface de mapping des colonnes"""
        # Groupe pour le mapping
        mapping_group = QGroupBox("Mapping des Colonnes")
        mapping_layout = QFormLayout()
        
        # Colonnes obligatoires
        self.direction_col_edit = QLineEdit()
        mapping_layout.addRow("Colonne Direction:", self.direction_col_edit)
        
        self.categorie_col_edit = QLineEdit()
        mapping_layout.addRow("Colonne Catégorie:", self.categorie_col_edit)
        
        self.annee_col_edit = QLineEdit()
        mapping_layout.addRow("Colonne Année:", self.annee_col_edit)
        
        self.mois_col_edit = QLineEdit()
        mapping_layout.addRow("Colonne Mois:", self.mois_col_edit)
        
        self.jours_col_edit = QLineEdit()
        mapping_layout.addRow("Colonne Jours:", self.jours_col_edit)
        
        # Zone pour colonnes de mois (structure en colonnes)
        self.month_columns_group = QGroupBox("Colonnes de Mois (pour structure en colonnes)")
        self.month_columns_layout = QVBoxLayout()
        self.month_columns_group.setLayout(self.month_columns_layout)
        
        add_month_btn = QPushButton("Ajouter Colonne Mois")
        add_month_btn.clicked.connect(self.add_month_column)
        self.month_columns_layout.addWidget(add_month_btn)
        
        mapping_group.setLayout(mapping_layout)
        parent_layout.addRow(mapping_group)
        parent_layout.addRow(self.month_columns_group)
    
    def setup_import_tab(self):
        """Configure l'onglet d'import"""
        import_tab = QWidget()
        layout = QVBoxLayout()
        
        # Sélection du projet
        project_layout = QHBoxLayout()
        project_layout.addWidget(QLabel("Projet:"))
        self.project_combo = QComboBox()
        project_layout.addWidget(self.project_combo)
        project_layout.addStretch()
        layout.addLayout(project_layout)
        
        # Sélection du fichier Excel
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("Fichier Excel:"))
        self.file_path_edit = QLineEdit()
        file_layout.addWidget(self.file_path_edit)
        
        browse_btn = QPushButton("Parcourir")
        browse_btn.clicked.connect(self.browse_excel_file)
        file_layout.addWidget(browse_btn)
        
        layout.addLayout(file_layout)
        
        # Options d'import
        options_group = QGroupBox("Options d'Import")
        options_layout = QFormLayout()
        
        self.sheet_combo = QComboBox()
        options_layout.addRow("Feuille:", self.sheet_combo)
        
        self.header_row_spin = QSpinBox()
        self.header_row_spin.setRange(0, 100)
        self.header_row_spin.setValue(0)
        options_layout.addRow("Ligne d'en-tête:", self.header_row_spin)
        
        self.skip_rows_spin = QSpinBox()
        self.skip_rows_spin.setRange(0, 100)
        options_layout.addRow("Lignes à ignorer:", self.skip_rows_spin)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # Aperçu du fichier Excel
        preview_group = QGroupBox("Aperçu du Fichier")
        preview_layout = QVBoxLayout()
        
        self.excel_preview_table = QTableWidget()
        preview_layout.addWidget(self.excel_preview_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        import_tab.setLayout(layout)
        self.tab_widget.addTab(import_tab, "Import")
    
    def setup_preview_tab(self):
        """Configure l'onglet de prévisualisation"""
        preview_tab = QWidget()
        layout = QVBoxLayout()
        
        # Informations
        info_layout = QHBoxLayout()
        self.info_label = QLabel("Aucune donnée à prévisualiser")
        info_layout.addWidget(self.info_label)
        info_layout.addStretch()
        layout.addLayout(info_layout)
        
        # Table de prévisualisation
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(6)
        self.preview_table.setHorizontalHeaderLabels([
            "Projet ID", "Année", "Direction", "Catégorie", "Mois", "Jours"
        ])
        layout.addWidget(self.preview_table)
        
        preview_tab.setLayout(layout)
        self.tab_widget.addTab(preview_tab, "Prévisualisation")
    
    def load_projects(self):
        """Charge la liste des projets"""
        projects = self.mapper.get_available_projects()
        self.project_combo.clear()
        for project_id, code, nom in projects:
            self.project_combo.addItem(f"{code} - {nom}", project_id)
    
    def preselect_project(self, project_id):
        """Pré-sélectionne un projet dans la liste"""
        for i in range(self.project_combo.count()):
            if self.project_combo.itemData(i) == project_id:
                self.project_combo.setCurrentIndex(i)
                break
    
    def load_configs(self):
        """Charge la liste des configurations"""
        configs = self.config_manager.list_configs()
        self.config_combo.clear()
        self.config_combo.addItem("-- Nouvelle configuration --", None)
        for config in configs:
            self.config_combo.addItem(config, config)
    
    def new_config(self):
        """Crée une nouvelle configuration"""
        self.config_combo.setCurrentIndex(0)
        self.clear_config_form()
    
    def clear_config_form(self):
        """Vide le formulaire de configuration"""
        self.config_name_edit.clear()
        self.description_edit.clear()
        self.structure_combo.setCurrentText("rows")
        self.direction_col_edit.clear()
        self.categorie_col_edit.clear()
        self.annee_col_edit.clear()
        self.mois_col_edit.clear()
        self.jours_col_edit.clear()
        
        # Nettoyer les colonnes de mois
        while self.month_columns_layout.count() > 1:  # Garder le bouton "Ajouter"
            child = self.month_columns_layout.takeAt(1)
            if child.widget():
                child.widget().deleteLater()
    
    def save_config(self):
        """Sauvegarde la configuration actuelle"""
        name = self.config_name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Erreur", "Nom de configuration requis")
            return
        
        config = self.build_config_from_form()
        
        if self.config_manager.save_config(name, config):
            QMessageBox.information(self, "Succès", f"Configuration '{name}' sauvegardée")
            self.load_configs()
            # Sélectionner la config sauvegardée
            index = self.config_combo.findData(name)
            if index >= 0:
                self.config_combo.setCurrentIndex(index)
        else:
            QMessageBox.critical(self, "Erreur", "Impossible de sauvegarder la configuration")
    
    def build_config_from_form(self) -> Dict:
        """Construit la configuration à partir du formulaire"""
        config = {
            'name': self.config_name_edit.text().strip(),
            'description': self.description_edit.toPlainText().strip(),
            'data_structure': self.structure_combo.currentText(),
            'column_mapping': {
                'direction': {
                    'column': self.direction_col_edit.text().strip() or None
                },
                'categorie': {
                    'column': self.categorie_col_edit.text().strip() or None
                },
                'annee': {
                    'column': self.annee_col_edit.text().strip() or None
                },
                'mois': {
                    'column': self.mois_col_edit.text().strip() or None
                },
                'jours': {
                    'column': self.jours_col_edit.text().strip() or None
                }
            },
            'month_columns': []
        }
        
        # Récupérer les colonnes de mois
        for i in range(1, self.month_columns_layout.count()):
            widget = self.month_columns_layout.itemAt(i).widget()
            if hasattr(widget, 'get_config'):
                month_config = widget.get_config()
                if month_config:
                    config['month_columns'].append(month_config)
        
        return config
    
    def load_selected_config(self):
        """Charge la configuration sélectionnée"""
        config_name = self.config_combo.currentData()
        if config_name:
            config = self.config_manager.load_config(config_name)
            if config:
                self.load_config_to_form(config)
                self.current_config = config
        else:
            self.clear_config_form()
            self.current_config = None
    
    def load_config_to_form(self, config: Dict):
        """Charge une configuration dans le formulaire"""
        self.config_name_edit.setText(config.get('name', ''))
        self.description_edit.setPlainText(config.get('description', ''))
        self.structure_combo.setCurrentText(config.get('data_structure', 'rows'))
        
        column_mapping = config.get('column_mapping', {})
        self.direction_col_edit.setText(column_mapping.get('direction', {}).get('column', ''))
        self.categorie_col_edit.setText(column_mapping.get('categorie', {}).get('column', ''))
        self.annee_col_edit.setText(column_mapping.get('annee', {}).get('column', ''))
        self.mois_col_edit.setText(column_mapping.get('mois', {}).get('column', ''))
        self.jours_col_edit.setText(column_mapping.get('jours', {}).get('column', ''))
        
        # Charger les colonnes de mois
        self.clear_month_columns()
        for month_config in config.get('month_columns', []):
            self.add_month_column(month_config)
    
    def delete_config(self):
        """Supprime la configuration sélectionnée"""
        config_name = self.config_combo.currentData()
        if config_name:
            reply = QMessageBox.question(
                self, "Confirmation", 
                f"Supprimer la configuration '{config_name}' ?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                if self.config_manager.delete_config(config_name):
                    QMessageBox.information(self, "Succès", "Configuration supprimée")
                    self.load_configs()
                else:
                    QMessageBox.critical(self, "Erreur", "Impossible de supprimer la configuration")
    
    def on_structure_changed(self):
        """Réagit au changement de structure de données"""
        if self.structure_combo.currentText() == "columns":
            self.month_columns_group.show()
            self.mois_col_edit.setEnabled(False)
        else:
            self.month_columns_group.hide()
            self.mois_col_edit.setEnabled(True)
    
    def add_month_column(self, config=None):
        """Ajoute une colonne de mois"""
        widget = MonthColumnWidget(config)
        self.month_columns_layout.insertWidget(
            self.month_columns_layout.count() - 1, widget
        )
    
    def clear_month_columns(self):
        """Vide les colonnes de mois"""
        while self.month_columns_layout.count() > 1:
            child = self.month_columns_layout.takeAt(1)
            if child.widget():
                child.widget().deleteLater()
    
    def browse_excel_file(self):
        """Sélection du fichier Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Sélectionner le fichier Excel",
            "", "Fichiers Excel (*.xlsx *.xls);;Tous les fichiers (*)"
        )
        
        if file_path:
            self.file_path_edit.setText(file_path)
            self.load_excel_preview(file_path)
    
    def load_excel_preview(self, file_path):
        """Charge un aperçu du fichier Excel"""
        try:
            # Lire les noms des feuilles
            excel_file = pd.ExcelFile(file_path)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(excel_file.sheet_names)
            
            # Charger la première feuille
            if excel_file.sheet_names:
                self.load_sheet_preview(file_path, excel_file.sheet_names[0])
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de lire le fichier Excel:\n{str(e)}")
    
    def load_sheet_preview(self, file_path, sheet_name):
        """Charge un aperçu d'une feuille spécifique"""
        try:
            # Lire un échantillon de la feuille
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                header=self.header_row_spin.value(),
                skiprows=self.skip_rows_spin.value(),
                nrows=20  # Limite pour l'aperçu
            )
            
            # Afficher dans le tableau
            self.excel_preview_table.setRowCount(len(df))
            self.excel_preview_table.setColumnCount(len(df.columns))
            self.excel_preview_table.setHorizontalHeaderLabels([str(col) for col in df.columns])
            
            for row in range(len(df)):
                for col in range(len(df.columns)):
                    value = df.iloc[row, col]
                    self.excel_preview_table.setItem(
                        row, col, QTableWidgetItem(str(value) if pd.notna(value) else "")
                    )
            
            self.current_df = df
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Impossible de lire la feuille:\n{str(e)}")
    
    def test_configuration(self):
        """Teste la configuration avec le fichier Excel"""
        if not self.file_path_edit.text():
            QMessageBox.warning(self, "Erreur", "Sélectionnez un fichier Excel")
            return
        
        if not self.config_name_edit.text().strip():
            QMessageBox.warning(self, "Erreur", "Définissez un nom de configuration")
            return
        
        projet_id = self.project_combo.currentData()
        if not projet_id:
            QMessageBox.warning(self, "Erreur", "Sélectionnez un projet")
            return
        
        try:
            # Construire la configuration
            config = self.build_config_from_form()
            
            # Lire le fichier Excel complet
            df = pd.read_excel(
                self.file_path_edit.text(),
                sheet_name=self.sheet_combo.currentText(),
                header=self.header_row_spin.value(),
                skiprows=self.skip_rows_spin.value()
            )
            
            # Traiter les données
            results = self.mapper.process_excel_data(df, config, projet_id)
            
            # Afficher la prévisualisation
            self.show_preview(results)
            self.tab_widget.setCurrentIndex(2)  # Aller à l'onglet preview
            
            if results:
                self.import_btn.setEnabled(True)
                QMessageBox.information(
                    self, "Test réussi", 
                    f"Configuration testée avec succès!\n{len(results)} entrées détectées."
                )
            else:
                QMessageBox.warning(
                    self, "Aucune donnée", 
                    "Aucune donnée valide détectée avec cette configuration."
                )
        
        except Exception as e:
            QMessageBox.critical(self, "Erreur de test", f"Erreur lors du test:\n{str(e)}")
            traceback.print_exc()
    
    def show_preview(self, results):
        """Affiche la prévisualisation des données"""
        self.preview_table.setRowCount(len(results))
        self.info_label.setText(f"Prévisualisation: {len(results)} entrées détectées")
        
        for row, entry in enumerate(results):
            self.preview_table.setItem(row, 0, QTableWidgetItem(str(entry['projet_id'])))
            self.preview_table.setItem(row, 1, QTableWidgetItem(str(entry['annee'])))
            self.preview_table.setItem(row, 2, QTableWidgetItem(entry['direction']))
            self.preview_table.setItem(row, 3, QTableWidgetItem(entry['categorie']))
            self.preview_table.setItem(row, 4, QTableWidgetItem(entry['mois']))
            self.preview_table.setItem(row, 5, QTableWidgetItem(str(entry['jours'])))
        
        self.preview_data = results
    
    def perform_import(self):
        """Effectue l'import des données"""
        if not hasattr(self, 'preview_data') or not self.preview_data:
            QMessageBox.warning(self, "Erreur", "Aucune donnée à importer. Testez d'abord la configuration.")
            return
        
        reply = QMessageBox.question(
            self, "Confirmation d'import",
            f"Importer {len(self.preview_data)} entrées dans la base de données ?\n\n"
            "Les données existantes avec les mêmes clés seront remplacées.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            success, message = self.mapper.save_to_database(self.preview_data)
            
            if success:
                QMessageBox.information(self, "Import réussi", message)
                self.import_btn.setEnabled(False)
                self.preview_data = []
            else:
                QMessageBox.critical(self, "Erreur d'import", message)

class MonthColumnWidget(QWidget):
    """Widget pour configurer une colonne de mois"""
    
    def __init__(self, config=None):
        super().__init__()
        layout = QHBoxLayout()
        
        layout.addWidget(QLabel("Colonne:"))
        self.column_edit = QLineEdit()
        layout.addWidget(self.column_edit)
        
        layout.addWidget(QLabel("Mois:"))
        self.month_combo = QComboBox()
        self.month_combo.addItems([
            "01", "02", "03", "04", "05", "06",
            "07", "08", "09", "10", "11", "12"
        ])
        layout.addWidget(self.month_combo)
        
        remove_btn = QPushButton("X")
        remove_btn.setMaximumWidth(30)
        remove_btn.clicked.connect(self.remove_self)
        layout.addWidget(remove_btn)
        
        self.setLayout(layout)
        
        if config:
            self.column_edit.setText(config.get('column', ''))
            month_value = config.get('month', '01')
            if isinstance(month_value, int):
                month_value = f"{month_value:02d}"
            index = self.month_combo.findText(str(month_value))
            if index >= 0:
                self.month_combo.setCurrentIndex(index)
    
    def get_config(self):
        """Retourne la configuration de cette colonne"""
        column = self.column_edit.text().strip()
        if column:
            return {
                'column': column,
                'month': self.month_combo.currentText()
            }
        return None
    
    def remove_self(self):
        """Supprime ce widget"""
        self.deleteLater()

# Fonction principale pour lancer l'interface
def open_excel_import_dialog(parent=None, preselected_project_id=None):
    """Ouvre la boîte de dialogue d'import Excel"""
    dialog = ExcelImportDialog(parent)
    
    # Pré-sélectionner le projet si fourni
    if preselected_project_id:
        dialog.preselect_project(preselected_project_id)
    
    return dialog.exec()

if __name__ == "__main__":
    from PyQt6.QtWidgets import QApplication
    import sys
    
    app = QApplication(sys.argv)
    open_excel_import_dialog()
    sys.exit(app.exec())
