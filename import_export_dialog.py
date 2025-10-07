from PyQt6.QtWidgets import QDialog, QVBoxLayout, QPushButton, QLabel, QCheckBox, QHBoxLayout, QFileDialog, QMessageBox, QListWidget, QRadioButton, QButtonGroup
from PyQt6.QtCore import Qt
import sqlite3
import os
import shutil

from database import DB_PATH, get_connection

class ImportExportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Gestion de la base de données')
        self.setMinimumWidth(400)
        layout = QVBoxLayout()

        # Section Export
        layout.addWidget(QLabel('Export de la base de données :'))
        
        # Options d'export avec boutons radio
        self.global_radio = QRadioButton('Exporter toute la base de données')
        self.global_radio.setChecked(True)  # Option par défaut
        self.specific_projects_radio = QRadioButton('Exporter seulement certains projets')
        self.custom_export_radio = QRadioButton('Export personnalisé (tables spécifiques)')
        
        # Groupe de boutons pour l'exclusivité
        self.option_group = QButtonGroup()
        self.option_group.addButton(self.global_radio)
        self.option_group.addButton(self.specific_projects_radio)
        self.option_group.addButton(self.custom_export_radio)

        layout.addWidget(self.global_radio)
        layout.addWidget(self.specific_projects_radio)
        layout.addWidget(self.custom_export_radio)

        # Liste des projets pour sélection (désactivée par défaut)
        self.project_list = QListWidget()
        self.project_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.project_list.setEnabled(False)  # Désactivée par défaut
        self.project_label = QLabel('Sélectionner les projets à exporter :')
        layout.addWidget(self.project_label)
        layout.addWidget(self.project_list)

        # Charger les projets dans la liste
        self.load_projects()

        # Liste des tables pour export personnalisé (désactivée par défaut)
        self.table_list = QListWidget()
        self.table_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.table_list.setEnabled(False)  # Désactivée par défaut
        self.table_label = QLabel('Sélectionner les tables de données à exporter :')
        layout.addWidget(self.table_label)
        layout.addWidget(self.table_list)

        # Charger les tables dans la liste
        self.load_tables()

        # Connecter les boutons radio pour activer/désactiver les listes
        self.specific_projects_radio.toggled.connect(self.toggle_project_list)
        self.custom_export_radio.toggled.connect(self.toggle_table_list)

        # Bouton Export
        self.btn_export = QPushButton('Exporter')
        layout.addWidget(self.btn_export)

        # Séparateur
        separator = QLabel('─' * 50)
        separator.setStyleSheet("color: gray;")
        layout.addWidget(separator)
        
        # Section Import
        layout.addWidget(QLabel('Import de la base de données :'))
        import_info = QLabel('L\'import fusionnera intelligemment les données :\n• Met à jour les projets existants\n• Ajoute les nouveaux projets')
        import_info.setStyleSheet("color: #666; font-size: 11px; margin: 5px 0px;")
        layout.addWidget(import_info)
        
        self.btn_import = QPushButton('Importer')
        layout.addWidget(self.btn_import)

        # Séparateur
        separator2 = QLabel('─' * 50)
        separator2.setStyleSheet("color: gray;")
        layout.addWidget(separator2)

        # Section Suppression (en rouge)
        layout.addWidget(QLabel('Suppression de la base de données :'))
        danger_info = QLabel('⚠️ Cette action supprimera définitivement toutes les données !')
        danger_info.setStyleSheet("color: red; font-weight: bold; margin: 5px 0px;")
        layout.addWidget(danger_info)
        
        self.btn_delete = QPushButton('Supprimer la base de données')
        self.btn_delete.setStyleSheet("background-color: #dc3545; color: white; font-weight: bold; padding: 8px;")
        layout.addWidget(self.btn_delete)

        self.setLayout(layout)

        # Connexions des boutons
        self.btn_import.clicked.connect(self.import_database)
        self.btn_export.clicked.connect(self.export_database)
        self.btn_delete.clicked.connect(self.delete_database)

    def toggle_project_list(self, checked):
        """Active ou désactive la liste des projets en fonction du bouton radio."""
        self.project_list.setEnabled(checked)

    def toggle_table_list(self, checked):
        """Active ou désactive la liste des tables en fonction du bouton radio."""
        self.table_list.setEnabled(checked)

    def load_projects(self):
        """Charge les projets dans la liste avec leur nom et code."""
        try:
            with get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT code, nom FROM projets ORDER BY nom')
                for code, nom in cursor.fetchall():
                    self.project_list.addItem(f"{code} - {nom}")
        except Exception as e:
            print(f"Erreur lors du chargement des projets: {e}")

    def load_tables(self):
        """Charge toutes les tables de la base de données dans la liste."""
        try:
            with get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")
                tables = [row[0] for row in cursor.fetchall()]

                for table in tables:
                    self.table_list.addItem(table)
        except Exception as e:
            print(f"Erreur lors du chargement des tables: {e}")

    def get_selected_projects(self):
        """Récupère les projets sélectionnés par l'utilisateur."""
        selected_items = self.project_list.selectedItems()
        return [item.text().split(' - ')[0] for item in selected_items]  # Retourne uniquement les codes des projets

    def get_selected_tables(self):
        """Récupère les tables sélectionnées par l'utilisateur."""
        selected_items = self.table_list.selectedItems()
        return [item.text() for item in selected_items]  # Retourne les noms des tables directement

    def import_database(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Sélectionner un fichier de base de données', '', 'Base de données SQLite (*.db)')
        if file_path:
            try:
                # Toujours fusionner intelligemment
                self.merge_databases(file_path, DB_PATH)
                QMessageBox.information(self, 'Succès', 'Base de données importée et fusionnée avec succès.')
                
                # Rafraîchir l'interface parent si possible
                if hasattr(self.parent(), 'load_projects'):
                    self.parent().load_projects()
                    
            except Exception as e:
                QMessageBox.critical(self, 'Erreur', f'Échec de l\'importation : {e}')

    def delete_database(self):
        """Supprime complètement la base de données après confirmation"""
        # Double confirmation pour éviter les accidents
        reply = QMessageBox.question(
            self, 
            'Confirmation de suppression', 
            '⚠️ ATTENTION ⚠️\n\nVous êtes sur le point de supprimer définitivement toute la base de données.\n\nCette action est IRRÉVERSIBLE !\n\nÊtes-vous absolument certain de vouloir continuer ?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # Seconde confirmation
            reply2 = QMessageBox.question(
                self,
                'Dernière confirmation',
                'Tapez "SUPPRIMER" dans le champ suivant pour confirmer définitivement :',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply2 == QMessageBox.StandardButton.Yes:
                try:
                    # Fermer toutes les connexions existantes
                    if os.path.exists(DB_PATH):
                        os.remove(DB_PATH)
                        QMessageBox.information(self, 'Suppression effectuée', 'La base de données a été supprimée avec succès.')
                        
                        # Fermer la fenêtre et rafraîchir l'interface parent
                        self.accept()
                        if hasattr(self.parent(), 'load_projects'):
                            self.parent().load_projects()
                    else:
                        QMessageBox.information(self, 'Information', 'Aucune base de données à supprimer.')
                        
                except Exception as e:
                    QMessageBox.critical(self, 'Erreur', f'Impossible de supprimer la base de données : {e}')

    def merge_databases(self, source_db_path, target_db_path):
        """Fusionne la base de données source avec la base cible"""
        source_conn = sqlite3.connect(source_db_path)
        target_conn = sqlite3.connect(target_db_path)
        
        source_cursor = source_conn.cursor()
        target_cursor = target_conn.cursor()
        
        # Obtenir la liste de toutes les tables
        source_cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [row[0] for row in source_cursor.fetchall()]
        
        for table_name in tables:
            try:
                if table_name == 'projets':
                    # Traitement spécial pour la table projets
                    self.merge_projects_table(source_cursor, target_cursor)
                else:
                    # Pour les autres tables, comportement normal
                    source_cursor.execute(f"SELECT * FROM {table_name}")
                    rows = source_cursor.fetchall()
                    
                    if rows:
                        placeholders = ', '.join(['?' for _ in range(len(rows[0]))])
                        for row in rows:
                            target_cursor.execute(f"INSERT OR REPLACE INTO {table_name} VALUES ({placeholders})", row)
                        
            except Exception as e:
                print(f"Erreur lors de la fusion de la table {table_name}: {e}")
                continue
        
        target_conn.commit()
        source_conn.close()
        target_conn.close()

    def merge_projects_table(self, source_cursor, target_cursor):
        """Fusionne intelligemment la table projets"""
        source_cursor.execute("SELECT * FROM projets")
        source_projects = source_cursor.fetchall()
        
        for project_row in source_projects:
            # Structure: id, code, nom, details, date_debut, date_fin, livrables, chef, etat, cir, subvention, theme_principal
            old_id, code, nom = project_row[0], project_row[1], project_row[2]
            
            # Vérifier si un projet avec même code existe (plus restrictif)
            target_cursor.execute("SELECT id, nom FROM projets WHERE code=?", (code,))
            existing = target_cursor.fetchone()
            
            if existing:
                existing_id, existing_nom = existing
                # Si même code ET même nom → Remplacer
                if existing_nom == nom:
                    # Remplacer en gardant l'ID existant
                    target_cursor.execute("""UPDATE projets SET 
                        nom=?, details=?, date_debut=?, date_fin=?, livrables=?, 
                        chef=?, etat=?, cir=?, subvention=?, theme_principal=? 
                        WHERE id=?""", 
                        project_row[2:] + (existing_id,))
                    
                    # Supprimer les anciennes données liées et les remplacer
                    self.replace_project_related_data(source_cursor, target_cursor, old_id, existing_id)
                else:
                    # Même code mais nom différent → Créer nouveau avec code modifié
                    new_code = f"{code}_imported"
                    # Vérifier que le nouveau code n'existe pas déjà
                    counter = 1
                    while True:
                        target_cursor.execute("SELECT id FROM projets WHERE code=?", (new_code,))
                        if not target_cursor.fetchone():
                            break
                        new_code = f"{code}_imported_{counter}"
                        counter += 1
                    
                    # Insérer avec le nouveau code
                    new_row = (new_code,) + project_row[2:]
                    placeholders = ', '.join(['?' for _ in range(len(new_row))])
                    target_cursor.execute(f"INSERT INTO projets (code, nom, details, date_debut, date_fin, livrables, chef, etat, cir, subvention, theme_principal) VALUES ({placeholders})", new_row)
                    
                    # Récupérer le nouvel ID généré
                    new_id = target_cursor.lastrowid
                    
                    # Copier les données liées avec le nouvel ID
                    self.copy_project_related_data(source_cursor, target_cursor, old_id, new_id)
            else:
                # Nouveau projet → Créer avec nouvel ID
                new_row = project_row[1:]  # Exclut l'ancien ID
                placeholders = ', '.join(['?' for _ in range(len(new_row))])
                target_cursor.execute(f"INSERT INTO projets (code, nom, details, date_debut, date_fin, livrables, chef, etat, cir, subvention, theme_principal) VALUES ({placeholders})", new_row)
                
                # Récupérer le nouvel ID généré
                new_id = target_cursor.lastrowid
                
                # Copier les données liées avec le nouvel ID
                self.copy_project_related_data(source_cursor, target_cursor, old_id, new_id)

    def replace_project_related_data(self, source_cursor, target_cursor, old_id, existing_id):
        """Remplace les données liées d'un projet existant"""
        related_tables = ['equipe', 'investissements', 'subventions', 'images', 'projet_themes', 
                         'temps_travail', 'recettes', 'depenses', 'autres_depenses', 'actualites']
        
        for table_name in related_tables:
            try:
                # Supprimer les anciennes données
                target_cursor.execute(f"DELETE FROM {table_name} WHERE projet_id=?", (existing_id,))
                
                # Copier les nouvelles données
                self.copy_project_related_data(source_cursor, target_cursor, old_id, existing_id, [table_name])
            except Exception as e:
                print(f"Erreur lors du remplacement de {table_name}: {e}")

    def copy_project_related_data(self, source_cursor, target_cursor, old_id, new_id, tables=None):
        """Copie les données liées d'un projet avec un nouvel ID"""
        if tables is None:
            tables = ['equipe', 'investissements', 'subventions', 'images', 'projet_themes', 
                     'temps_travail', 'recettes', 'depenses', 'autres_depenses', 'actualites']
        
        for table_name in tables:
            try:
                source_cursor.execute(f"SELECT * FROM {table_name} WHERE projet_id=?", (old_id,))
                related_rows = source_cursor.fetchall()
                
                for row in related_rows:
                    # Remplacer l'ancien projet_id par le nouveau
                    new_row = (new_id,) + row[1:]  # Remplace le premier élément (projet_id)
                    placeholders = ', '.join(['?' for _ in range(len(new_row))])
                    target_cursor.execute(f"INSERT INTO {table_name} VALUES ({placeholders})", new_row)
            except Exception as e:
                print(f"Erreur lors de la copie de {table_name}: {e}")

    def export_database(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            'Enregistrer la base de données',
            os.path.basename(DB_PATH),
            'Base de données SQLite (*.db)'
        )
        if file_path:
            try:
                # Supprimer le fichier s'il existe déjà pour éviter les conflits
                if os.path.exists(file_path):
                    os.remove(file_path)
                
                conn = get_connection()
                backup_conn = sqlite3.connect(file_path)

                if self.global_radio.isChecked():
                    # Exporter toute la base de données
                    with backup_conn:
                        conn.backup(backup_conn)
                        
                elif self.specific_projects_radio.isChecked():
                    # Export par projets spécifiques
                    selected_projects = self.get_selected_projects()
                    if not selected_projects:
                        QMessageBox.warning(self, 'Attention', 'Veuillez sélectionner au moins un projet à exporter.')
                        return
                    self.export_specific_projects(conn, backup_conn, selected_projects)
                    
                elif self.custom_export_radio.isChecked():
                    # Export personnalisé par tables
                    selected_tables = self.get_selected_tables()
                    if not selected_tables:
                        QMessageBox.warning(self, 'Attention', 'Veuillez sélectionner au moins une table à exporter.')
                        return
                    self.export_custom_tables(conn, backup_conn, selected_tables)

                conn.close()
                backup_conn.close()
                QMessageBox.information(self, 'Succès', 'Base de données exportée avec succès.')
            except Exception as e:
                QMessageBox.critical(self, 'Erreur', f'Échec de l\'exportation : {e}')

    def export_specific_projects(self, conn, backup_conn, selected_projects):
        """Exporte uniquement les projets sélectionnés avec toutes leurs données liées"""
        cursor = conn.cursor()
        backup_cursor = backup_conn.cursor()

        # Créer les tables nécessaires dans la base de données exportée (exclure les tables système)
        cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
        for (create_table_sql,) in cursor.fetchall():
            if create_table_sql:  # Éviter les None
                backup_cursor.execute(create_table_sql)

        # Tables à exporter avec leurs relations
        tables_to_export = [
            'themes',  # Thèmes (pas liés directement mais nécessaires)
            'directions',  # Directions (pas liées directement mais nécessaires)
            'chefs_projet',  # Chefs de projet (pas liés directement mais nécessaires)
            'categorie_cout'  # Catégories de coût (pas liées directement mais nécessaires)
        ]
        
        # Exporter d'abord les tables indépendantes
        for table_name in tables_to_export:
            cursor.execute(f'SELECT * FROM {table_name}')
            rows = cursor.fetchall()
            if rows:
                # Construire la requête INSERT dynamiquement
                placeholders = ', '.join(['?' for _ in range(len(rows[0]))])
                for row in rows:
                    backup_cursor.execute(f'INSERT OR IGNORE INTO {table_name} VALUES ({placeholders})', row)

        # Exporter les projets sélectionnés et leurs données liées
        for project_code in selected_projects:
            cursor.execute('SELECT * FROM projets WHERE code=?', (project_code,))
            project_data = cursor.fetchone()
            if project_data:
                projet_id = project_data[0]
                
                # Exporter le projet principal
                placeholders = ', '.join(['?' for _ in range(len(project_data))])
                backup_cursor.execute(f'INSERT INTO projets VALUES ({placeholders})', project_data)
                
                # Exporter toutes les données liées au projet
                related_tables = [
                    'equipe', 'investissements', 'subventions', 'images', 'projet_themes',
                    'temps_travail', 'recettes', 'depenses', 'autres_depenses', 'actualites'
                ]
                
                for table_name in related_tables:
                    try:
                        cursor.execute(f'SELECT * FROM {table_name} WHERE projet_id=?', (projet_id,))
                        related_rows = cursor.fetchall()
                        for row in related_rows:
                            placeholders = ', '.join(['?' for _ in range(len(row))])
                            backup_cursor.execute(f'INSERT INTO {table_name} VALUES ({placeholders})', row)
                    except Exception as e:
                        # Table n'existe peut-être pas, continuer
                        print(f"Erreur lors de l'export de {table_name}: {e}")

        backup_conn.commit()

    def export_custom_tables(self, conn, backup_conn, selected_tables):
        """Exporte uniquement les tables sélectionnées"""
        cursor = conn.cursor()
        backup_cursor = backup_conn.cursor()

        # Créer d'abord toutes les tables sélectionnées
        for table_name in selected_tables:
            try:
                cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
                create_sql = cursor.fetchone()
                if create_sql and create_sql[0]:
                    backup_cursor.execute(create_sql[0])
            except Exception as e:
                print(f"Erreur lors de la création de la table {table_name}: {e}")

        # Exporter les données des tables sélectionnées
        for table_name in selected_tables:
            try:
                cursor.execute(f'SELECT * FROM {table_name}')
                rows = cursor.fetchall()
                if rows:
                    placeholders = ', '.join(['?' for _ in range(len(rows[0]))])
                    for row in rows:
                        backup_cursor.execute(f'INSERT INTO {table_name} VALUES ({placeholders})', row)
            except Exception as e:
                print(f"Erreur lors de l'export de {table_name}: {e}")

        backup_conn.commit()
