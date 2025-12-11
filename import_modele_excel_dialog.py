"""
Script d'import de données depuis un fichier Excel modèle vers la base de données
Import les données de Temps de travail, Dépenses externes, Autres dépenses et Recettes
"""

import pandas as pd
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QComboBox, QFileDialog, QMessageBox,
                             QProgressBar, QTextEdit, QGroupBox)
from PyQt6.QtCore import Qt
from database import get_connection
import traceback


class ImportExcelModeleDialog(QDialog):
    """
    Dialogue pour importer les données depuis un fichier Excel modèle
    """
    
    def __init__(self, parent=None, projet_id=None):
        super().__init__(parent)
        self.setWindowTitle("Import depuis Excel Modèle")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        
        self.excel_file = None
        self.projet_id = projet_id
        self.data_temps = None
        self.data_depenses = None
        self.data_autres = None
        self.data_recettes = None
        
        self.init_ui()
    
    def init_ui(self):
        """Initialise l'interface utilisateur"""
        layout = QVBoxLayout()
        
        # Section fichier Excel
        file_group = QGroupBox("1. Sélection du fichier Excel")
        file_layout = QHBoxLayout()
        
        self.file_label = QLabel("Aucun fichier sélectionné")
        self.file_label.setStyleSheet("color: #666;")
        
        self.btn_browse = QPushButton("Parcourir...")
        self.btn_browse.clicked.connect(self.browse_file)
        
        file_layout.addWidget(self.file_label, 1)
        file_layout.addWidget(self.btn_browse)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # Section aperçu
        preview_group = QGroupBox("2. Aperçu des données à importer")
        preview_layout = QVBoxLayout()
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(200)
        preview_layout.addWidget(self.preview_text)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # Barre de progression
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Section log
        log_group = QGroupBox("Journal d'import")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # Boutons d'action
        btn_layout = QHBoxLayout()
        
        self.btn_import = QPushButton("Importer")
        self.btn_import.clicked.connect(self.import_data)
        self.btn_import.setEnabled(False)
        self.btn_import.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px; font-weight: bold;")
        
        self.btn_close = QPushButton("Fermer")
        self.btn_close.clicked.connect(self.close)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_import)
        btn_layout.addWidget(self.btn_close)
        
        layout.addLayout(btn_layout)
        self.setLayout(layout)
    

    
    def browse_file(self):
        """Ouvre le dialogue de sélection de fichier"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Sélectionner le fichier Excel",
            "",
            "Fichiers Excel (*.xlsx *.xls)"
        )
        
        if file_path:
            self.excel_file = file_path
            self.file_label.setText(file_path)
            self.file_label.setStyleSheet("color: #000;")
            self.load_excel_data()
            self.check_ready_to_import()
    
    def load_excel_data(self):
        """Charge et valide les données du fichier Excel"""
        try:
            self.log("Chargement du fichier Excel...")
            
            # Charger les différentes feuilles
            excel_data = {}
            
            # Feuille Temps de Travail
            try:
                df_temps = pd.read_excel(self.excel_file, sheet_name="Temps de Travail")
                # Filtrer les lignes de description (ligne 1 avec les descriptions)
                df_temps = df_temps[df_temps['Année'].notna()]
                # Filtrer les descriptions qui contiennent "Année" ou "format"
                df_temps = df_temps[~df_temps['Année'].astype(str).str.contains('Année|format|ex:', case=False, na=False)]
                self.data_temps = df_temps
                excel_data['temps'] = len(df_temps)
            except Exception as e:
                self.data_temps = None
                excel_data['temps'] = 0
                self.log(f"⚠ Feuille 'Temps de Travail' non chargée : {e}")
            
            # Feuille Dépenses Externes
            try:
                df_depenses = pd.read_excel(self.excel_file, sheet_name="Dépenses Externes", header=0)
                df_depenses = df_depenses[df_depenses['Année'].notna()]
                df_depenses = df_depenses[~df_depenses['Année'].astype(str).str.contains('Année|format|ex:', case=False, na=False)]
                self.data_depenses = df_depenses
                excel_data['depenses'] = len(df_depenses)
            except Exception as e:
                self.data_depenses = None
                excel_data['depenses'] = 0
                self.log(f"⚠ Feuille 'Dépenses Externes' non chargée : {e}")
            
            # Feuille Autres Dépenses
            try:
                df_autres = pd.read_excel(self.excel_file, sheet_name="Autres Dépenses")
                df_autres = df_autres[df_autres['Année'].notna()]
                df_autres = df_autres[~df_autres['Année'].astype(str).str.contains('Année|format|ex:', case=False, na=False)]
                self.data_autres = df_autres
                excel_data['autres'] = len(df_autres)
            except Exception as e:
                self.data_autres = None
                excel_data['autres'] = 0
                self.log(f"⚠ Feuille 'Autres Dépenses' non chargée : {e}")
            
            # Feuille Recettes
            try:
                df_recettes = pd.read_excel(self.excel_file, sheet_name="Recettes")
                df_recettes = df_recettes[df_recettes['Année'].notna()]
                df_recettes = df_recettes[~df_recettes['Année'].astype(str).str.contains('Année|format|ex:', case=False, na=False)]
                self.data_recettes = df_recettes
                excel_data['recettes'] = len(df_recettes)
            except Exception as e:
                self.data_recettes = None
                excel_data['recettes'] = 0
                self.log(f"⚠ Feuille 'Recettes' non chargée : {e}")
            
            # Afficher l'aperçu
            preview = "Données détectées dans le fichier Excel :\n\n"
            preview += f"• Temps de Travail : {excel_data['temps']} lignes\n"
            preview += f"• Dépenses Externes : {excel_data['depenses']} lignes\n"
            preview += f"• Autres Dépenses : {excel_data['autres']} lignes\n"
            preview += f"• Recettes : {excel_data['recettes']} lignes\n"
            preview += f"\nTotal : {sum(excel_data.values())} lignes à importer"
            
            self.preview_text.setText(preview)
            self.log("✓ Fichier chargé avec succès")
            
        except Exception as e:
            self.log(f"✗ Erreur lors du chargement : {e}")
            QMessageBox.critical(self, "Erreur", f"Erreur lors du chargement du fichier :\n{e}\n\n{traceback.format_exc()}")
    
    def check_ready_to_import(self):
        """Vérifie si tout est prêt pour l'import"""
        ready = (self.excel_file is not None and 
                 (self.data_temps is not None or 
                  self.data_depenses is not None or 
                  self.data_autres is not None or 
                  self.data_recettes is not None))
        
        self.btn_import.setEnabled(ready)
    
    def log(self, message):
        """Ajoute un message au journal"""
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def import_data(self):
        """Importe les données dans la base de données"""
        try:
            self.log("\n=== DÉBUT DE L'IMPORT ===")
            self.progress_bar.setVisible(True)
            self.btn_import.setEnabled(False)
            
            total_inserted = 0
            total_errors = 0
            
            conn = get_connection()
            cursor = conn.cursor()
            
            # Import Temps de Travail
            if self.data_temps is not None and len(self.data_temps) > 0:
                self.log(f"\nImport de {len(self.data_temps)} lignes de Temps de Travail...")
                inserted, errors = self.import_temps_travail(cursor)
                total_inserted += inserted
                total_errors += errors
                self.log(f"✓ {inserted} lignes insérées, {errors} erreurs")
            
            # Import Dépenses Externes
            if self.data_depenses is not None and len(self.data_depenses) > 0:
                self.log(f"\nImport de {len(self.data_depenses)} lignes de Dépenses Externes...")
                inserted, errors = self.import_depenses(cursor)
                total_inserted += inserted
                total_errors += errors
                self.log(f"✓ {inserted} lignes insérées, {errors} erreurs")
            
            # Import Autres Dépenses
            if self.data_autres is not None and len(self.data_autres) > 0:
                self.log(f"\nImport de {len(self.data_autres)} lignes d'Autres Dépenses...")
                inserted, errors = self.import_autres_depenses(cursor)
                total_inserted += inserted
                total_errors += errors
                self.log(f"✓ {inserted} lignes insérées, {errors} erreurs")
            
            # Import Recettes
            if self.data_recettes is not None and len(self.data_recettes) > 0:
                self.log(f"\nImport de {len(self.data_recettes)} lignes de Recettes...")
                inserted, errors = self.import_recettes(cursor)
                total_inserted += inserted
                total_errors += errors
                self.log(f"✓ {inserted} lignes insérées, {errors} erreurs")
            
            conn.commit()
            conn.close()
            
            self.log(f"\n=== IMPORT TERMINÉ ===")
            self.log(f"Total : {total_inserted} lignes insérées, {total_errors} erreurs")
            
            self.progress_bar.setVisible(False)
            
            QMessageBox.information(
                self,
                "Import terminé",
                f"Import terminé avec succès !\n\n"
                f"Lignes insérées : {total_inserted}\n"
                f"Erreurs : {total_errors}"
            )
            
            # Fermer le dialog après un import réussi
            self.accept()
            
        except Exception as e:
            self.log(f"\n✗ ERREUR GRAVE : {e}")
            self.log(traceback.format_exc())
            QMessageBox.critical(self, "Erreur", f"Erreur lors de l'import :\n{e}")
        finally:
            self.progress_bar.setVisible(False)
            self.btn_import.setEnabled(True)
    
    def import_temps_travail(self, cursor):
        """Importe les données de temps de travail"""
        inserted = 0
        errors = 0
        
        for idx, row in self.data_temps.iterrows():
            try:
                annee = int(row['Année'])
                direction = str(row['Direction']).strip()
                categorie = str(row['Catégorie']).strip()
                membre_id = str(row['Membre ID']).strip()
                mois = str(row['Mois']).strip()
                # Accepter point ou virgule comme séparateur décimal
                jours = float(str(row['Jours']).replace(',', '.'))
                
                # Vérifier si l'entrée existe déjà
                cursor.execute("""
                    SELECT COUNT(*) FROM temps_travail 
                    WHERE projet_id = ? AND annee = ? AND membre_id = ? AND mois = ?
                """, (self.projet_id, annee, membre_id, mois))
                
                exists = cursor.fetchone()[0] > 0
                
                if exists:
                    # Mettre à jour
                    cursor.execute("""
                        UPDATE temps_travail 
                        SET direction = ?, categorie = ?, jours = ?
                        WHERE projet_id = ? AND annee = ? AND membre_id = ? AND mois = ?
                    """, (direction, categorie, jours, self.projet_id, annee, membre_id, mois))
                else:
                    # Insérer
                    cursor.execute("""
                        INSERT INTO temps_travail 
                        (projet_id, annee, direction, categorie, membre_id, mois, jours)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (self.projet_id, annee, direction, categorie, membre_id, mois, jours))
                
                inserted += 1
                
            except Exception as e:
                errors += 1
                self.log(f"  ✗ Erreur ligne {idx + 3}: {e}")
        
        return inserted, errors
    
    def import_depenses(self, cursor):
        """Importe les dépenses externes"""
        inserted = 0
        errors = 0
        
        for idx, row in self.data_depenses.iterrows():
            try:
                annee = int(row['Année'])
                mois = str(row['Mois']).strip()
                detail = str(row['Libellé']).strip()
                # Accepter point ou virgule comme séparateur décimal
                montant = float(str(row['Montant']).replace(',', '.'))
                categorie = "Dépenses Externes"  # Catégorie par défaut
                
                cursor.execute("""
                    INSERT INTO depenses (projet_id, annee, categorie, mois, montant, detail)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (self.projet_id, annee, categorie, mois, montant, detail))
                
                inserted += 1
                
            except Exception as e:
                errors += 1
                self.log(f"  ✗ Erreur ligne {idx + 3}: {e}")
        
        return inserted, errors
    
    def import_autres_depenses(self, cursor):
        """Importe les autres dépenses"""
        inserted = 0
        errors = 0
        
        for idx, row in self.data_autres.iterrows():
            try:
                annee = int(row['Année'])
                mois = str(row['Mois']).strip()
                detail = str(row['Libellé']).strip()
                # Accepter point ou virgule comme séparateur décimal
                montant = float(str(row['Montant']).replace(',', '.'))
                ligne_index = idx + 1  # Utiliser l'index de ligne
                
                cursor.execute("""
                    INSERT INTO autres_depenses (projet_id, annee, ligne_index, mois, montant, detail)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (self.projet_id, annee, ligne_index, mois, montant, detail))
                
                inserted += 1
                
            except Exception as e:
                errors += 1
                self.log(f"  ✗ Erreur ligne {idx + 3}: {e}")
        
        return inserted, errors
    
    def import_recettes(self, cursor):
        """Importe les recettes"""
        inserted = 0
        errors = 0
        
        for idx, row in self.data_recettes.iterrows():
            try:
                annee = int(row['Année'])
                mois = str(row['Mois']).strip()
                detail = str(row['Libellé']).strip()
                # Accepter point ou virgule comme séparateur décimal
                montant = float(str(row['Montant']).replace(',', '.'))
                ligne_index = idx + 1  # Utiliser l'index de ligne
                
                cursor.execute("""
                    INSERT INTO recettes (projet_id, annee, ligne_index, mois, montant, detail)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (self.projet_id, annee, ligne_index, mois, montant, detail))
                
                inserted += 1
                
            except Exception as e:
                errors += 1
                self.log(f"  ✗ Erreur ligne {idx + 3}: {e}")
        
        return inserted, errors


def show_import_dialog(parent=None):
    """Fonction utilitaire pour afficher le dialogue d'import"""
    dialog = ImportExcelModeleDialog(parent)
    dialog.exec()
