import sqlite3
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, 
                            QTableWidgetItem, QPushButton, QComboBox, QGroupBox, 
                            QCheckBox, QSpinBox, QMessageBox, QFrame)
from PyQt6.QtCore import Qt
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
import datetime

DB_PATH = 'gestion_budget.db'

# Vérifier si la table cir_coeffs existe, sinon la créer
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Vérifier et créer la table cir_coeffs
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='cir_coeffs'")
    if not cursor.fetchone():
        cursor.execute('''
            CREATE TABLE cir_coeffs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                annee INTEGER,
                k1 REAL,
                k2 REAL,
                k3 REAL
            )
        ''')
    
    # Vérifier et créer la table amortissements
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='amortissements'")
    if not cursor.fetchone():
        cursor.execute('''
            CREATE TABLE amortissements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                annee INTEGER,
                mois TEXT,
                montant REAL,
                detail TEXT,
                FOREIGN KEY (projet_id) REFERENCES projets (id)
            )
        ''')
    
    conn.commit()
    conn.close()

class CompteResultatDialog(QDialog):
    def __init__(self, parent, projet_id):
        super().__init__(parent)
        self.parent = parent
        self.projet_id = projet_id
        
        # Vérifier que le projet existe
        if not self.projet_id:
            QMessageBox.critical(self, "Erreur", "Aucun projet sélectionné!")
            self.reject()
            return
            
        # Récupérer les informations du projet
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT code, nom FROM projets WHERE id = ?", (self.projet_id,))
        projet = cursor.fetchone()
        conn.close()
        
        if not projet:
            QMessageBox.critical(self, "Erreur", "Projet non trouvé!")
            self.reject()
            return
            
        self.projet_code, self.projet_nom = projet
        
        self.setWindowTitle(f"Compte de Résultat - {self.projet_code}")
        self.setMinimumSize(800, 600)
        
        # Initialiser les variables pour stocker les montants
        self.depenses = {
            'achats_sous_traitance': 0,
            'autres_achats': 0,
            'cout_direct': 0,
            'dotation_amortissements': 0
        }
        self.recettes = {
            'recettes': 0,
            'subventions': 0
        }
        self.charges_financieres = 0
        self.charges_exceptionnelles = 0  # Crédit d'impôt (en négatif)
        
        # Créer l'interface
        self.init_ui()
        
        # Charger les données initiales
        self.refresh_data()
        
    def init_ui(self):
        """Initialise l'interface utilisateur"""
        layout = QVBoxLayout(self)
        
        # En-tête avec informations du projet
        header_layout = QHBoxLayout()
        project_info = QLabel(f"<b>Projet:</b> {self.projet_code} - {self.projet_nom}")
        header_layout.addWidget(project_info)
        layout.addLayout(header_layout)
        
        # Filtre de période
        filter_group = QGroupBox("Période")
        filter_layout = QHBoxLayout()
        
        self.period_type = QComboBox()
        self.period_type.addItems(["Année complète", "Mois spécifique"])
        filter_layout.addWidget(QLabel("Type de période:"))
        filter_layout.addWidget(self.period_type)
        
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2000, 2100)
        self.year_spin.setValue(datetime.datetime.now().year)
        filter_layout.addWidget(QLabel("Année:"))
        filter_layout.addWidget(self.year_spin)
        
        self.month_combo = QComboBox()
        months = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", 
                 "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        self.month_combo.addItems(months)
        self.month_combo.setCurrentIndex(datetime.datetime.now().month - 1)
        filter_layout.addWidget(QLabel("Mois:"))
        filter_layout.addWidget(self.month_combo)
        self.month_combo.setVisible(False)  # Caché par défaut
        
        refresh_btn = QPushButton("Actualiser")
        refresh_btn.clicked.connect(self.refresh_data)
        filter_layout.addWidget(refresh_btn)
        
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
        
        # Tableau du compte de résultat
        self.result_table = QTableWidget(0, 2)
        self.result_table.setHorizontalHeaderLabels(["Poste", "Montant (€)"])
        layout.addWidget(self.result_table)
        
        # Options supplémentaires
        options_layout = QHBoxLayout()
        
        options_layout.addWidget(QLabel("Pourcentage des frais d'amortissements:"))
        self.charges_fin_percent = QSpinBox()
        self.charges_fin_percent.setRange(0, 100)
        self.charges_fin_percent.setValue(0)
        self.charges_fin_percent.setSuffix("%")
        self.charges_fin_percent.valueChanged.connect(self.update_charges_financieres)
        options_layout.addWidget(self.charges_fin_percent)
        
        options_layout.addStretch()
        
        layout.addLayout(options_layout)
        
        # Boutons d'action
        button_layout = QHBoxLayout()
        
        print_btn = QPushButton("Imprimer")
        print_btn.clicked.connect(self.print_report)
        
        close_btn = QPushButton("Fermer")
        close_btn.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(print_btn)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        # Connexions des signaux
        self.period_type.currentIndexChanged.connect(self.toggle_month_visibility)
        
    def toggle_month_visibility(self, index):
        """Affiche ou masque le sélecteur de mois en fonction du type de période"""
        self.month_combo.setVisible(index == 1)  # Visible si "Mois spécifique" est sélectionné
    
    def get_period_filter(self):
        """Retourne les conditions SQL et paramètres pour le filtre de période"""
        if self.period_type.currentIndex() == 0:  # Année complète
            return "annee = ?", [self.year_spin.value()]
        else:  # Mois spécifique
            month_name = self.month_combo.currentText()
            return "annee = ? AND mois = ?", [self.year_spin.value(), month_name]
    
    def refresh_data(self):
        """Actualise les données du compte de résultat"""
        # Réinitialiser les variables
        self.depenses = {
            'achats_sous_traitance': 0,
            'autres_achats': 0,
            'cout_direct': 0,
            'dotation_amortissements': 0
        }
        self.recettes = {
            'recettes': 0,
            'subventions': 0
        }
        self.charges_financieres = 0
        self.charges_exceptionnelles = 0
        
        # Récupérer le filtre de période
        period_condition, period_params = self.get_period_filter()
        
        # Collecter les données
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        try:
            # 1. Achats et sous-traitance (dépenses externes)
            cursor.execute(f"""
                SELECT SUM(montant) FROM depenses 
                WHERE projet_id = ? AND {period_condition} 
                AND detail LIKE '%sous-traitance%' OR detail LIKE '%externe%'
            """, [self.projet_id] + period_params)
            result = cursor.fetchone()[0]
            self.depenses['achats_sous_traitance'] = result if result else 0
            
            # 2. Autres achats
            cursor.execute(f"""
                SELECT SUM(montant) FROM depenses 
                WHERE projet_id = ? AND {period_condition} 
                AND detail NOT LIKE '%sous-traitance%' AND detail NOT LIKE '%externe%'
            """, [self.projet_id] + period_params)
            result = cursor.fetchone()[0]
            self.depenses['autres_achats'] = result if result else 0
            
            # 3. Coût direct (depuis temps_travail et categorie_cout)
            cursor.execute(f"""
                SELECT t.categorie, SUM(t.jours), c.cout_production
                FROM temps_travail t
                JOIN categorie_cout c ON t.annee = c.annee AND t.categorie = c.categorie
                WHERE t.projet_id = ? AND {period_condition}
                GROUP BY t.categorie
            """, [self.projet_id] + period_params)
            
            for categorie, jours, cout_production in cursor.fetchall():
                if cout_production:
                    self.depenses['cout_direct'] += jours * cout_production
            
            # 4. Dotation aux amortissements
            # Vérifier si la table existe
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='amortissements'")
            if cursor.fetchone():
                cursor.execute(f"""
                    SELECT SUM(montant) FROM amortissements 
                    WHERE projet_id = ? AND {period_condition}
                """, [self.projet_id] + period_params)
                result = cursor.fetchone()[0]
                self.depenses['dotation_amortissements'] = result if result else 0
            
            # 5. Recettes
            cursor.execute(f"""
                SELECT SUM(montant) FROM recettes 
                WHERE projet_id = ? AND {period_condition} 
                AND detail NOT LIKE '%subvention%'
            """, [self.projet_id] + period_params)
            result = cursor.fetchone()[0]
            self.recettes['recettes'] = result if result else 0
            
            # 6. Subventions
            cursor.execute(f"""
                SELECT SUM(montant) FROM recettes 
                WHERE projet_id = ? AND {period_condition} 
                AND detail LIKE '%subvention%'
            """, [self.projet_id] + period_params)
            result = cursor.fetchone()[0]
            self.recettes['subventions'] = result if result else 0
            
            # 7. Crédit d'impôt (charges exceptionnelles, en négatif)
            cursor.execute(f"""
                SELECT k1, k2, k3 FROM cir_coeffs 
                WHERE annee IN (SELECT DISTINCT annee FROM temps_travail 
                               WHERE projet_id = ? AND {period_condition})
            """, [self.projet_id] + period_params)
            
            cir_coeffs = cursor.fetchone()
            if cir_coeffs and all(cir_coeffs):
                k1, k2, k3 = cir_coeffs
                
                # Calcul du crédit d'impôt en fonction des jours éligibles
                cursor.execute(f"""
                    SELECT SUM(jours) FROM temps_travail 
                    WHERE projet_id = ? AND {period_condition}
                """, [self.projet_id] + period_params)
                
                total_jours = cursor.fetchone()[0] or 0
                
                # Le crédit d'impôt est calculé comme un pourcentage des coûts éligibles
                jours_eligibles = total_jours * k1
                credit_impot = jours_eligibles * k3
                
                # En négatif car c'est un gain
                self.charges_exceptionnelles = -credit_impot
            
            # 8. Charges financières (calculées en fonction du pourcentage)
            self.update_charges_financieres()
            
            # Afficher les résultats
            self.display_compte_resultat()
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la récupération des données: {str(e)}")
        
        finally:
            conn.close()
    
    def update_charges_financieres(self):
        """Met à jour les charges financières en fonction du pourcentage défini"""
        percent = self.charges_fin_percent.value() / 100.0
        self.charges_financieres = self.depenses['dotation_amortissements'] * percent
        
        # Mettre à jour l'affichage si le tableau existe déjà
        if hasattr(self, 'result_table') and self.result_table.rowCount() > 0:
            self.display_compte_resultat()
    
    def display_compte_resultat(self):
        """Affiche le compte de résultat dans le tableau"""
        # Effacer le tableau
        self.result_table.setRowCount(0)
        
        # Style pour les titres de section
        title_style = "background-color: #f0f0f0; font-weight: bold;"
        
        # Fonction pour ajouter une ligne au tableau
        def add_row(label, value=None, style=None):
            row = self.result_table.rowCount()
            self.result_table.insertRow(row)
            
            item_label = QTableWidgetItem(label)
            if style:
                item_label.setBackground(Qt.GlobalColor.lightGray)
                font = item_label.font()
                font.setBold(True)
                item_label.setFont(font)
            
            self.result_table.setItem(row, 0, item_label)
            
            if value is not None:
                item_value = QTableWidgetItem(f"{value:.2f}")
                self.result_table.setItem(row, 1, item_value)
        
        # SECTION DÉPENSES
        add_row("DÉPENSES", style=title_style)
        add_row("Achats et sous-traitance", self.depenses['achats_sous_traitance'])
        add_row("Autres achats", self.depenses['autres_achats'])
        add_row("Coût direct", self.depenses['cout_direct'])
        add_row("Dotation aux amortissements", self.depenses['dotation_amortissements'])
        
        # Total dépenses
        total_depenses = sum(self.depenses.values())
        add_row("TOTAL DÉPENSES", total_depenses, style=title_style)
        
        # SECTION RECETTES
        add_row("RECETTES", style=title_style)
        add_row("Recettes", self.recettes['recettes'])
        add_row("Subvention", self.recettes['subventions'])
        
        # Total recettes
        total_recettes = sum(self.recettes.values())
        add_row("TOTAL RECETTES", total_recettes, style=title_style)
        
        # Charges financières
        add_row("Charges financières", self.charges_financieres)
        
        # Charges exceptionnelles (crédit d'impôt)
        add_row("Charges exceptionnelles (crédit d'impôt)", self.charges_exceptionnelles)
        
        # Calcul du résultat
        somme_charges = total_depenses + self.charges_financieres + self.charges_exceptionnelles
        add_row("Somme des charges", somme_charges)
        add_row("Somme des produits", total_recettes)
        
        # Résultat financier
        resultat = total_recettes - somme_charges
        add_row("RÉSULTAT FINANCIER", resultat, style=title_style)
        
        # Ajuster la taille des colonnes
        self.result_table.resizeColumnsToContents()
        
        # Ajouter une ligne de séparation horizontale
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
    
    def print_report(self):
        """Imprime le compte de résultat"""
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self)
        
        if dialog.exec() == QPrintDialog.DialogCode.Accepted:
            QMessageBox.information(self, "Impression", "Impression en cours...")
            # Pour une implémentation complète, il faudrait utiliser QPainter

def show_compte_resultat(parent, projet_id):
    """Affiche la fenêtre de compte de résultat"""
    # Initialiser la base de données si nécessaire
    init_db()
    
    dialog = CompteResultatDialog(parent, projet_id)
    dialog.exec()
