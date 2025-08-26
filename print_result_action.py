import sqlite3
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, 
                            QTableWidgetItem, QPushButton, QComboBox, QGroupBox, 
                            QCheckBox, QListWidget, QListWidgetItem, QDateEdit,
                            QRadioButton, QButtonGroup, QGridLayout, QSpinBox,
                            QSplitter, QWidget, QMessageBox, QScrollArea, QFrame,
                            QFileDialog)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QColor
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
import datetime

DB_PATH = 'gestion_budget.db'

class PrintConfigDialog(QDialog):
    def __init__(self, parent, projet_id=None):
        super().__init__(parent)
        self.parent = parent
        self.projet_id = projet_id
        self.setWindowTitle("Configuration du Compte de Résultat")
        self.setMinimumSize(600, 400)
        
        layout = QVBoxLayout()
        
        # Titre
        title = QLabel("Configuration du Compte de Résultat")
        title.setStyleSheet("font-size: 16pt; font-weight: bold; margin-bottom: 20px;")
        layout.addWidget(title)
        
        # Sélection du projet
        project_group = QGroupBox("Sélection du projet")
        project_layout = QVBoxLayout()
        
        self.project_combo = QComboBox()
        self.load_projects()
        project_layout.addWidget(QLabel("Projet:"))
        project_layout.addWidget(self.project_combo)
        
        project_group.setLayout(project_layout)
        layout.addWidget(project_group)
        
        # Sélection de la période
        period_group = QGroupBox("Période")
        period_layout = QVBoxLayout()
        
        # Type de période
        self.period_type = QComboBox()
        self.period_type.addItems(["Année complète", "Mois spécifique", "Période complète du projet"])
        self.period_type.currentIndexChanged.connect(self.update_period_widgets)
        period_layout.addWidget(QLabel("Type de période:"))
        period_layout.addWidget(self.period_type)
        
        # Sélection de l'année
        year_layout = QHBoxLayout()
        year_layout.addWidget(QLabel("Année:"))
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2000, 2100)
        self.year_spin.setValue(datetime.datetime.now().year)
        year_layout.addWidget(self.year_spin)
        year_layout.addStretch()
        period_layout.addLayout(year_layout)
        
        # Sélection du mois
        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("Mois:"))
        self.month_combo = QComboBox()
        months = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", 
                 "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        self.month_combo.addItems(months)
        self.month_combo.setCurrentIndex(datetime.datetime.now().month - 1)
        month_layout.addWidget(self.month_combo)
        month_layout.addStretch()
        
        self.month_widget = QWidget()
        self.month_widget.setLayout(month_layout)
        period_layout.addWidget(self.month_widget)
        
        period_group.setLayout(period_layout)
        layout.addWidget(period_group)
        
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
        self.update_period_widgets(0)
    
    def load_projects(self):
        """Charge les projets depuis la base de données"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT id, code, nom FROM projets ORDER BY code")
        projects = cursor.fetchall()
        conn.close()
        
        self.project_combo.clear()
        for project_id, code, name in projects:
            self.project_combo.addItem(f"{code} - {name}", project_id)
            
            # Présélectionner le projet courant si défini
            if project_id == self.projet_id:
                self.project_combo.setCurrentIndex(self.project_combo.count() - 1)
    
    def update_period_widgets(self, index):
        """Met à jour l'affichage selon le type de période sélectionné"""
        if index == 0:  # Année complète
            self.month_widget.hide()
        elif index == 1:  # Mois spécifique
            self.month_widget.show()
        else:  # Période complète du projet
            self.month_widget.hide()
    
    def get_selected_project_id(self):
        """Retourne l'ID du projet sélectionné"""
        return self.project_combo.currentData()
    
    def get_selected_period(self):
        """Retourne les informations de période sélectionnée"""
        project_id = self.get_selected_project_id()
        year = self.year_spin.value()
        
        if self.period_type.currentIndex() == 0:  # Année complète
            return project_id, year, None
        elif self.period_type.currentIndex() == 1:  # Mois spécifique
            month = self.month_combo.currentIndex() + 1
            return project_id, year, month
        else:  # Période complète du projet
            return project_id, None, "complete_project"
    
    def generate_compte_resultat(self):
        """Génère le compte de résultat"""
        project_id, year, month = self.get_selected_period()
        
        if project_id is None:
            QMessageBox.warning(self, "Avertissement", "Veuillez sélectionner un projet.")
            return
        
        # Déterminer le type de période
        if month == "complete_project":
            period_type = "complete_project"
            year = None
            month = None
        else:
            period_type = None
        
        # Créer et afficher la fenêtre de compte de résultat
        compte_resultat_dialog = CompteResultatDialog(self.parent, project_id, year, month, period_type)
        compte_resultat_dialog.exec()
        self.accept()

class CompteResultatDialog(QDialog):
    def __init__(self, parent, projet_id, year, month=None, period_type=None):
        super().__init__(parent)
        self.projet_id = projet_id
        self.year = year
        self.month = month
        self.period_type = period_type
        
        # Configuration de la fenêtre
        if period_type == "complete_project":
            period_str = "Période complète du projet"
        elif month:
            period_str = f"{month:02d}/{year}"
        else:
            period_str = str(year)
        self.setWindowTitle(f"Compte de Résultat - {period_str}")
        self.setMinimumSize(800, 600)
        
        layout = QVBoxLayout()
        
        # En-tête avec titre et boutons d'actions
        header_layout = QHBoxLayout()
        
        # Titre
        project_name = self.get_project_name()
        title = QLabel(f"COMPTE DE RÉSULTAT\n{project_name}\nPériode: {period_str}")
        title.setStyleSheet("font-size: 14pt; font-weight: bold; margin-bottom: 10px;")
        header_layout.addWidget(title)
        
        header_layout.addStretch()
        
        # Boutons d'actions
        excel_btn = QPushButton("Export Excel")
        excel_btn.clicked.connect(self.export_to_excel)
        header_layout.addWidget(excel_btn)
        
        pdf_btn = QPushButton("Export PDF")
        pdf_btn.clicked.connect(self.export_to_pdf)
        header_layout.addWidget(pdf_btn)
        
        print_btn = QPushButton("Imprimer")
        print_btn.clicked.connect(self.print_compte_resultat)
        header_layout.addWidget(print_btn)
        
        layout.addLayout(header_layout)
        
        # Configuration du pourcentage d'amortissement
        config_layout = QHBoxLayout()
        config_layout.addWidget(QLabel("Pourcentage d'amortissement à appliquer:"))
        
        self.amortissement_percent = QSpinBox()
        self.amortissement_percent.setRange(0, 100)
        self.amortissement_percent.setValue(10)
        self.amortissement_percent.setSuffix("%")
        self.amortissement_percent.valueChanged.connect(self.refresh_compte_resultat)
        config_layout.addWidget(self.amortissement_percent)
        
        refresh_btn = QPushButton("Actualiser")
        refresh_btn.clicked.connect(self.refresh_compte_resultat)
        config_layout.addWidget(refresh_btn)
        config_layout.addStretch()
        
        layout.addLayout(config_layout)
        
        # Tableau du compte de résultat
        self.compte_table = QTableWidget(0, 2)
        self.compte_table.setHorizontalHeaderLabels(["Poste", "Montant (€)"])
        self.compte_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.compte_table)
        
        # Bouton fermer
        close_layout = QHBoxLayout()
        close_layout.addStretch()
        close_btn = QPushButton("Fermer")
        close_btn.clicked.connect(self.accept)
        close_layout.addWidget(close_btn)
        layout.addLayout(close_layout)
        
        self.setLayout(layout)
        
        # Générer le compte de résultat
        self.refresh_compte_resultat()
    
    def get_project_name(self):
        """Récupère le nom du projet"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT code, nom FROM projets WHERE id = ?", (self.projet_id,))
        result = cursor.fetchone()
        conn.close()
        
        if result:
            code, nom = result
            return f"{code} - {nom}"
        return "Projet inconnu"
    
    def refresh_compte_resultat(self):
        """Actualise le compte de résultat"""
        try:
            # Collecter les données
            donnees = self.collecter_donnees_compte_resultat()
            
            # Afficher dans le tableau
            self.display_compte_resultat(donnees)
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la génération du compte de résultat: {str(e)}")
    
    def collecter_donnees_compte_resultat(self):
        """Collecte toutes les données nécessaires pour le compte de résultat"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        try:
            # Conditions de filtrage selon la période
            if self.period_type == "complete_project":
                # Période complète du projet - pas de filtre de date
                condition = "1=1"
                params = []
            elif self.month:
                # Filtre par mois et année - utiliser les noms de mois français
                months_fr = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", 
                           "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
                month_name = months_fr[self.month - 1]
                
                condition = "annee = ? AND mois = ?"
                params = [self.year, month_name]
            else:
                # Filtre par année seulement
                condition = "annee = ?"
                params = [self.year]
            
            donnees = {
                'recettes': 0,
                'subventions': 0,
                'depenses_externes': 0,
                'autres_achats': 0,
                'cout_direct': 0,
                'dotation_amortissements': 0,
                'credit_impot': 0
            }
            
            # 1. RECETTES
            cursor.execute(f"""
                SELECT COALESCE(SUM(montant), 0) FROM recettes 
                WHERE projet_id = ? AND {condition}
            """, [self.projet_id] + params)
            donnees['recettes'] = cursor.fetchone()[0]
            
            # 2. SUBVENTIONS - cette table n'a pas de structure temporelle, on prend tout
            if self.period_type == "complete_project":
                cursor.execute("""
                    SELECT COALESCE(SUM(cd), 0) FROM subventions 
                    WHERE projet_id = ?
                """, [self.projet_id])
            else:
                # Pour les subventions, on considère qu'elles s'appliquent sur toute la période du projet
                cursor.execute("""
                    SELECT COALESCE(SUM(cd), 0) FROM subventions 
                    WHERE projet_id = ?
                """, [self.projet_id])
            donnees['subventions'] = cursor.fetchone()[0]
            
            # 3. DÉPENSES EXTERNES (table depenses)
            cursor.execute(f"""
                SELECT COALESCE(SUM(montant), 0) FROM depenses 
                WHERE projet_id = ? AND {condition}
            """, [self.projet_id] + params)
            donnees['depenses_externes'] = cursor.fetchone()[0]
            
            # 4. AUTRES ACHATS (table autres_depenses)
            cursor.execute(f"""
                SELECT COALESCE(SUM(montant), 0) FROM autres_depenses 
                WHERE projet_id = ? AND {condition}
            """, [self.projet_id] + params)
            donnees['autres_achats'] = cursor.fetchone()[0]
            
            # 5. COÛT DIRECT (temps_travail * cout_production)
            cursor.execute(f"""
                SELECT COALESCE(SUM(t.jours * c.cout_production), 0)
                FROM temps_travail t
                JOIN categorie_cout c ON t.categorie = c.categorie AND t.annee = c.annee
                WHERE t.projet_id = ? AND {condition.replace('annee', 't.annee').replace('mois', 't.mois')}
            """, [self.projet_id] + params)
            
            donnees['cout_direct'] = cursor.fetchone()[0]
            
            # 6. DOTATION AMORTISSEMENTS (investissements)
            # Pour les investissements, on regarde la date d'achat
            if self.period_type == "complete_project":
                cursor.execute("""
                    SELECT COALESCE(SUM(montant / duree), 0) FROM investissements 
                    WHERE projet_id = ?
                """, [self.projet_id])
            else:
                # Pour simplifier, on prend les amortissements de l'année
                cursor.execute("""
                    SELECT COALESCE(SUM(montant / duree), 0) FROM investissements 
                    WHERE projet_id = ? AND strftime('%Y', date_achat) = ?
                """, [self.projet_id, str(self.year)])
            
            donnees['dotation_amortissements'] = cursor.fetchone()[0]
            
            # 7. CRÉDIT D'IMPÔT - Calculé selon les coefficients CIR
            # Pour simplifier, on met 0 pour l'instant
            donnees['credit_impot'] = 0
            
            return donnees
            
        finally:
            conn.close()
    
    def display_compte_resultat(self, donnees):
        """Affiche le compte de résultat dans le tableau"""
        self.compte_table.setRowCount(0)
        
        # PRODUITS D'EXPLOITATION
        self.add_table_row("=== PRODUITS D'EXPLOITATION ===", "", bold=True)
        self.add_table_row("Recettes", f"{donnees['recettes']:.2f}")
        self.add_table_row("Subventions d'exploitation", f"{donnees['subventions']:.2f}")
        
        total_produits_exploitation = donnees['recettes'] + donnees['subventions']
        self.add_table_row("TOTAL PRODUITS D'EXPLOITATION", f"{total_produits_exploitation:.2f}", bold=True)
        
        self.add_table_row("", "")  # Ligne vide
        
        # CHARGES D'EXPLOITATION
        self.add_table_row("=== CHARGES D'EXPLOITATION ===", "", bold=True)
        self.add_table_row("Dépenses externes", f"{donnees['depenses_externes']:.2f}")
        self.add_table_row("Autres achats", f"{donnees['autres_achats']:.2f}")
        self.add_table_row("Coût direct du personnel", f"{donnees['cout_direct']:.2f}")
        self.add_table_row("Dotation aux amortissements", f"{donnees['dotation_amortissements']:.2f}")
        
        total_charges_exploitation = (donnees['depenses_externes'] + donnees['autres_achats'] + 
                                    donnees['cout_direct'] + donnees['dotation_amortissements'])
        self.add_table_row("TOTAL CHARGES D'EXPLOITATION", f"{total_charges_exploitation:.2f}", bold=True)
        
        self.add_table_row("", "")  # Ligne vide
        
        # RÉSULTAT D'EXPLOITATION
        resultat_exploitation = total_produits_exploitation - total_charges_exploitation
        self.add_table_row("RÉSULTAT D'EXPLOITATION", f"{resultat_exploitation:.2f}", bold=True, 
                          color="green" if resultat_exploitation >= 0 else "red")
        
        self.add_table_row("", "")  # Ligne vide
        
        # PRODUITS EXCEPTIONNELS
        if donnees['credit_impot'] > 0:
            self.add_table_row("=== PRODUITS EXCEPTIONNELS ===", "", bold=True)
            self.add_table_row("Crédit d'impôt recherche", f"{donnees['credit_impot']:.2f}")
            self.add_table_row("TOTAL PRODUITS EXCEPTIONNELS", f"{donnees['credit_impot']:.2f}", bold=True)
            self.add_table_row("", "")  # Ligne vide
        
        # RÉSULTAT NET
        resultat_net = resultat_exploitation + donnees['credit_impot']
        self.add_table_row("RÉSULTAT NET", f"{resultat_net:.2f}", bold=True, 
                          color="green" if resultat_net >= 0 else "red", size=16)
        
        # Ajuster les dimensions du tableau
        self.compte_table.resizeColumnsToContents()
        self.compte_table.resizeRowsToContents()
    
    def add_table_row(self, poste, montant, bold=False, color=None, size=None):
        """Ajoute une ligne au tableau"""
        row = self.compte_table.rowCount()
        self.compte_table.insertRow(row)
        
        # Cellule poste
        poste_item = QTableWidgetItem(poste)
        if bold:
            font = poste_item.font()
            font.setBold(True)
            if size:
                font.setPointSize(size)
            poste_item.setFont(font)
        if color:
            if color == "green":
                poste_item.setBackground(QColor(144, 238, 144))  # lightgreen
            elif color == "red":
                poste_item.setBackground(QColor(255, 182, 193))  # lightcoral
        
        # Cellule montant
        montant_item = QTableWidgetItem(montant)
        montant_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        if bold:
            font = montant_item.font()
            font.setBold(True)
            if size:
                font.setPointSize(size)
            montant_item.setFont(font)
        if color:
            if color == "green":
                montant_item.setBackground(QColor(144, 238, 144))  # lightgreen
            elif color == "red":
                montant_item.setBackground(QColor(255, 182, 193))  # lightcoral
        
        self.compte_table.setItem(row, 0, poste_item)
        self.compte_table.setItem(row, 1, montant_item)
    
    def export_to_excel(self):
        """Exporte le compte de résultat vers Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, PatternFill
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Exporter le compte de résultat",
                f"compte_resultat_{self.year}_{self.month or 'annuel'}.xlsx",
                "Fichiers Excel (*.xlsx)"
            )
            
            if not file_path:
                return
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Compte de Résultat"
            
            # Styles
            title_font = Font(bold=True, size=14)
            header_font = Font(bold=True, size=12)
            bold_font = Font(bold=True)
            
            # Titre
            project_name = self.get_project_name()
            period_str = f"{self.month:02d}/{self.year}" if self.month else str(self.year)
            
            ws['A1'] = f"COMPTE DE RÉSULTAT - {project_name}"
            ws['A1'].font = title_font
            ws['A2'] = f"Période: {period_str}"
            ws['A2'].font = header_font
            
            # En-têtes
            ws['A4'] = "Poste"
            ws['B4'] = "Montant (€)"
            ws['A4'].font = bold_font
            ws['B4'].font = bold_font
            
            # Données
            row = 5
            for i in range(self.compte_table.rowCount()):
                poste_item = self.compte_table.item(i, 0)
                montant_item = self.compte_table.item(i, 1)
                
                if poste_item and montant_item:
                    ws[f'A{row}'] = poste_item.text()
                    ws[f'B{row}'] = montant_item.text()
                    
                    # Appliquer le style gras si nécessaire
                    if poste_item.font().bold():
                        ws[f'A{row}'].font = bold_font
                        ws[f'B{row}'].font = bold_font
                
                row += 1
            
            # Ajuster les colonnes
            ws.column_dimensions['A'].width = 35
            ws.column_dimensions['B'].width = 15
            
            wb.save(file_path)
            QMessageBox.information(self, "Export réussi", f"Fichier exporté: {file_path}")
            
        except ImportError:
            QMessageBox.critical(self, "Erreur", "Le module openpyxl n'est pas installé.\nInstallez-le avec: pip install openpyxl")
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'export", f"Erreur lors de l'export Excel: {str(e)}")
    
    def export_to_pdf(self):
        """Exporte le compte de résultat vers PDF"""
        try:
            from PyQt6.QtPrintSupport import QPrinter
            from PyQt6.QtGui import QTextDocument
            from PyQt6.QtCore import QMarginsF
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Exporter le compte de résultat",
                f"compte_resultat_{self.year}_{self.month or 'annuel'}.pdf",
                "Fichiers PDF (*.pdf)"
            )
            
            if not file_path:
                return
            
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(file_path)
            printer.setPageMargins(QMarginsF(15, 15, 15, 15), QPrinter.Unit.Millimeter)
            
            # Générer le HTML
            html_content = self.generate_html()
            
            document = QTextDocument()
            document.setHtml(html_content)
            document.print(printer)
            
            QMessageBox.information(self, "Export réussi", f"Fichier exporté: {file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'export", f"Erreur lors de l'export PDF: {str(e)}")
    
    def print_compte_resultat(self):
        """Imprime le compte de résultat"""
        try:
            from PyQt6.QtPrintSupport import QPrintDialog, QPrinter
            from PyQt6.QtGui import QTextDocument
            
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            
            dialog = QPrintDialog(printer, self)
            if dialog.exec() != QPrintDialog.DialogCode.Accepted:
                return
            
            # Générer le HTML et imprimer
            html_content = self.generate_html()
            
            document = QTextDocument()
            document.setHtml(html_content)
            document.print(printer)
            
            QMessageBox.information(self, "Impression", "Document envoyé à l'imprimante.")
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'impression", f"Erreur lors de l'impression: {str(e)}")
    
    def generate_html(self):
        """Génère le contenu HTML pour l'export PDF/impression"""
        project_name = self.get_project_name()
        period_str = f"{self.month:02d}/{self.year}" if self.month else str(self.year)
        
        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                h1 {{ text-align: center; color: #333; }}
                h2 {{ text-align: center; color: #666; font-size: 14pt; }}
                table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
                th, td {{ border: 1px solid #ddd; padding: 10px; text-align: left; }}
                th {{ background-color: #f2f2f2; font-weight: bold; }}
                .bold {{ font-weight: bold; }}
                .amount {{ text-align: right; }}
                .total {{ background-color: #e8f4f8; font-weight: bold; }}
                .result {{ background-color: #d4edda; font-weight: bold; font-size: 14pt; }}
            </style>
        </head>
        <body>
            <h1>COMPTE DE RÉSULTAT</h1>
            <h2>{project_name}</h2>
            <h2>Période: {period_str}</h2>
            
            <table>
                <tr>
                    <th>Poste</th>
                    <th>Montant (€)</th>
                </tr>
        """
        
        # Ajouter les données du tableau
        for i in range(self.compte_table.rowCount()):
            poste_item = self.compte_table.item(i, 0)
            montant_item = self.compte_table.item(i, 1)
            
            if poste_item and montant_item:
                poste = poste_item.text()
                montant = montant_item.text()
                
                # Déterminer le style CSS
                css_class = ""
                if poste_item.font().bold():
                    if "RÉSULTAT NET" in poste:
                        css_class = "result"
                    elif "TOTAL" in poste or "RÉSULTAT" in poste:
                        css_class = "total"
                    else:
                        css_class = "bold"
                
                html += f'<tr class="{css_class}"><td>{poste}</td><td class="amount">{montant}</td></tr>'
        
        html += """
            </table>
        </body>
        </html>
        """
        
        return html

def show_print_config_dialog(parent, projet_id=None):
    """Fonction d'entrée pour afficher la configuration d'impression"""
    dialog = PrintConfigDialog(parent, projet_id)
    return dialog.exec()