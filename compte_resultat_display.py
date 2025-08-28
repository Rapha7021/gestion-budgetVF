import sqlite3
import datetime
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, 
                            QTableWidgetItem, QPushButton, QMessageBox, 
                            QFileDialog, QHeaderView)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog

DB_PATH = 'gestion_budget.db'

class CompteResultatDisplay(QDialog):
    def __init__(self, parent, config_data):
        super().__init__(parent)
        self.parent = parent
        self.config_data = config_data
        
        # Extraction des paramètres de configuration
        self.project_ids = config_data.get('project_ids', [])
        self.years = config_data.get('years', [])
        self.period_type = config_data.get('period_type', 'yearly')  # 'yearly' ou 'monthly'
        self.granularity = config_data.get('granularity', 'yearly')  # 'yearly' ou 'monthly'
        self.cost_type = config_data.get('cost_type', 'cout_production')  # Type de coût sélectionné
        
        self.setWindowTitle("Compte de Résultat")
        self.setMinimumSize(1000, 700)
        
        self.init_ui()
        self.load_data()
    
    def init_ui(self):
        """Initialise l'interface utilisateur"""
        layout = QVBoxLayout()
        
        # En-tête avec titre et informations
        header_layout = self.create_header()
        layout.addLayout(header_layout)
        
        # Pas de section de configuration
        
        # Tableau principal du compte de résultat
        self.table = QTableWidget()
        self.setup_table()
        layout.addWidget(self.table)
        
        # Boutons d'actions
        buttons_layout = self.create_buttons()
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
    
    def create_header(self):
        """Crée l'en-tête avec titre et informations"""
        header_layout = QHBoxLayout()
        
        # Titre principal
        title = QLabel("COMPTE DE RÉSULTAT")
        title.setStyleSheet("font-size: 18pt; font-weight: bold; color: #2c3e50; margin-bottom: 10px;")
        header_layout.addWidget(title)
        
        header_layout.addStretch()
        
        # Informations sur la sélection
        info_text = self.get_selection_info()
        info_label = QLabel(info_text)
        info_label.setStyleSheet("font-size: 10pt; color: #7f8c8d;")
        header_layout.addWidget(info_label)
        
        return header_layout
    
    def get_selection_info(self):
        """Génère le texte d'information sur la sélection"""
        info_lines = []
        
        # Projets
        if len(self.project_ids) == 1:
            project_name = self.get_project_name(self.project_ids[0])
            info_lines.append(f"Projet: {project_name}")
        else:
            info_lines.append(f"Projets: {len(self.project_ids)} sélectionnés")
        
        # Années
        if len(self.years) == 1:
            info_lines.append(f"Année: {self.years[0]}")
        else:
            years_str = f"{min(self.years)} - {max(self.years)}" if len(self.years) > 1 else str(self.years[0])
            info_lines.append(f"Période: {years_str}")
        
        # Granularité
        granularity_text = "Mensuel" if self.granularity == 'monthly' else "Annuel"
        info_lines.append(f"Granularité: {granularity_text}")
        
        # Type de coût
        cost_type_mapping = {
            'montant_charge': 'Montant chargé',
            'cout_production': 'Coût de production',
            'cout_complet': 'Coût complet'
        }
        cost_type_text = cost_type_mapping.get(self.cost_type, self.cost_type)
        info_lines.append(f"Type de coût: {cost_type_text}")
        
        return " | ".join(info_lines)
    
    def get_project_name(self, project_id):
        """Récupère le nom d'un projet"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT code, nom FROM projets WHERE id = ?", (project_id,))
        result = cursor.fetchone()
        conn.close()
        
        if result:
            code, nom = result
            return f"{code} - {nom}"
        return "Projet inconnu"
    
    def get_active_months_for_year(self, year):
        """Détermine les mois actifs pour une année donnée en fonction des dates du projet"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        try:
            # Récupérer les dates de début et fin de tous les projets
            active_months = set()
            
            for project_id in self.project_ids:
                cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id = ?", (project_id,))
                project_info = cursor.fetchone()
                
                if not project_info or not project_info[0] or not project_info[1]:
                    continue
                
                try:
                    debut_projet = datetime.datetime.strptime(project_info[0], '%m/%Y')
                    fin_projet = datetime.datetime.strptime(project_info[1], '%m/%Y')
                    
                    # Si l'année demandée est dans la période du projet
                    if debut_projet.year <= year <= fin_projet.year:
                        # Déterminer le mois de début pour cette année
                        start_month = debut_projet.month if year == debut_projet.year else 1
                        # Déterminer le mois de fin pour cette année
                        end_month = fin_projet.month if year == fin_projet.year else 12
                        
                        # Ajouter tous les mois actifs pour ce projet
                        for month in range(start_month, end_month + 1):
                            active_months.add(month)
                            
                except (ValueError, TypeError):
                    continue
            
            return sorted(active_months)
            
        finally:
            conn.close()
    
    def setup_table(self):
        """Configure le tableau du compte de résultat"""
        # Le nombre de colonnes dépend de la granularité et des années
        if self.granularity == 'monthly':
            # Une colonne par mois actif pour chaque année
            columns = ["Poste"]
            for year in sorted(self.years):
                active_months = self.get_active_months_for_year(year)
                for month in active_months:
                    columns.append(f"{month:02d}/{year}")
            columns.append("TOTAL")
        else:
            # Une colonne par année
            columns = ["Poste"] + [str(year) for year in sorted(self.years)]
            if len(self.years) > 1:
                columns.append("TOTAL")
        
        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)
        
        # Configuration de l'apparence
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        
        # Rendre le tableau en lecture seule
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    
    def create_buttons(self):
        """Crée les boutons d'actions"""
        buttons_layout = QHBoxLayout()
        
        # Export Excel
        excel_btn = QPushButton("Export Excel")
        excel_btn.clicked.connect(self.export_to_excel)
        excel_btn.setStyleSheet("QPushButton { background-color: #27ae60; color: white; font-weight: bold; padding: 8px; }")
        buttons_layout.addWidget(excel_btn)
        
        # Export PDF
        pdf_btn = QPushButton("Export PDF")
        pdf_btn.clicked.connect(self.export_to_pdf)
        pdf_btn.setStyleSheet("QPushButton { background-color: #e74c3c; color: white; font-weight: bold; padding: 8px; }")
        buttons_layout.addWidget(pdf_btn)
        
        # Imprimer
        print_btn = QPushButton("Imprimer")
        print_btn.clicked.connect(self.print_compte_resultat)
        print_btn.setStyleSheet("QPushButton { background-color: #3498db; color: white; font-weight: bold; padding: 8px; }")
        buttons_layout.addWidget(print_btn)
        
        buttons_layout.addStretch()
        
        # Fermer
        close_btn = QPushButton("Fermer")
        close_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(close_btn)
        
        return buttons_layout
    
    def load_data(self):
        """Charge et affiche les données du compte de résultat"""
        try:
            # Collecter toutes les données
            data = self.collect_financial_data()
            
            # Afficher dans le tableau
            self.populate_table(data)
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors du chargement des données: {str(e)}")
    
    def collect_financial_data(self):
        """Collecte toutes les données financières"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        try:
            data = {}
            
            for year in self.years:
                if self.granularity == 'monthly':
                    active_months = self.get_active_months_for_year(year)
                    for month in active_months:
                        period_key = f"{month:02d}/{year}"
                        data[period_key] = self.collect_period_data(cursor, year, month)
                else:
                    data[str(year)] = self.collect_period_data(cursor, year)
            
            return data
            
        finally:
            conn.close()
    
    def collect_period_data(self, cursor, year, month=None):
        """Collecte les données pour une période spécifique"""
        # Conditions de filtrage
        if month:
            # Filtrage mensuel - utilise les noms de mois français
            month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                          "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
            month_condition = f"AND mois = '{month_names[month-1]}'"
            year_condition = f"AND annee = {year}"
        else:
            # Filtrage annuel
            month_condition = ""
            year_condition = f"AND annee = {year}"
        
        project_condition = f"AND projet_id IN ({','.join(['?'] * len(self.project_ids))})"
        
        data = {
            # RECETTES
            'recettes': 0,
            'subventions': 0,
            
            # DÉPENSES
            'achats_sous_traitance': 0,
            'autres_achats': 0,
            'cout_direct': 0,
            'dotation_amortissements': 0,
            
            # CHARGES
            'charges_financieres': 0,
            'credit_impot': 0
        }
        
        # 1. RECETTES - table recettes
        try:
            query = f"""
                SELECT COALESCE(SUM(montant), 0) FROM recettes 
                WHERE 1=1 {year_condition} {month_condition} {project_condition}
            """
            cursor.execute(query, self.project_ids)
            data['recettes'] = cursor.fetchone()[0] or 0
        except sqlite3.OperationalError:
            # Table n'existe pas encore
            data['recettes'] = 0
        
        # 2. SUBVENTIONS - Répartition équitable par année/mois
        try:
            subventions_total = 0
            for project_id in self.project_ids:
                # Récupérer les informations du projet (dates, etc.)
                cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id = ?", (project_id,))
                projet_info = cursor.fetchone()
                
                if not projet_info or not projet_info[0] or not projet_info[1]:
                    continue  # Pas de dates de projet, skip
                
                # Calculer la subvention répartie pour cette période
                subvention_periode = self.calculate_simple_distributed_subvention(
                    cursor, project_id, year, month, projet_info)
                subventions_total += subvention_periode
            
            data['subventions'] = subventions_total
        except sqlite3.OperationalError:
            data['subventions'] = 0
        
        # 3. ACHATS ET SOUS-TRAITANCE - table depenses
        try:
            query = f"""
                SELECT COALESCE(SUM(montant), 0) FROM depenses 
                WHERE 1=1 {year_condition} {month_condition} {project_condition}
            """
            cursor.execute(query, self.project_ids)
            data['achats_sous_traitance'] = cursor.fetchone()[0] or 0
        except sqlite3.OperationalError:
            data['achats_sous_traitance'] = 0
        
        # 4. AUTRES ACHATS - table autres_depenses
        try:
            query = f"""
                SELECT COALESCE(SUM(montant), 0) FROM autres_depenses 
                WHERE 1=1 {year_condition} {month_condition} {project_condition}
            """
            cursor.execute(query, self.project_ids)
            data['autres_achats'] = cursor.fetchone()[0] or 0
        except sqlite3.OperationalError:
            data['autres_achats'] = 0
        
        # 5. COÛT DIRECT - temps_travail * (type de coût sélectionné)
        try:
            query = f"""
                SELECT COALESCE(SUM(t.jours * c.{self.cost_type}), 0)
                FROM temps_travail t
                JOIN categorie_cout c ON t.categorie = c.libelle AND t.annee = c.annee
                WHERE t.annee = {year} {month_condition.replace('mois', 't.mois') if month_condition else ''} 
                AND t.projet_id IN ({','.join(['?'] * len(self.project_ids))})
            """
            cursor.execute(query, self.project_ids)
            data['cout_direct'] = cursor.fetchone()[0] or 0
            
            # Vérification pour debug : si le coût direct est 0, vérifier pourquoi
            if data['cout_direct'] == 0:
                # Compter les entrées de temps de travail
                count_query = f"""
                    SELECT COUNT(*) FROM temps_travail t
                    WHERE t.annee = {year} {month_condition.replace('mois', 't.mois') if month_condition else ''} 
                    AND t.projet_id IN ({','.join(['?'] * len(self.project_ids))})
                """
                cursor.execute(count_query, self.project_ids)
                temps_count = cursor.fetchone()[0]
                
                # Compter les catégories de coût disponibles
                cursor.execute(f"SELECT COUNT(*) FROM categorie_cout WHERE annee = {year}")
                cat_count = cursor.fetchone()[0]
                
                print(f"DEBUG: Coût direct = 0 pour année {year}")
                print(f"  - Entrées temps_travail trouvées: {temps_count}")
                print(f"  - Catégories de coût pour {year}: {cat_count}")
                
                if temps_count > 0:
                    # Vérifier les catégories qui ne matchent pas
                    cursor.execute(f"""
                        SELECT DISTINCT t.categorie 
                        FROM temps_travail t 
                        WHERE t.annee = {year} AND t.projet_id IN ({','.join(['?'] * len(self.project_ids))})
                        AND t.categorie NOT IN (SELECT libelle FROM categorie_cout WHERE annee = {year})
                    """, self.project_ids)
                    missing_cats = [row[0] for row in cursor.fetchall()]
                    if missing_cats:
                        print(f"  - Catégories temps_travail sans correspondance: {missing_cats}")
                        
        except sqlite3.OperationalError as e:
            print(f"Erreur SQL coût direct: {e}")
            data['cout_direct'] = 0
        
        # 6. DOTATION AUX AMORTISSEMENTS
        try:
            amortissements_total = 0
            
            # Calculer les amortissements pour chaque projet
            for project_id in self.project_ids:
                amortissement_projet = self.calculate_amortissement_for_period(
                    cursor, project_id, year, month)
                amortissements_total += amortissement_projet
            
            data['dotation_amortissements'] = amortissements_total
        except sqlite3.OperationalError:
            data['dotation_amortissements'] = 0
        
        # 7. CRÉDIT D'IMPÔT RECHERCHE (CIR) - Calculé selon les coefficients
        try:
            # Récupérer les coefficients CIR pour l'année
            cursor.execute('SELECT k1, k2, k3 FROM cir_coeffs WHERE annee = ?', (year,))
            cir_coeffs = cursor.fetchone()
            
            if cir_coeffs:
                k1, k2, k3 = cir_coeffs
                # CIR = ((Coût_direct × K1 + Amortissements × K2) - Subventions) × K3
                cir_base = (data['cout_direct'] * (k1 or 0)) + (data['dotation_amortissements'] * (k2 or 0))
                cir_assiette = cir_base - data['subventions']  # Soustraire les subventions
                cir_credit = cir_assiette * (k3 or 0)
                # Le crédit d'impôt est négatif (diminue les charges)
                data['credit_impot'] = -abs(cir_credit) if cir_credit > 0 else 0
            else:
                data['credit_impot'] = 0
        except sqlite3.OperationalError:
            data['credit_impot'] = 0
        
        return data
    
    def get_cost_type_label(self):
        """Retourne le libellé du type de coût sélectionné"""
        cost_type_mapping = {
            'montant_charge': 'Montant chargé',
            'cout_production': 'Coût de production',
            'cout_complet': 'Coût complet'
        }
        return cost_type_mapping.get(self.cost_type, 'Coût direct')

    def populate_table(self, data):
        """Remplit le tableau avec les données"""
        # Structure du compte de résultat selon vos spécifications
        structure = [
            ("=== PRODUITS ===", "header"),
            ("Recettes", "recettes"),
            ("Subventions", "subventions"),
            ("TOTAL PRODUITS", "total_produits"),
            ("", "separator"),
            ("=== CHARGES ===", "header"),
            ("Achats et sous-traitance", "achats_sous_traitance"),
            ("Autres achats", "autres_achats"),
            (self.get_cost_type_label(), "cout_direct"),  # Nom dynamique selon le type de coût
            ("Dotation aux amortissements", "dotation_amortissements"),
            ("Crédit d'impôt (négatif)", "credit_impot"),
            ("TOTAL CHARGES", "total_charges"),
            ("", "separator"),
            ("RÉSULTAT FINANCIER", "resultat_financier")
        ]
        
        self.table.setRowCount(len(structure))
        
        # Colonnes de données (exclut la première colonne "Poste")
        data_columns = list(data.keys())
        if len(data_columns) > 1:
            data_columns.append("TOTAL")  # Ajouter une colonne total si plusieurs périodes
        
        for row, (label, data_key) in enumerate(structure):
            # Colonne des postes
            item = QTableWidgetItem(label)
            
            if data_key == "header":
                font = QFont()
                font.setBold(True)
                item.setFont(font)
                item.setBackground(QColor(52, 73, 94))
                item.setForeground(QColor(255, 255, 255))
            elif data_key.startswith("total_") or data_key.startswith("resultat_"):
                font = QFont()
                font.setBold(True)
                item.setFont(font)
                if "resultat" in data_key:
                    item.setBackground(QColor(46, 204, 113) if data_key == "resultat_net" else QColor(52, 152, 219))
                    item.setForeground(QColor(255, 255, 255))
            
            self.table.setItem(row, 0, item)
            
            # Colonnes de données
            for col, period in enumerate(sorted(data.keys()), 1):
                if data_key == "separator" or data_key == "header":
                    # Lignes vides ou en-têtes
                    self.table.setItem(row, col, QTableWidgetItem(""))
                elif data_key.startswith("total_") or data_key.startswith("resultat_"):
                    # Calculs des totaux
                    value = self.calculate_total(data[period], data_key)
                    item = QTableWidgetItem(f"{value:,.2f}" if value != 0 else "")
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    
                    font = QFont()
                    font.setBold(True)
                    item.setFont(font)
                    
                    if "resultat_financier" in data_key:
                        color = QColor(46, 204, 113) if value >= 0 else QColor(231, 76, 60)
                        item.setBackground(color)
                        item.setForeground(QColor(255, 255, 255))
                    elif "total_" in data_key:
                        item.setBackground(QColor(52, 152, 219))
                        item.setForeground(QColor(255, 255, 255))
                    
                    self.table.setItem(row, col, item)
                else:
                    # Données simples
                    value = data[period].get(data_key, 0)
                    if data_key == "credit_impot":
                        # Crédit d'impôt négatif - à implémenter selon vos besoins
                        value = 0  # Pour l'instant
                    
                    item = QTableWidgetItem(f"{value:,.2f}" if value != 0 else "")
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    self.table.setItem(row, col, item)
            
            # Colonne TOTAL si plusieurs périodes
            if len(data) > 1 and col < self.table.columnCount() - 1:
                total_value = self.calculate_row_total(data, data_key)
                if total_value is not None:
                    item = QTableWidgetItem(f"{total_value:,.2f}" if total_value != 0 else "")
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    
                    if data_key.startswith("total_") or data_key.startswith("resultat_"):
                        font = QFont()
                        font.setBold(True)
                        item.setFont(font)
                        
                        if "resultat_financier" in data_key:
                            color = QColor(46, 204, 113) if total_value >= 0 else QColor(231, 76, 60)
                            item.setBackground(color)
                            item.setForeground(QColor(255, 255, 255))
                        elif "total_" in data_key:
                            item.setBackground(QColor(52, 152, 219))
                            item.setForeground(QColor(255, 255, 255))
                    
                    self.table.setItem(row, self.table.columnCount() - 1, item)
    
    def calculate_total(self, period_data, total_type):
        """Calcule les totaux selon le type"""
        if total_type == "total_produits":
            return period_data.get('recettes', 0) + period_data.get('subventions', 0)
        elif total_type == "total_charges":
            return (period_data.get('achats_sous_traitance', 0) + 
                   period_data.get('autres_achats', 0) + 
                   period_data.get('cout_direct', 0) + 
                   period_data.get('dotation_amortissements', 0) + 
                   period_data.get('credit_impot', 0))  # Crédit d'impôt inclus dans les charges
        elif total_type == "resultat_financier":
            total_produits = self.calculate_total(period_data, "total_produits")
            total_charges = self.calculate_total(period_data, "total_charges")
            return total_produits - total_charges
        
        return 0
    
    def calculate_row_total(self, all_data, data_key):
        """Calcule le total d'une ligne sur toutes les périodes"""
        if data_key == "separator" or data_key == "header":
            return None
        
        if data_key.startswith("total_") or data_key.startswith("resultat_"):
            # Pour les totaux calculés, sommer les résultats de chaque période
            total = 0
            for period_data in all_data.values():
                total += self.calculate_total(period_data, data_key)
            return total
        else:
            # Pour les données simples, sommer directement
            total = 0
            for period_data in all_data.values():
                total += period_data.get(data_key, 0)
            return total
    
    def export_to_excel(self):
        """Exporte vers Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, PatternFill
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Exporter le compte de résultat",
                f"compte_resultat_consolide.xlsx",
                "Fichiers Excel (*.xlsx)"
            )
            
            if not file_path:
                return
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Compte de Résultat"
            
            # Exporter les données du tableau
            for row in range(self.table.rowCount()):
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    if item:
                        ws.cell(row + 1, col + 1, item.text())
            
            wb.save(file_path)
            QMessageBox.information(self, "Export réussi", f"Fichier exporté: {file_path}")
            
        except ImportError:
            QMessageBox.critical(self, "Erreur", "Le module openpyxl n'est pas installé.\nInstallez-le avec: pip install openpyxl")
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'export", f"Erreur lors de l'export Excel: {str(e)}")
    
    def export_to_pdf(self):
        """Exporte vers PDF"""
        try:
            from PyQt6.QtGui import QTextDocument
            from PyQt6.QtCore import QMarginsF
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Exporter le compte de résultat",
                f"compte_resultat_consolide.pdf",
                "Fichiers PDF (*.pdf)"
            )
            
            if not file_path:
                return
            
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(file_path)
            printer.setPageMargins(QMarginsF(15, 15, 15, 15), QPrinter.Unit.Millimeter)
            
            html_content = self.generate_html_content()
            
            document = QTextDocument()
            document.setHtml(html_content)
            document.print(printer)
            
            QMessageBox.information(self, "Export réussi", f"Fichier exporté: {file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'export", f"Erreur lors de l'export PDF: {str(e)}")
    
    def print_compte_resultat(self):
        """Imprime le compte de résultat"""
        try:
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            
            dialog = QPrintDialog(printer, self)
            if dialog.exec() != QPrintDialog.DialogCode.Accepted:
                return
            
            html_content = self.generate_html_content()
            
            from PyQt6.QtGui import QTextDocument
            document = QTextDocument()
            document.setHtml(html_content)
            document.print(printer)
            
            QMessageBox.information(self, "Impression", "Document envoyé à l'imprimante.")
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'impression", f"Erreur lors de l'impression: {str(e)}")
    
    def generate_html_content(self):
        """Génère le contenu HTML pour l'export"""
        html = """
        <html>
        <head>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                h1 { text-align: center; color: #2c3e50; }
                h2 { text-align: center; color: #7f8c8d; font-size: 12pt; }
                table { border-collapse: collapse; width: 100%; margin-top: 20px; }
                th, td { border: 1px solid #bdc3c7; padding: 8px; text-align: left; }
                th { background-color: #ecf0f1; font-weight: bold; }
                .header { background-color: #34495e; color: white; font-weight: bold; }
                .total { background-color: #3498db; color: white; font-weight: bold; }
                .result { background-color: #2ecc71; color: white; font-weight: bold; }
                .amount { text-align: right; }
            </style>
        </head>
        <body>
            <h1>COMPTE DE RÉSULTAT CONSOLIDÉ</h1>
            <h2>""" + self.get_selection_info() + """</h2>
            <table>
        """
        
        # En-têtes de colonnes
        html += "<tr>"
        for col in range(self.table.columnCount()):
            header_item = self.table.horizontalHeaderItem(col)
            html += f"<th>{header_item.text() if header_item else ''}</th>"
        html += "</tr>"
        
        # Données
        for row in range(self.table.rowCount()):
            html += "<tr>"
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                value = item.text() if item else ""
                css_class = "amount" if col > 0 else ""
                html += f'<td class="{css_class}">{value}</td>'
            html += "</tr>"
        
        html += """
            </table>
        </body>
        </html>
        """
        
        return html
    
    def calculate_simple_distributed_subvention(self, cursor, project_id, year, month, projet_info):
        """Calcule la subvention répartie équitablement selon la règle simple :
        Subvention totale / Nb mois total du projet
        UTILISE LA MÊME LOGIQUE QUE subvention_dialog.py"""
        import datetime
        
        try:
            # Récupérer les paramètres de subvention pour ce projet
            cursor.execute('''
                SELECT depenses_temps_travail, coef_temps_travail, depenses_externes, coef_externes,
                       depenses_autres_achats, coef_autres_achats, depenses_dotation_amortissements, 
                       coef_dotation_amortissements, cd, taux 
                FROM subventions WHERE projet_id = ?
            ''', (project_id,))
            
            subventions_config = cursor.fetchall()
            if not subventions_config:
                return 0
            
            # Calculer le montant total de subvention sur tout le projet
            montant_total_projet = 0
            
            for subvention in subventions_config:
                (dep_temps, coef_temps, dep_ext, coef_ext, dep_autres, coef_autres, 
                 dep_amort, coef_amort, cd, taux) = subvention
                
                montant_subvention_config = 0
                
                # 1. TEMPS DE TRAVAIL - MÊME LOGIQUE QUE subvention_dialog.py
                if dep_temps and coef_temps:
                    cout_total_temps = self.calculate_temps_travail_total(cursor, project_id)
                    temps_travail = cout_total_temps * cd
                    montant_subvention_config += coef_temps * temps_travail
                
                # 2. DÉPENSES EXTERNES - MÊME LOGIQUE
                if dep_ext and coef_ext:
                    cursor.execute('''
                        SELECT COALESCE(SUM(montant), 0) FROM depenses 
                        WHERE projet_id = ?
                    ''', (project_id,))
                    depenses_ext_total = cursor.fetchone()[0] or 0
                    montant_subvention_config += coef_ext * depenses_ext_total
                
                # 3. AUTRES ACHATS - MÊME LOGIQUE
                if dep_autres and coef_autres:
                    cursor.execute('''
                        SELECT COALESCE(SUM(montant), 0) FROM autres_depenses 
                        WHERE projet_id = ?
                    ''', (project_id,))
                    autres_achats_total = cursor.fetchone()[0] or 0
                    montant_subvention_config += coef_autres * autres_achats_total
                
                # 4. AMORTISSEMENTS - MÊME LOGIQUE QUE subvention_dialog.py
                if dep_amort and coef_amort:
                    amortissements_total = self.calculate_amortissements_total_subvention_style(
                        cursor, project_id, projet_info)
                    montant_subvention_config += coef_amort * amortissements_total
                
                # Appliquer le taux de subvention
                montant_subvention_config = montant_subvention_config * (taux / 100)
                montant_total_projet += montant_subvention_config
            
            # Calculer les dates du projet
            debut_projet = datetime.datetime.strptime(projet_info[0], '%m/%Y')
            fin_projet = datetime.datetime.strptime(projet_info[1], '%m/%Y')
            
            # Calculer le nombre total de mois du projet
            nb_mois_total = (fin_projet.year - debut_projet.year) * 12 + (fin_projet.month - debut_projet.month) + 1
            
            # Vérifier si l'année demandée est dans la période du projet
            if year < debut_projet.year or year > fin_projet.year:
                return 0  # L'année n'est pas dans le projet
            
            if month:
                # Pour un mois : Subvention totale / Nombre total de mois du projet
                return montant_total_projet / nb_mois_total
            else:
                # Pour une année complète : Calculer combien de mois actifs dans cette année
                start_month = debut_projet.month if year == debut_projet.year else 1
                end_month = fin_projet.month if year == fin_projet.year else 12
                nb_mois_annee = end_month - start_month + 1
                
                # Subvention = (Subvention totale / Nb mois total) * Nb mois dans cette année
                return (montant_total_projet / nb_mois_total) * nb_mois_annee
            
        except Exception as e:
            print(f"Erreur dans calculate_simple_distributed_subvention: {e}")
            return 0
    
    def calculate_temps_travail_total(self, cursor, project_id):
        """Calcule le temps de travail total exactement comme dans subvention_dialog.py"""
        try:
            cursor.execute("""
                SELECT tt.annee, tt.categorie, tt.mois, tt.jours 
                FROM temps_travail tt 
                WHERE tt.projet_id = ?
            """, (project_id,))
            
            temps_travail_rows = cursor.fetchall()
            cout_total_temps = 0
            
            for annee, categorie, mois, jours in temps_travail_rows:
                # Convertir la catégorie du temps de travail au format de categorie_cout
                # MÊME LOGIQUE QUE subvention_dialog.py lignes 111-127
                categorie_code = ""
                if "Stagiaire" in categorie:
                    categorie_code = "STP"
                elif "Assistante" in categorie or "opérateur" in categorie:
                    categorie_code = "AOP"
                elif "Technicien" in categorie:
                    categorie_code = "TEP"
                elif "Junior" in categorie:
                    categorie_code = "IJP"
                elif "Senior" in categorie:
                    categorie_code = "ISP"
                elif "Expert" in categorie:
                    categorie_code = "EDP"
                elif "moyen" in categorie:
                    categorie_code = "MOY"
                
                # Si on n'a pas trouvé de correspondance, continuer
                if not categorie_code:
                    continue
                    
                # Récupérer le montant chargé pour cette catégorie et cette année
                cursor.execute("""
                    SELECT montant_charge 
                    FROM categorie_cout 
                    WHERE categorie = ? AND annee = ?
                """, (categorie_code, annee))
                
                cout_row = cursor.fetchone()
                if cout_row and cout_row[0]:
                    montant_charge = float(cout_row[0])
                    cout_total_temps += jours * montant_charge
                else:
                    # Si pas de coût pour cette année/catégorie, utiliser une valeur par défaut
                    cout_total_temps += jours * 500  # 500€ par jour par défaut
            
            return cout_total_temps
            
        except Exception as e:
            print(f"Erreur dans calculate_temps_travail_total: {e}")
            return 0
    
    def calculate_amortissements_total_subvention_style(self, cursor, project_id, projet_info):
        """Calcule les amortissements totaux exactement comme dans subvention_dialog.py"""
        import datetime
        
        try:
            if not projet_info or not projet_info[0] or not projet_info[1]:
                return 0
            
            # Convertir les dates MM/yyyy en objets datetime - MÊME LOGIQUE
            try:
                debut_projet = datetime.datetime.strptime(projet_info[0], '%m/%Y')
                fin_projet = datetime.datetime.strptime(projet_info[1], '%m/%Y')
            except ValueError:
                return 0
            
            cursor.execute("""
                SELECT montant, date_achat, duree 
                FROM investissements 
                WHERE projet_id = ?
            """, (project_id,))
            
            amortissements_total = 0
            
            for montant, date_achat, duree in cursor.fetchall():
                try:
                    # MÊME LOGIQUE QUE subvention_dialog.py lignes 177-213
                    # Convertir la date d'achat en datetime
                    achat_date = datetime.datetime.strptime(date_achat, '%m/%Y')
                    
                    # La dotation commence le mois suivant l'achat
                    debut_amort = achat_date.replace(day=1)
                    debut_amort = datetime.datetime(debut_amort.year, debut_amort.month, 1) + datetime.timedelta(days=32)
                    debut_amort = debut_amort.replace(day=1)
                    
                    # La fin de l'amortissement est soit la fin du projet, soit la fin de la période d'amortissement
                    fin_amort = achat_date.replace(day=1)
                    # Ajouter durée années à la date d'achat
                    fin_amort = datetime.datetime(fin_amort.year + int(duree), fin_amort.month, 1)
                    
                    # Prendre la date la plus proche entre fin du projet et fin d'amortissement
                    fin_effective = min(fin_projet, fin_amort)
                    
                    # Si le début d'amortissement est après la fin du projet, pas d'amortissement
                    if debut_amort > fin_projet:
                        continue
                        
                    # Calculer le nombre de mois d'amortissement effectif
                    mois_amort = (fin_effective.year - debut_amort.year) * 12 + fin_effective.month - debut_amort.month + 1
                    
                    # Calculer la dotation mensuelle (montant / durée en mois)
                    dotation_mensuelle = float(montant) / (int(duree) * 12)
                    
                    # Ajouter au total des amortissements
                    amortissements_total += dotation_mensuelle * mois_amort
                except Exception:
                    # En cas d'erreur dans le calcul, ignorer cet investissement
                    continue
            
            return amortissements_total
            
        except Exception as e:
            print(f"Erreur dans calculate_amortissements_total_subvention_style: {e}")
            return 0
    
    
    def calculate_amortissement_for_period(self, cursor, project_id, year, month=None):
        """
        Calcule les amortissements pour une période donnée selon la règle :
        - Amortissement mensuel = montant total / (durée amortissement × 12)
        - Début d'amortissement = mois suivant le mois d'achat
        """
        import datetime
        
        try:
            # Récupérer tous les investissements du projet
            cursor.execute('''
                SELECT montant, date_achat, duree 
                FROM investissements 
                WHERE projet_id = ?
            ''', (project_id,))
            
            amortissements_total = 0
            
            for montant_inv, date_achat, duree in cursor.fetchall():
                try:
                    # Convertir la date d'achat (format 'MM/YYYY')
                    achat_date = datetime.datetime.strptime(date_achat, '%m/%Y')
                    
                    # Calculer l'amortissement mensuel
                    amortissement_mensuel = float(montant_inv) / (int(duree) * 12)
                    
                    # Le premier mois d'amortissement = mois suivant l'achat
                    if achat_date.month == 12:
                        premier_mois_amort = datetime.datetime(achat_date.year + 1, 1, 1)
                    else:
                        premier_mois_amort = datetime.datetime(achat_date.year, achat_date.month + 1, 1)
                    
                    # Le dernier mois d'amortissement
                    mois_restants = int(duree) * 12
                    dernier_mois_amort = premier_mois_amort
                    for _ in range(mois_restants - 1):
                        if dernier_mois_amort.month == 12:
                            dernier_mois_amort = datetime.datetime(dernier_mois_amort.year + 1, 1, 1)
                        else:
                            dernier_mois_amort = datetime.datetime(dernier_mois_amort.year, dernier_mois_amort.month + 1, 1)
                    
                    if month:
                        # Calcul pour un mois spécifique
                        mois_demande = datetime.datetime(year, month, 1)
                        
                        # Vérifier si ce mois tombe dans la période d'amortissement
                        if premier_mois_amort <= mois_demande <= dernier_mois_amort:
                            amortissements_total += amortissement_mensuel
                    
                    else:
                        # Calcul pour une année complète
                        debut_annee = datetime.datetime(year, 1, 1)
                        fin_annee = datetime.datetime(year, 12, 1)
                        
                        # Compter les mois d'amortissement qui tombent dans cette année
                        mois_amort_annee = 0
                        mois_courant = premier_mois_amort
                        
                        while mois_courant <= dernier_mois_amort:
                            if debut_annee <= mois_courant <= fin_annee:
                                mois_amort_annee += 1
                            
                            # Passer au mois suivant
                            if mois_courant.month == 12:
                                mois_courant = datetime.datetime(mois_courant.year + 1, 1, 1)
                            else:
                                mois_courant = datetime.datetime(mois_courant.year, mois_courant.month + 1, 1)
                        
                        amortissements_total += amortissement_mensuel * mois_amort_annee
                
                except (ValueError, TypeError) as e:
                    print(f"Erreur lors du traitement de l'investissement {montant_inv}: {e}")
                    continue
            
            return amortissements_total
            
        except Exception as e:
            print(f"Erreur dans calculate_amortissement_for_period: {e}")
            return 0
    
    def calculate_amortissement_for_year(self, cursor, project_id, year, month, projet_info):
        """Calcule les amortissements pour une année donnée"""
        import datetime
        
        try:
            if not projet_info or not projet_info[0] or not projet_info[1]:
                return 0
                
            debut_projet = datetime.datetime.strptime(projet_info[0], '%m/%Y')
            fin_projet = datetime.datetime.strptime(projet_info[1], '%m/%Y')
            
            # Récupérer tous les investissements du projet
            cursor.execute('''
                SELECT montant, date_achat, duree 
                FROM investissements 
                WHERE projet_id = ?
            ''', (project_id,))
            
            amortissements_total = 0
            
            for montant_inv, date_achat, duree in cursor.fetchall():
                try:
                    # Convertir la date d'achat en datetime
                    achat_date = datetime.datetime.strptime(date_achat, '%m/%Y')
                    
                    # La dotation commence le mois suivant l'achat
                    debut_amort = achat_date.replace(day=1)
                    debut_amort = datetime.datetime(debut_amort.year, debut_amort.month, 1) + datetime.timedelta(days=32)
                    debut_amort = debut_amort.replace(day=1)
                    
                    # La fin de l'amortissement
                    fin_amort = datetime.datetime(achat_date.year + int(duree), achat_date.month, 1)
                    fin_effective = min(fin_projet, fin_amort)
                    
                    # Si le début d'amortissement est après la fin du projet, pas d'amortissement
                    if debut_amort > fin_projet:
                        continue
                    
                    # Pour l'année spécifique, calculer la part d'amortissement
                    year_start = datetime.datetime(year, 1, 1)
                    year_end = datetime.datetime(year, 12, 31)
                    
                    # Intersection avec l'année demandée
                    period_start = max(debut_amort, year_start)
                    period_end = min(fin_effective, year_end)
                    
                    if period_start <= period_end:
                        # Calculer le nombre de mois dans cette année
                        if month:
                            # Pour un mois spécifique
                            month_start = datetime.datetime(year, month, 1)
                            month_end = datetime.datetime(year, month, 28)  # Approximation
                            if period_start <= month_end and period_end >= month_start:
                                mois_amort = 1
                            else:
                                mois_amort = 0
                        else:
                            # Pour toute l'année
                            mois_amort = (period_end.year - period_start.year) * 12 + period_end.month - period_start.month + 1
                        
                        # Calculer la dotation mensuelle
                        dotation_mensuelle = float(montant_inv) / (int(duree) * 12)
                        amortissements_total += dotation_mensuelle * mois_amort
                
                except Exception:
                    continue
            
            return amortissements_total
            
        except Exception as e:
            print(f"Erreur dans calculate_amortissement_for_year: {e}")
            return 0

def show_compte_resultat(parent, config_data):
    """Fonction pour afficher le compte de résultat"""
    dialog = CompteResultatDisplay(parent, config_data)
    return dialog.exec()
