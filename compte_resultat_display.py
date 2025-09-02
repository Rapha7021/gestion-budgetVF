import sqlite3
import datetime
from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, 
                            QTableWidgetItem, QPushButton, QMessageBox, 
                            QFileDialog, QHeaderView, QGroupBox, QGridLayout, 
                            QColorDialog, QLineEdit, QComboBox)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
import traceback
import re

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
        
        # Vérifier si au moins un projet a le CIR activé
        self.has_cir_projects = self.check_cir_projects()
        
        self.setWindowTitle("Compte de Résultat")
        self.setMinimumSize(1000, 700)
        
        self.init_ui()
        self.load_data()
    
    def load_export_settings(self):
        """Charge les paramètres d'export depuis la base de données"""
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            
            cursor.execute('SELECT * FROM export_settings WHERE id = 1')
            result = cursor.fetchone()
            
            if result:
                settings = type('Settings', (), {})()
                settings.title_color = result[1]
                settings.header_color = result[2]
                settings.total_color = result[3]
                settings.result_color = result[4]
                settings.logo_path = result[5]
                settings.logo_position = result[6]
            else:
                # Paramètres par défaut
                settings = type('Settings', (), {})()
                settings.title_color = '#2c3e50'
                settings.header_color = '#34495e'
                settings.total_color = '#3498db'
                settings.result_color = '#2ecc71'
                settings.logo_path = ''
                settings.logo_position = 'Haut gauche'
            
            conn.close()
            return settings
            
        except Exception:
            # En cas d'erreur, utiliser les paramètres par défaut
            settings = type('Settings', (), {})()
            settings.title_color = '#2c3e50'
            settings.header_color = '#34495e'
            settings.total_color = '#3498db'
            settings.result_color = '#2ecc71'
            settings.logo_path = ''
            settings.logo_position = 'Haut gauche'
            return settings
    
    def check_cir_projects(self):
        """Vérifie si au moins un projet a le CIR activé"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        try:
            # Récupérer les projets avec CIR activé
            placeholders = ','.join(['?'] * len(self.project_ids))
            cursor.execute(f"""
                SELECT COUNT(*) FROM projets 
                WHERE id IN ({placeholders}) AND cir = 1
            """, self.project_ids)
            
            count = cursor.fetchone()[0] or 0
            return count > 0
            
        except sqlite3.OperationalError:
            # Table n'existe pas ou colonne CIR manquante
            return False
        finally:
            conn.close()
    
    def format_currency(self, value, with_decimals=True):
        """Formate une valeur monétaire avec le formatage français :
        - Séparateur des milliers : espace
        - Séparateur des décimales : virgule
        """
        if value == 0:
            return ""
        
        if with_decimals:
            # Formatage avec 2 décimales
            formatted = f"{value:,.2f}"
        else:
            # Formatage sans décimales (pour les jours par exemple)
            formatted = f"{value:,.0f}"
        
        # Remplacer la virgule (séparateur des milliers) par un espace
        # et le point (séparateur des décimales) par une virgule
        formatted = formatted.replace(",", "TEMP").replace(".", ",").replace("TEMP", " ")
        
        return formatted
    
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
        
        # Paramètres Export
        settings_btn = QPushButton("Paramètres Export")
        settings_btn.clicked.connect(self.open_export_settings)
        settings_btn.setStyleSheet("QPushButton { background-color: #9b59b6; color: white; font-weight: bold; padding: 8px; }")
        buttons_layout.addWidget(settings_btn)
        
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
    
    def open_export_settings(self):
        """Ouvre le dialogue des paramètres d'export"""
        dialog = ExportSettingsDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            QMessageBox.information(self, "Paramètres", "Paramètres d'export sauvegardés avec succès!")
    
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
        # Conditions de filtrage et définition des noms de mois
        month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                      "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        
        if month:
            # Filtrage mensuel - utilise les noms de mois français
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
            'nb_jours_total': 0,
            'cout_moyen_par_jour': 0,
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
                subvention_periode = self.calculate_proportional_distributed_subvention(
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
            
            # Calculer le nombre total de jours
            query_jours = f"""
                SELECT COALESCE(SUM(t.jours), 0)
                FROM temps_travail t
                WHERE t.annee = {year} {month_condition.replace('mois', 't.mois') if month_condition else ''} 
                AND t.projet_id IN ({','.join(['?'] * len(self.project_ids))})
            """
            cursor.execute(query_jours, self.project_ids)
            data['nb_jours_total'] = cursor.fetchone()[0] or 0
            
            # Calculer le coût moyen par jour
            if data['nb_jours_total'] > 0:
                data['cout_moyen_par_jour'] = data['cout_direct'] / data['nb_jours_total']
            else:
                data['cout_moyen_par_jour'] = 0
            
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
            data['cout_direct'] = 0
            data['nb_jours_total'] = 0
            data['cout_moyen_par_jour'] = 0
        
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
        
        # 7. CRÉDIT D'IMPÔT RECHERCHE (CIR) - Calculé sur l'ensemble du projet puis réparti
        try:
            data['credit_impot'] = 0
            data['credit_impot_note'] = ""
            
            if self.has_cir_projects:
                # Calculer le CIR total du projet et le répartir par année
                cir_result = self.calculate_distributed_cir(cursor, year, month)
                
                if cir_result == "CIR_NON_APPLICABLE":
                    data['credit_impot'] = 0
                    data['credit_impot_note'] = "CIR non applicable (subventions > dépenses éligibles)"
                else:
                    data['credit_impot'] = cir_result
                
        except sqlite3.OperationalError as e:
            # Erreur avec les tables - message d'alerte mais continuer
            if not hasattr(self, '_cir_error_shown'):
                QMessageBox.warning(self.parent, "Erreur CIR", 
                                  f"Erreur lors du calcul du CIR : {str(e)}\n"
                                  f"Le calcul du CIR sera ignoré.")
                self._cir_error_shown = True
            data['credit_impot'] = 0
        except Exception as e:
            print(f"DEBUG CIR: Erreur générale = {str(e)}")
            data['credit_impot'] = 0
        
        return data
    
    def get_cost_type_label(self):
        """Retourne le libellé du type de coût sélectionné"""
        cost_type_mapping = {
            'montant_charge': 'Salaire (montant chargé)',
            'cout_production': 'Salaire (coût de production)',
            'cout_complet': 'Salaire (coût complet)'
        }
        return cost_type_mapping.get(self.cost_type, 'Salaire (coût direct)')
    
    def generate_filename(self, extension):
        """Génère un nom de fichier basé sur la configuration"""
        # 1. Partie projets
        if len(self.project_ids) <= 3:
            # Récupérer les codes des projets
            project_codes = []
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            try:
                for project_id in self.project_ids:
                    cursor.execute("SELECT code FROM projets WHERE id = ?", (project_id,))
                    result = cursor.fetchone()
                    if result and result[0]:
                        # Nettoyer le code projet pour le nom de fichier
                        clean_code = re.sub(r'[^\w\-_]', '', result[0])
                        project_codes.append(clean_code)
                    else:
                        project_codes.append(f"PROJ{project_id}")
            finally:
                conn.close()
            
            projects_part = "_".join(project_codes) if project_codes else "projets"
        else:
            projects_part = "multi_projets"
        
        # 2. Partie années
        if len(self.years) == 1:
            years_part = str(self.years[0])
        else:
            years_sorted = sorted(self.years)
            years_part = f"{years_sorted[0]}_{years_sorted[-1]}"
        
        # 3. Partie granularité
        granularity_part = "mensuel" if self.granularity == 'monthly' else "annuel"
        
        # 4. Partie type de coût (version courte pour le nom de fichier)
        cost_type_mapping = {
            'montant_charge': 'montant_charge',
            'cout_production': 'cout_production', 
            'cout_complet': 'cout_complet'
        }
        cost_part = cost_type_mapping.get(self.cost_type, 'cout_production')
        
        # 5. Assembler le nom de fichier
        filename = f"compte_resultat_{projects_part}_{years_part}_{granularity_part}_{cost_part}.{extension}"
        
        # 6. Nettoyer le nom de fichier pour Windows
        # Remplacer les caractères problématiques
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        # Limiter la longueur (255 caractères max pour Windows)
        if len(filename) > 255:
            base_name = filename[:255-len(extension)-1]
            filename = f"{base_name}.{extension}"
        
        return filename

    def calculate_distributed_cir(self, cursor, target_year, target_month=None):
        """
        Calcule le CIR total du projet et le répartit proportionnellement aux dépenses éligibles de la période.
        
        NOUVELLE LOGIQUE DE RÉPARTITION :
        - Le CIR total est calculé sur l'ensemble du projet
        - Il est ensuite réparti proportionnellement aux dépenses éligibles (temps de travail * k1 + amortissements * k2) 
          de la période demandée (mois ou année)
        - Cela permet une répartition mensuelle cohérente basée sur l'activité réelle de chaque période
        
        Retourne:
        - Un nombre négatif si CIR applicable (diminue les charges)
        - 0 si pas de CIR
        - "CIR_NON_APPLICABLE" si subventions > dépenses éligibles
        """
        try:
            # Récupérer les projets CIR
            cir_project_ids = []
            for project_id in self.project_ids:
                cursor.execute("SELECT cir FROM projets WHERE id = ? AND cir = 1", (project_id,))
                if cursor.fetchone():
                    cir_project_ids.append(project_id)
            
            if not cir_project_ids:
                return 0
            
            # 1. Calculer les totaux sur toutes les années du projet pour tous les types de dépenses éligibles
            total_montant_charge = 0
            total_amortissements = 0
            total_subventions = 0
            
            # Calculer les coûts éligibles de la période cible pour la répartition
            # (sera recalculé plus bas avec les bons coefficients)
            
            month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                          "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
            
            for year in self.years:
                # Montant chargé pour cette année
                year_montant_charge = 0
                cir_project_condition = f"AND tt.projet_id IN ({','.join(['?'] * len(cir_project_ids))})"
                
                query = f"""
                    SELECT COALESCE(SUM(tt.jours * cc.montant_charge), 0)
                    FROM temps_travail tt
                    JOIN categorie_cout cc ON tt.categorie = cc.libelle AND tt.annee = cc.annee
                    WHERE tt.annee = {year} {cir_project_condition}
                """
                cursor.execute(query, cir_project_ids)
                year_montant_charge = cursor.fetchone()[0] or 0
                total_montant_charge += year_montant_charge                # Amortissements pour cette année
                year_amort = 0
                for project_id in cir_project_ids:
                    amort = self.calculate_amortissement_for_period(cursor, project_id, year)
                    year_amort += amort
                total_amortissements += year_amort
                
                # Si c'est l'année cible, on traitera les calculs plus tard avec les bons coefficients
                
                # Subventions pour cette année
                year_subvention = 0
                for project_id in cir_project_ids:
                    cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id = ?", (project_id,))
                    projet_info = cursor.fetchone()
                    if projet_info:
                        subv = self.calculate_proportional_distributed_subvention(
                            cursor, project_id, year, None, projet_info)
                        year_subvention += subv
                total_subventions += year_subvention
            
            # 2. Utiliser les coefficients de la première année disponible
            k1, k2, k3 = None, None, None
            for year in sorted(self.years):
                cursor.execute('SELECT k1, k2, k3 FROM cir_coeffs WHERE annee = ?', (year,))
                cir_coeffs = cursor.fetchone()
                if cir_coeffs:
                    k1, k2, k3 = cir_coeffs
                    break
            
            if not k1:
                if not hasattr(self, '_cir_warning_shown'):
                    QMessageBox.warning(self.parent, "Avertissement CIR", 
                                      f"Aucun coefficient CIR trouvé pour les années {self.years}.\n"
                                      f"Le calcul du CIR sera ignoré.")
                    self._cir_warning_shown = True
                return 0
            
            # 3. Calculer le CIR total du projet
            montant_eligible_total = (total_montant_charge * k1) + (total_amortissements * k2)
            montant_net_eligible_total = montant_eligible_total - total_subventions
            cir_total = montant_net_eligible_total * k3
            
            # 4. Vérifier si le CIR est applicable
            if montant_net_eligible_total <= 0:
                # Subventions dépassent les dépenses éligibles
                return "CIR_NON_APPLICABLE"
            
            # 5. Calculer le total des coûts éligibles sur tout le projet (pour la proportion)
            total_eligible_costs = (total_montant_charge * k1) + (total_amortissements * k2)
            
            # 6. Répartir proportionnellement aux dépenses éligibles de la période
            if total_eligible_costs > 0:
                
                # Appliquer les coefficients k1 et k2 aux coûts de la période cible
                # target_period_eligible_costs contient déjà montant_charge + amortissements
                # Il faut séparer pour appliquer les bons coefficients
                
                # Recalculer séparément pour appliquer les coefficients correctement
                period_montant_charge_eligible = 0
                period_amortissements_eligible = 0
                
                if target_month:
                    # Pour un mois spécifique
                    query_month = f"""
                        SELECT COALESCE(SUM(tt.jours * cc.montant_charge), 0)
                        FROM temps_travail tt
                        JOIN categorie_cout cc ON tt.categorie = cc.libelle AND tt.annee = cc.annee
                        WHERE tt.annee = {target_year} AND tt.mois = '{month_names[target_month-1]}' {cir_project_condition}
                    """
                    cursor.execute(query_month, cir_project_ids)
                    period_montant_charge_eligible = cursor.fetchone()[0] or 0
                    
                    for project_id in cir_project_ids:
                        amort = self.calculate_amortissement_for_period(cursor, project_id, target_year, target_month)
                        period_amortissements_eligible += amort
                else:
                    # Pour une année complète
                    query_year = f"""
                        SELECT COALESCE(SUM(tt.jours * cc.montant_charge), 0)
                        FROM temps_travail tt
                        JOIN categorie_cout cc ON tt.categorie = cc.libelle AND tt.annee = cc.annee
                        WHERE tt.annee = {target_year} {cir_project_condition}
                    """
                    cursor.execute(query_year, cir_project_ids)
                    period_montant_charge_eligible = cursor.fetchone()[0] or 0
                    
                    for project_id in cir_project_ids:
                        amort = self.calculate_amortissement_for_period(cursor, project_id, target_year)
                        period_amortissements_eligible += amort
                
                # Appliquer les coefficients aux coûts de la période
                target_period_weighted_costs = (period_montant_charge_eligible * k1) + (period_amortissements_eligible * k2)
                
                # Calculer la proportion basée sur les coûts pondérés
                proportion = (target_period_weighted_costs / total_eligible_costs) if total_eligible_costs > 0 else 0
                
                # Répartir le CIR proportionnellement
                cir_reparti = cir_total * proportion
                
                # Le crédit d'impôt est négatif (diminue les charges)
                # Si le CIR calculé est positif, on le rend négatif pour diminuer les charges
                if cir_reparti > 0:
                    result = -abs(cir_reparti)
                else:
                    result = 0
                    
                return result
            else:
                return 0
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            return 0

    def populate_table(self, data):
        """Remplit le tableau avec les données"""
        # Structure du compte de résultat selon vos spécifications
        structure = [
            ("PRODUITS", "header"),
            ("Recettes", "recettes"),
            ("Subventions", "subventions"),
            ("TOTAL PRODUITS", "total_produits"),
            ("CHARGES", "header"),
            ("Achats et sous-traitance", "achats_sous_traitance"),
            ("Autres achats", "autres_achats"),
            ("Dotation aux amortissements", "dotation_amortissements"),
            (self.get_cost_type_label(), "cout_direct"),  # Nom dynamique selon le type de coût
            ("  → Nombre de jours TOTAL", "nb_jours_total"),
            ("  → Coût moyen par jour", "cout_moyen_par_jour"),
            ("CHARGES EXCEPTIONNELLES", "header"),
        ]
        
        # Ajouter la ligne CIR seulement si au moins un projet a le CIR activé
        if self.has_cir_projects:
            structure.append(("Crédit d'impôt", "credit_impot"))
        
        structure.extend([
            ("TOTAL CHARGES", "total_charges"),
            ("", "separator"),
            ("RÉSULTAT FINANCIER", "resultat_financier")
        ])
        
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
                    item = QTableWidgetItem(self.format_currency(value) if value != 0 else "")
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
                    
                    # Formatage spécial pour les nouveaux indicateurs
                    if data_key == "nb_jours_total":
                        # Afficher le nombre de jours sans décimales
                        item = QTableWidgetItem(f"{self.format_currency(value, False)} jours" if value != 0 else "")
                    elif data_key == "cout_moyen_par_jour":
                        # Afficher le coût moyen par jour avec 2 décimales et €/jour
                        item = QTableWidgetItem(f"{self.format_currency(value)} €/jour" if value != 0 else "")
                    elif data_key == "credit_impot":
                        # Formatage spécial pour le CIR avec gestion de la note explicative
                        note = data[period].get('credit_impot_note', "")
                        if note:
                            item = QTableWidgetItem(note)
                        else:
                            item = QTableWidgetItem(self.format_currency(value) if value != 0 else "")
                    else:
                        # Formatage normal pour les autres données
                        item = QTableWidgetItem(self.format_currency(value) if value != 0 else "")
                    
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    
                    # Style spécial pour les indicateurs (légèrement en retrait visuellement)
                    if data_key in ["nb_jours_total", "cout_moyen_par_jour"]:
                        item.setForeground(QColor(100, 100, 100))  # Couleur légèrement grisée
                    
                    self.table.setItem(row, col, item)
            
            # Colonne TOTAL si plusieurs périodes
            if len(data) > 1 and col < self.table.columnCount() - 1:
                total_value = self.calculate_row_total(data, data_key)
                if total_value is not None:
                    # Formatage spécial pour les nouveaux indicateurs dans la colonne TOTAL
                    if data_key == "nb_jours_total":
                        item = QTableWidgetItem(f"{self.format_currency(total_value, False)} jours" if total_value != 0 else "")
                    elif data_key == "cout_moyen_par_jour":
                        item = QTableWidgetItem(f"{self.format_currency(total_value)} €/jour" if total_value != 0 else "")
                    else:
                        item = QTableWidgetItem(self.format_currency(total_value) if total_value != 0 else "")
                    
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    
                    # Style spécial pour les indicateurs
                    if data_key in ["nb_jours_total", "cout_moyen_par_jour"]:
                        item.setForeground(QColor(100, 100, 100))
                    elif data_key.startswith("total_") or data_key.startswith("resultat_"):
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
            total_charges = (period_data.get('achats_sous_traitance', 0) + 
                           period_data.get('autres_achats', 0) + 
                           period_data.get('cout_direct', 0) + 
                           period_data.get('dotation_amortissements', 0))
            
            # Ajouter le crédit d'impôt seulement si au moins un projet a le CIR activé
            if self.has_cir_projects:
                total_charges += period_data.get('credit_impot', 0)
                
            return total_charges
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
        elif data_key == "cout_moyen_par_jour":
            # Pour le coût moyen par jour, recalculer la moyenne globale
            total_cout = 0
            total_jours = 0
            for period_data in all_data.values():
                total_cout += period_data.get('cout_direct', 0)
                total_jours += period_data.get('nb_jours_total', 0)
            
            return total_cout / total_jours if total_jours > 0 else 0
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
            
            # Charger les paramètres d'export
            settings = self.load_export_settings()
            
            # Générer le nom de fichier basé sur la configuration
            default_filename = self.generate_filename("xlsx")
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Exporter le compte de résultat",
                default_filename,
                "Fichiers Excel (*.xlsx)"
            )
            
            if not file_path:
                return
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Compte de Résultat"
            
            # Convertir les couleurs hex en couleurs openpyxl (sans le #)
            def hex_to_openpyxl(hex_color):
                return hex_color.lstrip('#').upper()
            
            # Déterminer la structure selon la présence du logo
            if settings.logo_path and settings.logo_path.strip():
                # Titre
                title_cell = ws.cell(row=1, column=1)
                title_cell.value = f"COMPTE DE RÉSULTAT"
                title_color = hex_to_openpyxl(settings.title_color)
                title_cell.font = Font(bold=True, size=16, color=title_color)
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=self.table.columnCount())
                
                # Sous-titre avec informations
                subtitle_cell = ws.cell(row=2, column=1)
                subtitle_cell.value = self.get_selection_info()
                subtitle_cell.font = Font(size=10, color="666666")
                ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=self.table.columnCount())
                
                # Essayer d'insérer le logo comme image
                try:
                    import os
                    from openpyxl.drawing.image import Image
                    
                    if os.path.exists(settings.logo_path):
                        logo_img = Image(settings.logo_path)
                        
                        # Redimensionner le logo (max 100px de hauteur)
                        max_height = 100
                        if logo_img.height > max_height:
                            ratio = max_height / logo_img.height
                            logo_img.height = max_height
                            logo_img.width = int(logo_img.width * ratio)
                        
                        # Position selon le paramètre
                        if settings.logo_position == "Haut droite":
                            logo_img.anchor = f"{chr(65 + self.table.columnCount() - 1)}3"  # Dernière colonne
                        elif settings.logo_position == "Haut centre":
                            middle_col = self.table.columnCount() // 2
                            logo_img.anchor = f"{chr(65 + middle_col)}3"
                        else:  # Haut gauche par défaut
                            logo_img.anchor = "A3"
                        
                        ws.add_image(logo_img)
                        
                        # Ajuster la hauteur des lignes pour le logo
                        ws.row_dimensions[3].height = max(60, logo_img.height * 0.75)
                        ws.row_dimensions[4].height = 20
                        
                        header_row = 5
                    else:
                        # Si le fichier n'existe pas, juste noter le nom
                        logo_note_cell = ws.cell(row=3, column=1)
                        logo_name = os.path.basename(settings.logo_path)
                        logo_note_cell.value = f"Logo configuré: {logo_name} (Position: {settings.logo_position})"
                        logo_note_cell.font = Font(italic=True, size=9, color="999999")
                        header_row = 4
                        
                except ImportError:
                    # Si openpyxl.drawing.image n'est pas disponible
                    logo_note_cell = ws.cell(row=3, column=1)
                    logo_note_cell.value = f"Logo: {os.path.basename(settings.logo_path)} (Position: {settings.logo_position})"
                    logo_note_cell.font = Font(italic=True, size=9, color="999999")
                    header_row = 4
                except Exception:
                    # Autre erreur avec le logo
                    header_row = 3
            else:
                header_row = 1
            
            # Exporter les en-têtes de colonnes à la bonne ligne
            for col in range(self.table.columnCount()):
                header_item = self.table.horizontalHeaderItem(col)
                cell = ws.cell(row=header_row, column=col+1)
                cell.value = header_item.text() if header_item else ""
                cell.font = Font(bold=True, color="FFFFFF")
                header_color = hex_to_openpyxl(settings.header_color)
                cell.fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
                cell.alignment = Alignment(horizontal='center')
            
            # Exporter les données du tableau
            data_start_row = header_row + 1
            for row in range(self.table.rowCount()):
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    cell = ws.cell(row=data_start_row + row, column=col+1)
                    
                    if item and item.text():
                        # Pour la première colonne (postes), garder le texte
                        if col == 0:
                            cell.value = item.text()
                            # Appliquer le style selon le type de ligne
                            if "PRODUITS" in item.text() or "CHARGES" in item.text():
                                cell.font = Font(bold=True, color="FFFFFF")
                                header_color = hex_to_openpyxl(settings.header_color)
                                cell.fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
                            elif "TOTAL" in item.text():
                                cell.font = Font(bold=True, color="FFFFFF")
                                total_color = hex_to_openpyxl(settings.total_color)
                                cell.fill = PatternFill(start_color=total_color, end_color=total_color, fill_type="solid")
                            elif "RÉSULTAT" in item.text():
                                cell.font = Font(bold=True, color="FFFFFF")
                                result_color = hex_to_openpyxl(settings.result_color)
                                cell.fill = PatternFill(start_color=result_color, end_color=result_color, fill_type="solid")
                        else:
                            # Pour les colonnes de données, convertir en nombre si possible
                            try:
                                # Enlever les espaces et virgules pour la conversion
                                # Le texte contient maintenant des espaces comme séparateurs de milliers et des virgules comme décimales
                                value_str = item.text().replace(" ", "").replace(",", ".")
                                if value_str and value_str not in ["jours", "€/jour"]:
                                    # Enlever les unités si présentes
                                    value_str = value_str.replace(" jours", "").replace(" €/jour", "")
                                    cell.value = float(value_str)
                                    # Format français : espace pour milliers, virgule pour décimales
                                    cell.number_format = '# ##0,00'
                                else:
                                    cell.value = ""
                            except ValueError:
                                cell.value = item.text()
                            
                            cell.alignment = Alignment(horizontal='right')
                            
                            # Appliquer les couleurs aux colonnes de données selon le type de ligne
                            first_col_item = self.table.item(row, 0)
                            if first_col_item:
                                first_col_text = first_col_item.text()
                                if "PRODUITS" in first_col_text or "CHARGES" in first_col_text:
                                    cell.font = Font(bold=True, color="FFFFFF")
                                    header_color = hex_to_openpyxl(settings.header_color)
                                    cell.fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
                                elif "TOTAL" in first_col_text:
                                    cell.font = Font(bold=True, color="FFFFFF")
                                    total_color = hex_to_openpyxl(settings.total_color)
                                    cell.fill = PatternFill(start_color=total_color, end_color=total_color, fill_type="solid")
                                elif "RÉSULTAT" in first_col_text:
                                    cell.font = Font(bold=True, color="FFFFFF")
                                    result_color = hex_to_openpyxl(settings.result_color)
                                    cell.fill = PatternFill(start_color=result_color, end_color=result_color, fill_type="solid")
            
            # Ajuster la largeur des colonnes
            try:
                for col_num in range(1, self.table.columnCount() + 1):
                    max_length = 0
                    for row_num in range(1, ws.max_row + 1):
                        try:
                            cell = ws.cell(row=row_num, column=col_num)
                            if cell.value and not hasattr(cell, '_merge_parent'):  # Éviter les cellules fusionnées
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                        except:
                            continue
                    
                    # Calculer la lettre de colonne
                    if col_num <= 26:
                        column_letter = chr(64 + col_num)  # A, B, C, etc.
                    else:
                        column_letter = chr(64 + (col_num - 1) // 26) + chr(65 + (col_num - 1) % 26)
                    
                    adjusted_width = min(max(max_length + 2, 10), 50)
                    ws.column_dimensions[column_letter].width = adjusted_width
            except Exception as e:
                # Si l'ajustement automatique échoue, utiliser des largeurs fixes
                for i in range(self.table.columnCount()):
                    column_letter = chr(65 + i)  # A, B, C, etc.
                    ws.column_dimensions[column_letter].width = 20
            
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
            
            # Générer le nom de fichier basé sur la configuration
            default_filename = self.generate_filename("pdf")
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Exporter le compte de résultat",
                default_filename,
                "Fichiers PDF (*.pdf)"
            )
            
            if not file_path:
                return
            
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(file_path)
            
            # Ne pas définir de marges personnalisées pour éviter les problèmes de compatibilité
            
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
        # Charger les paramètres d'export
        settings = self.load_export_settings()
        
        # Gérer le logo
        logo_html = ""
        if settings.logo_path and settings.logo_path.strip():
            import os
            import base64
            
            try:
                if os.path.exists(settings.logo_path):
                    with open(settings.logo_path, 'rb') as image_file:
                        encoded_string = base64.b64encode(image_file.read()).decode()
                    
                    # Déterminer le type MIME
                    ext = os.path.splitext(settings.logo_path)[1].lower()
                    mime_type = {
                        '.png': 'image/png',
                        '.jpg': 'image/jpeg',
                        '.jpeg': 'image/jpeg',
                        '.gif': 'image/gif',
                        '.bmp': 'image/bmp'
                    }.get(ext, 'image/png')
                    
                    # Position du logo
                    position_style = {
                        'Haut gauche': 'float: left;',
                        'Haut droite': 'float: right;',
                        'Haut centre': 'display: block; margin: 0 auto;'
                    }.get(settings.logo_position, 'float: left;')
                    
                    logo_html = f"""
                    <div style="margin-bottom: 20px;">
                        <img src="data:{mime_type};base64,{encoded_string}" 
                             style="max-height: 80px; max-width: 200px; {position_style}" 
                             alt="Logo">
                    </div>
                    """
            except Exception:
                # Si le logo ne peut pas être chargé, on continue sans
                pass
        
        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                h1 {{ text-align: center; color: {settings.title_color}; }}
                h2 {{ text-align: center; color: #7f8c8d; font-size: 12pt; }}
                table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
                th, td {{ border: 1px solid #bdc3c7; padding: 8px; text-align: left; }}
                th {{ background-color: {settings.header_color}; color: white; font-weight: bold; }}
                .header {{ background-color: {settings.header_color}; color: white; font-weight: bold; }}
                .total {{ background-color: {settings.total_color}; color: white; font-weight: bold; }}
                .result {{ background-color: {settings.result_color}; color: white; font-weight: bold; }}
                .amount {{ text-align: right; }}
                .logo-container {{ overflow: auto; margin-bottom: 20px; }}
            </style>
        </head>
        <body>
            {logo_html}
            <div style="clear: both;"></div>
            <h1>COMPTE DE RÉSULTAT</h1>
            <h2>{self.get_selection_info()}</h2>
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
                
                # Déterminer la classe CSS selon le contenu de la première colonne
                css_class = ""
                if col == 0:  # Première colonne (libellés)
                    if "PRODUITS" in value or "CHARGES" in value:
                        css_class = "header"
                    elif "TOTAL" in value:
                        css_class = "total"
                    elif "RÉSULTAT" in value:
                        css_class = "result"
                else:  # Colonnes de données
                    first_col_item = self.table.item(row, 0)
                    if first_col_item:
                        first_col_text = first_col_item.text()
                        if "PRODUITS" in first_col_text or "CHARGES" in first_col_text:
                            css_class = "header"
                        elif "TOTAL" in first_col_text:
                            css_class = "total"
                        elif "RÉSULTAT" in first_col_text:
                            css_class = "result"
                        else:
                            css_class = "amount"
                    else:
                        css_class = "amount"
                
                html += f'<td class="{css_class}">{value}</td>'
            html += "</tr>"
        
        html += """
            </table>
        </body>
        </html>
        """
        
        return html
    
    def calculate_proportional_distributed_subvention(self, cursor, project_id, year, month, projet_info):
        """
        Calcule la subvention répartie proportionnellement aux dépenses éligibles de la période.
        
        NOUVELLE LOGIQUE DE RÉPARTITION :
        - La subvention totale est calculée sur l'ensemble du projet selon les paramètres configurés
        - Elle est ensuite répartie proportionnellement aux dépenses éligibles (selon les coefficients) 
          de la période demandée (mois ou année)
        - Cela permet une répartition mensuelle cohérente basée sur l'activité réelle de chaque période
        """
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
            
            # Calculer les dates du projet
            debut_projet = datetime.datetime.strptime(projet_info[0], '%m/%Y')
            fin_projet = datetime.datetime.strptime(projet_info[1], '%m/%Y')
            
            # Vérifier si l'année demandée est dans la période du projet
            if year < debut_projet.year or year > fin_projet.year:
                return 0
            
            month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                          "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
            
            subvention_total_projet = 0
            
            for subvention in subventions_config:
                (dep_temps, coef_temps, dep_ext, coef_ext, dep_autres, coef_autres, 
                 dep_amort, coef_amort, cd, taux) = subvention
                
                # 1. Calculer le montant total de subvention sur tout le projet
                montant_subvention_config = 0
                
                # 1.1 TEMPS DE TRAVAIL total du projet
                total_depenses_eligible_temps = 0
                if dep_temps and coef_temps:
                    cout_total_temps = self.calculate_temps_travail_total(cursor, project_id)
                    temps_travail = cout_total_temps * cd
                    total_depenses_eligible_temps = temps_travail
                    montant_subvention_config += coef_temps * temps_travail
                
                # 1.2 DÉPENSES EXTERNES totales du projet
                total_depenses_eligible_externes = 0
                if dep_ext and coef_ext:
                    cursor.execute('''
                        SELECT COALESCE(SUM(montant), 0) FROM depenses 
                        WHERE projet_id = ?
                    ''', (project_id,))
                    depenses_ext_total = cursor.fetchone()[0] or 0
                    total_depenses_eligible_externes = depenses_ext_total
                    montant_subvention_config += coef_ext * depenses_ext_total
                
                # 1.3 AUTRES ACHATS totaux du projet
                total_depenses_eligible_autres = 0
                if dep_autres and coef_autres:
                    cursor.execute('''
                        SELECT COALESCE(SUM(montant), 0) FROM autres_depenses 
                        WHERE projet_id = ?
                    ''', (project_id,))
                    autres_achats_total = cursor.fetchone()[0] or 0
                    total_depenses_eligible_autres = autres_achats_total
                    montant_subvention_config += coef_autres * autres_achats_total
                
                # 1.4 AMORTISSEMENTS totaux du projet
                total_depenses_eligible_amortissements = 0
                if dep_amort and coef_amort:
                    amortissements_total = self.calculate_amortissements_total_subvention_style(
                        cursor, project_id, projet_info)
                    total_depenses_eligible_amortissements = amortissements_total
                    montant_subvention_config += coef_amort * amortissements_total
                
                # Appliquer le taux de subvention
                montant_subvention_config = montant_subvention_config * (taux / 100)
                
                # 2. Calculer les dépenses éligibles de la période cible
                period_depenses_eligible_temps = 0
                period_depenses_eligible_externes = 0
                period_depenses_eligible_autres = 0
                period_depenses_eligible_amortissements = 0
                
                # 2.1 TEMPS DE TRAVAIL de la période
                if dep_temps and coef_temps:
                    if month:
                        # Pour un mois spécifique
                        cursor.execute("""
                            SELECT tt.annee, tt.categorie, tt.mois, tt.jours 
                            FROM temps_travail tt 
                            WHERE tt.projet_id = ? AND tt.annee = ? AND tt.mois = ?
                        """, (project_id, year, month_names[month-1]))
                    else:
                        # Pour une année complète
                        cursor.execute("""
                            SELECT tt.annee, tt.categorie, tt.mois, tt.jours 
                            FROM temps_travail tt 
                            WHERE tt.projet_id = ? AND tt.annee = ?
                        """, (project_id, year))
                    
                    temps_travail_period_rows = cursor.fetchall()
                    cout_period_temps = 0
                    
                    for annee, categorie, mois, jours in temps_travail_period_rows:
                        # Convertir la catégorie comme dans la méthode originale
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
                        
                        if categorie_code:
                            cursor.execute("""
                                SELECT montant_charge 
                                FROM categorie_cout 
                                WHERE categorie = ? AND annee = ?
                            """, (categorie_code, annee))
                            
                            cout_row = cursor.fetchone()
                            if cout_row and cout_row[0]:
                                montant_charge = float(cout_row[0])
                                cout_period_temps += jours * montant_charge
                            else:
                                cout_period_temps += jours * 500  # valeur par défaut
                    
                    period_depenses_eligible_temps = cout_period_temps * cd
                
                # 2.2 DÉPENSES EXTERNES de la période
                if dep_ext and coef_ext:
                    if month:
                        cursor.execute('''
                            SELECT COALESCE(SUM(montant), 0) FROM depenses 
                            WHERE projet_id = ? AND annee = ? AND mois = ?
                        ''', (project_id, year, month_names[month-1]))
                    else:
                        cursor.execute('''
                            SELECT COALESCE(SUM(montant), 0) FROM depenses 
                            WHERE projet_id = ? AND annee = ?
                        ''', (project_id, year))
                    
                    period_depenses_eligible_externes = cursor.fetchone()[0] or 0
                
                # 2.3 AUTRES ACHATS de la période
                if dep_autres and coef_autres:
                    if month:
                        cursor.execute('''
                            SELECT COALESCE(SUM(montant), 0) FROM autres_depenses 
                            WHERE projet_id = ? AND annee = ? AND mois = ?
                        ''', (project_id, year, month_names[month-1]))
                    else:
                        cursor.execute('''
                            SELECT COALESCE(SUM(montant), 0) FROM autres_depenses 
                            WHERE projet_id = ? AND annee = ?
                        ''', (project_id, year))
                    
                    period_depenses_eligible_autres = cursor.fetchone()[0] or 0
                
                # 2.4 AMORTISSEMENTS de la période
                if dep_amort and coef_amort:
                    period_depenses_eligible_amortissements = self.calculate_amortissement_for_period(
                        cursor, project_id, year, month)
                
                # 3. Calculer les totaux pondérés selon les coefficients
                total_depenses_ponderees = (
                    (total_depenses_eligible_temps * coef_temps if dep_temps else 0) +
                    (total_depenses_eligible_externes * coef_ext if dep_ext else 0) +
                    (total_depenses_eligible_autres * coef_autres if dep_autres else 0) +
                    (total_depenses_eligible_amortissements * coef_amort if dep_amort else 0)
                )
                
                period_depenses_ponderees = (
                    (period_depenses_eligible_temps * coef_temps if dep_temps else 0) +
                    (period_depenses_eligible_externes * coef_ext if dep_ext else 0) +
                    (period_depenses_eligible_autres * coef_autres if dep_autres else 0) +
                    (period_depenses_eligible_amortissements * coef_amort if dep_amort else 0)
                )
                
                # 4. Répartir proportionnellement
                if total_depenses_ponderees > 0:
                    proportion = period_depenses_ponderees / total_depenses_ponderees
                    subvention_period = montant_subvention_config * proportion
                    subvention_total_projet += subvention_period
                else:
                    # Si pas de dépenses, répartition équitable comme avant
                    nb_mois_total = (fin_projet.year - debut_projet.year) * 12 + (fin_projet.month - debut_projet.month) + 1
                    if month:
                        subvention_total_projet += montant_subvention_config / nb_mois_total
                    else:
                        start_month = debut_projet.month if year == debut_projet.year else 1
                        end_month = fin_projet.month if year == fin_projet.year else 12
                        nb_mois_annee = end_month - start_month + 1
                        subvention_total_projet += (montant_subvention_config / nb_mois_total) * nb_mois_annee
            
            return subvention_total_projet
            
        except Exception as e:
            print(f"Erreur dans calculate_proportional_distributed_subvention: {e}")
            # En cas d'erreur, fallback vers la méthode simple
            return self.calculate_simple_distributed_subvention(cursor, project_id, year, month, projet_info)
    
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
                    continue
            
            return amortissements_total
            
        except Exception as e:
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
            return 0

def show_compte_resultat(parent, config_data):
    """Fonction pour afficher le compte de résultat"""
    dialog = CompteResultatDisplay(parent, config_data)
    return dialog.exec()


class ExportSettingsDialog(QDialog):
    """Dialogue pour configurer les paramètres d'export PDF et Excel"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("Paramètres d'Export")
        self.setMinimumSize(500, 600)
        
        # Charger les paramètres existants
        self.settings = self.load_export_settings()
        
        self.init_ui()
        self.load_current_settings()
        
    def init_ui(self):
        """Initialise l'interface utilisateur"""
        layout = QVBoxLayout()
        
        # Titre
        title = QLabel("Configuration des Exports PDF et Excel")
        title.setStyleSheet("font-size: 14pt; font-weight: bold; color: #2c3e50; margin-bottom: 10px;")
        layout.addWidget(title)
        
        # Section couleurs
        colors_group = self.create_colors_section()
        layout.addWidget(colors_group)
        
        # Section logo
        logo_group = self.create_logo_section()
        layout.addWidget(logo_group)
        
        # Aperçu des couleurs
        preview_group = self.create_preview_section()
        layout.addWidget(preview_group)
        
        # Boutons
        buttons_layout = self.create_buttons()
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
    
    def create_colors_section(self):
        """Crée la section de configuration des couleurs"""
        group = QGroupBox("Configuration des Couleurs")
        layout = QGridLayout()
        
        # Couleur du titre du document
        layout.addWidget(QLabel("Couleur du titre du document:"), 0, 0)
        self.title_color_btn = QPushButton()
        self.title_color_btn.setFixedSize(80, 30)
        self.title_color_btn.clicked.connect(lambda: self.select_color('title'))
        layout.addWidget(self.title_color_btn, 0, 1)
        
        # Couleur des en-têtes
        layout.addWidget(QLabel("Couleur des en-têtes:"), 1, 0)
        self.header_color_btn = QPushButton()
        self.header_color_btn.setFixedSize(80, 30)
        self.header_color_btn.clicked.connect(lambda: self.select_color('header'))
        layout.addWidget(self.header_color_btn, 1, 1)
        
        # Couleur des totaux
        layout.addWidget(QLabel("Couleur des totaux:"), 2, 0)
        self.total_color_btn = QPushButton()
        self.total_color_btn.setFixedSize(80, 30)
        self.total_color_btn.clicked.connect(lambda: self.select_color('total'))
        layout.addWidget(self.total_color_btn, 2, 1)
        
        # Couleur du résultat financier
        layout.addWidget(QLabel("Couleur du résultat financier:"), 3, 0)
        self.result_color_btn = QPushButton()
        self.result_color_btn.setFixedSize(80, 30)
        self.result_color_btn.clicked.connect(lambda: self.select_color('result'))
        layout.addWidget(self.result_color_btn, 3, 1)
        
        group.setLayout(layout)
        return group
    
    def create_logo_section(self):
        """Crée la section de configuration du logo"""
        group = QGroupBox("Configuration du Logo")
        layout = QGridLayout()
        
        # Chemin du logo
        layout.addWidget(QLabel("Fichier logo:"), 0, 0)
        self.logo_path = QLineEdit()
        self.logo_path.setPlaceholderText("Sélectionnez un fichier image...")
        layout.addWidget(self.logo_path, 0, 1)
        
        browse_btn = QPushButton("Parcourir")
        browse_btn.clicked.connect(self.browse_logo)
        layout.addWidget(browse_btn, 0, 2)
        
        # Position du logo
        layout.addWidget(QLabel("Position du logo:"), 1, 0)
        self.logo_position = QComboBox()
        self.logo_position.addItems(["Haut gauche", "Haut droite", "Haut centre"])
        layout.addWidget(self.logo_position, 1, 1)
        
        group.setLayout(layout)
        return group
    
    def create_preview_section(self):
        """Crée la section d'aperçu des couleurs"""
        group = QGroupBox("Aperçu des Couleurs")
        layout = QVBoxLayout()
        
        # Tableau d'aperçu
        self.preview_table = QTableWidget(6, 2)
        self.preview_table.setHorizontalHeaderLabels(["Élément", "Valeur"])
        self.preview_table.setMaximumHeight(200)
        
        # Données d'exemple
        preview_data = [
            ("COMPTE DE RÉSULTAT", "title"),
            ("Poste", "header"),
            ("Recettes", "100 000,00"),
            ("Subventions", "50 000,00"),
            ("TOTAL PRODUITS", "total"),
            ("RÉSULTAT FINANCIER", "result")
        ]
        
        for row, (label, data_type) in enumerate(preview_data):
            # Colonne label
            item_label = QTableWidgetItem(label)
            if data_type == "title":
                item_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
            elif data_type == "header":
                item_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            elif data_type in ["total", "result"]:
                item_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
            
            self.preview_table.setItem(row, 0, item_label)
            
            # Colonne valeur
            if data_type in ["title", "header", "total", "result"]:
                item_value = QTableWidgetItem("")
            else:
                item_value = QTableWidgetItem(data_type)
                item_value.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            
            self.preview_table.setItem(row, 1, item_value)
        
        self.preview_table.resizeColumnsToContents()
        layout.addWidget(self.preview_table)
        
        group.setLayout(layout)
        return group
    
    def create_buttons(self):
        """Crée les boutons de validation"""
        layout = QHBoxLayout()
        
        # Bouton Réinitialiser
        reset_btn = QPushButton("Réinitialiser")
        reset_btn.clicked.connect(self.reset_to_defaults)
        layout.addWidget(reset_btn)
        
        layout.addStretch()
        
        # Bouton Annuler
        cancel_btn = QPushButton("Annuler")
        cancel_btn.clicked.connect(self.reject)
        layout.addWidget(cancel_btn)
        
        # Bouton Sauvegarder
        save_btn = QPushButton("Sauvegarder")
        save_btn.clicked.connect(self.save_settings)
        save_btn.setStyleSheet("QPushButton { background-color: #27ae60; color: white; font-weight: bold; padding: 8px; }")
        layout.addWidget(save_btn)
        
        return layout
    
    def select_color(self, color_type):
        """Ouvre le sélecteur de couleur"""
        current_color = getattr(self.settings, f'{color_type}_color', '#000000')
        color = QColorDialog.getColor(QColor(current_color), self, f"Choisir la couleur - {color_type}")
        
        if color.isValid():
            setattr(self.settings, f'{color_type}_color', color.name())
            self.update_color_button(color_type, color.name())
            self.update_preview()
    
    def update_color_button(self, color_type, color_hex):
        """Met à jour l'apparence du bouton de couleur"""
        button = getattr(self, f'{color_type}_color_btn')
        button.setStyleSheet(f"background-color: {color_hex}; border: 1px solid #999999;")
    
    def browse_logo(self):
        """Ouvre le dialogue de sélection de fichier pour le logo"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Sélectionner un logo",
            "",
            "Images (*.png *.jpg *.jpeg *.bmp *.gif)"
        )
        
        if file_path:
            self.logo_path.setText(file_path)
            self.settings.logo_path = file_path
    
    def update_preview(self):
        """Met à jour l'aperçu des couleurs"""
        color_mapping = {
            0: ('title_color', '#2c3e50'),  # Titre
            1: ('header_color', '#34495e'),  # En-tête
            4: ('total_color', '#3498db'),   # Total
            5: ('result_color', '#2ecc71')   # Résultat
        }
        
        for row, (color_attr, default_color) in color_mapping.items():
            color_hex = getattr(self.settings, color_attr, default_color)
            
            # Mettre à jour la couleur de fond des cellules
            for col in range(2):
                item = self.preview_table.item(row, col)
                if item:
                    item.setBackground(QColor(color_hex))
                    item.setForeground(QColor('#ffffff'))
    
    def load_current_settings(self):
        """Charge les paramètres actuels dans l'interface"""
        # Mettre à jour les boutons de couleur
        colors = ['title', 'header', 'total', 'result']
        defaults = ['#2c3e50', '#34495e', '#3498db', '#2ecc71']
        
        for color_type, default in zip(colors, defaults):
            color_hex = getattr(self.settings, f'{color_type}_color', default)
            self.update_color_button(color_type, color_hex)
        
        # Mettre à jour le logo
        if hasattr(self.settings, 'logo_path') and self.settings.logo_path:
            self.logo_path.setText(self.settings.logo_path)
        
        if hasattr(self.settings, 'logo_position') and self.settings.logo_position:
            index = self.logo_position.findText(self.settings.logo_position)
            if index >= 0:
                self.logo_position.setCurrentIndex(index)
        
        # Mettre à jour l'aperçu
        self.update_preview()
    
    def reset_to_defaults(self):
        """Remet les paramètres par défaut"""
        self.settings = self.get_default_settings()
        self.load_current_settings()
    
    def save_settings(self):
        """Sauvegarde les paramètres dans la base de données"""
        try:
            # Récupérer les valeurs actuelles
            self.settings.logo_path = self.logo_path.text()
            self.settings.logo_position = self.logo_position.currentText()
            
            # Sauvegarder dans la base de données
            self.save_export_settings(self.settings)
            
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"Erreur lors de la sauvegarde: {str(e)}")
    
    def load_export_settings(self):
        """Charge les paramètres d'export depuis la base de données"""
        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            
            # Créer la table si elle n'existe pas
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS export_settings (
                    id INTEGER PRIMARY KEY,
                    title_color TEXT DEFAULT '#2c3e50',
                    header_color TEXT DEFAULT '#34495e',
                    total_color TEXT DEFAULT '#3498db',
                    result_color TEXT DEFAULT '#2ecc71',
                    logo_path TEXT DEFAULT '',
                    logo_position TEXT DEFAULT 'Haut gauche'
                )
            ''')
            
            # Récupérer les paramètres
            cursor.execute('SELECT * FROM export_settings WHERE id = 1')
            result = cursor.fetchone()
            
            if result:
                settings = type('Settings', (), {})()
                settings.title_color = result[1]
                settings.header_color = result[2]
                settings.total_color = result[3]
                settings.result_color = result[4]
                settings.logo_path = result[5]
                settings.logo_position = result[6]
            else:
                settings = self.get_default_settings()
            
            conn.close()
            return settings
            
        except Exception as e:
            QMessageBox.warning(self, "Avertissement", f"Impossible de charger les paramètres: {str(e)}")
            return self.get_default_settings()
    
    def get_default_settings(self):
        """Retourne les paramètres par défaut"""
        settings = type('Settings', (), {})()
        settings.title_color = '#2c3e50'
        settings.header_color = '#34495e'
        settings.total_color = '#3498db'
        settings.result_color = '#2ecc71'
        settings.logo_path = ''
        settings.logo_position = 'Haut gauche'
        return settings
    
    def save_export_settings(self, settings):
        """Sauvegarde les paramètres dans la base de données"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        try:
            # Créer la table si elle n'existe pas
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS export_settings (
                    id INTEGER PRIMARY KEY,
                    title_color TEXT DEFAULT '#2c3e50',
                    header_color TEXT DEFAULT '#34495e',
                    total_color TEXT DEFAULT '#3498db',
                    result_color TEXT DEFAULT '#2ecc71',
                    logo_path TEXT DEFAULT '',
                    logo_position TEXT DEFAULT 'Haut gauche'
                )
            ''')
            
            # Insérer ou mettre à jour
            cursor.execute('''
                INSERT OR REPLACE INTO export_settings 
                (id, title_color, header_color, total_color, result_color, logo_path, logo_position)
                VALUES (1, ?, ?, ?, ?, ?, ?)
            ''', (settings.title_color, settings.header_color, settings.total_color,
                  settings.result_color, settings.logo_path, settings.logo_position))
            
            conn.commit()
            
        finally:
            conn.close()
