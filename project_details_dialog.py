from PyQt6.QtWidgets import QDialog, QVBoxLayout, QGridLayout, QLabel, QHBoxLayout, QListWidget, QListWidgetItem, QPushButton, QInputDialog, QMessageBox, QTextEdit, QDialogButtonBox
from PyQt6.QtCore import Qt
import sqlite3
import datetime
from utils import format_montant, format_montant_aligne
DB_PATH = 'gestion_budget.db'

class ProjectDetailsDialog(QDialog):
    def __init__(self, parent, projet_id):
        super().__init__(parent)
        self.setWindowTitle('Détails du projet')
        screen = self.screen().geometry()
        self.resize(int(screen.width() * 0.9), int(screen.height() * 0.9))
        self.show()
        main_layout = QVBoxLayout()
        grid = QGridLayout()
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT p.code, p.nom, p.details, p.date_debut, p.date_fin, p.livrables, 
                       c.nom || ' ' || c.prenom || ' - ' || c.direction AS chef_complet, 
                       p.etat, p.cir, p.subvention, p.theme_principal
                FROM projets p
                LEFT JOIN chefs_projet c ON p.chef = c.id
                WHERE p.id=?
            ''', (projet_id,))
            projet = cursor.fetchone()
            
            # Thèmes
            cursor.execute('SELECT t.nom FROM themes t JOIN projet_themes pt ON t.id=pt.theme_id WHERE pt.projet_id=?', (projet_id,))
            themes = [nom for (nom,) in cursor.fetchall()]
            # Investissements
            cursor.execute('SELECT montant, date_achat, duree FROM investissements WHERE projet_id=?', (projet_id,))
            investissements = cursor.fetchall()
            # Equipe
            cursor.execute('SELECT direction, type, nombre FROM equipe WHERE projet_id=?', (projet_id,))
            equipe = cursor.fetchall()
            # Images
            cursor.execute('SELECT nom, data FROM images WHERE projet_id=?', (projet_id,))
            images = cursor.fetchall()
            
            # Calcul des coûts
            cursor.execute('''
                SELECT t.categorie, SUM(t.jours) AS total_jours
                FROM temps_travail t
                WHERE t.projet_id = ?
                GROUP BY t.categorie
            ''', (projet_id,))
            categories_jours = cursor.fetchall()
            
            # Mapping entre les libellés complets et les codes courts
            mapping_categories = {
                "Stagiaire Projet": "STP",
                "Assistante / opérateur": "AOP", 
                "Technicien": "TEP",
                "Junior": "IJP",
                "Senior": "ISP",
                "Expert": "EDP",
                "Collaborateur moyen": "MOY"
            }
            
            couts = {"charge": 0, "direct": 0, "complet": 0}
            for categorie, total_jours in categories_jours:
                # Convertir le libellé en code court
                code_categorie = mapping_categories.get(categorie, categorie)
                
                cursor.execute('''
                    SELECT montant_charge, cout_production, cout_complet
                    FROM categorie_cout
                    WHERE categorie = ?
                ''', (code_categorie,))
                res = cursor.fetchone()
                
                if res:
                    montant_charge, cout_production, cout_complet = res
                    couts["charge"] += (montant_charge or 0) * total_jours
                    couts["direct"] += (cout_production or 0) * total_jours
                    couts["complet"] += (cout_complet or 0) * total_jours

            # Ajouter les dépenses externes
            cursor.execute('''
                SELECT SUM(montant) 
                FROM depenses 
                WHERE projet_id = ?
            ''', (projet_id,))
            depenses_externes = cursor.fetchone()
            if depenses_externes and depenses_externes[0]:
                montant_depenses = float(depenses_externes[0])
                # Les dépenses externes s'ajoutent à tous les types de coûts
                couts["charge"] += montant_depenses
                couts["direct"] += montant_depenses
                couts["complet"] += montant_depenses

            # Ajouter les autres dépenses
            cursor.execute('''
                SELECT SUM(montant) 
                FROM autres_depenses 
                WHERE projet_id = ?
            ''', (projet_id,))
            autres_depenses = cursor.fetchone()
            if autres_depenses and autres_depenses[0]:
                montant_autres = float(autres_depenses[0])
                # Les autres dépenses s'ajoutent à tous les types de coûts
                couts["charge"] += montant_autres
                couts["direct"] += montant_autres
                couts["complet"] += montant_autres
                    
        # Affichage des coûts
        self.budget_vbox = QVBoxLayout()
        self.budget_vbox.addWidget(QLabel(f"<b>Budget Total :</b>"))
        
        # Créer les labels avec police monospace pour l'alignement
        cout_charge_label = QLabel(f"Coût chargé     : {format_montant_aligne(couts['charge'])}")
        cout_charge_label.setStyleSheet("font-family: 'Courier New', monospace;")
        self.budget_vbox.addWidget(cout_charge_label)
        
        cout_production_label = QLabel(f"Coût production : {format_montant_aligne(couts['direct'])}")
        cout_production_label.setStyleSheet("font-family: 'Courier New', monospace;")
        self.budget_vbox.addWidget(cout_production_label)
        
        cout_complet_label = QLabel(f"Coût complet    : {format_montant_aligne(couts['complet'])}")
        cout_complet_label.setStyleSheet("font-family: 'Courier New', monospace;")
        self.budget_vbox.addWidget(cout_complet_label)
        
        grid.addLayout(self.budget_vbox, 0, 2)
        
        # En haut à gauche
        left_vbox = QVBoxLayout()
        left_vbox.addWidget(QLabel(f"<b>Code projet :</b> {projet[0]}"))
        h_nom = QHBoxLayout()
        h_nom.addWidget(QLabel(f"<b>Nom projet :</b> {projet[1]}"))
        left_vbox.addLayout(h_nom)
        
        # Champ détails avec retour à la ligne automatique
        details_label = QLabel(f"<b>Détails :</b> {projet[2]}")
        details_label.setWordWrap(True)  # Permet le retour à la ligne automatique
        details_label.setMaximumWidth(400)  # Limite la largeur pour forcer les retours à la ligne
        left_vbox.addWidget(details_label)
        
        grid.addLayout(left_vbox, 0, 0)
        # En haut à droite
        right_vbox = QVBoxLayout()
        right_vbox.addWidget(QLabel(f"<b>Date début :</b> {projet[3]}"))
        right_vbox.addWidget(QLabel(f"<b>Date fin :</b> {projet[4]}"))
        right_vbox.addWidget(QLabel(f"<b>Livrables :</b> {projet[5]}"))
        right_vbox.addWidget(QLabel(f"<b>Chef(fe) de projet :</b> {projet[6]}"))
        if projet[10]:
            right_vbox.addWidget(QLabel(f"<b>Thème principal :</b> {projet[10]}"))
        if themes:
            right_vbox.addWidget(QLabel("<b>Thèmes :</b> " + ", ".join(themes)))
        grid.addLayout(right_vbox, 0, 1)
        # Centre haut
        center_vbox = QVBoxLayout()
        center_vbox.addWidget(QLabel(f"<b>Etat :</b> {projet[7]}"))
        center_vbox.addWidget(QLabel(f"<b>CIR :</b> {'Oui' if projet[8] else 'Non'}"))
        center_vbox.addWidget(QLabel(f"<b>Subvention :</b> {'Oui' if projet[9] else 'Non'}"))
        # Investissements
        invest_text = "<b>Investissements :</b>\n"
        if investissements:
            for montant, date_achat, duree in investissements:
                invest_text += f"- {format_montant(montant)} | Achat: {date_achat} | Durée: {duree} ans\n"
        else:
            invest_text += "Aucun"
        center_vbox.addWidget(QLabel(invest_text))
        # Equipe
        equipe_text = "<b>Equipe :</b>\n"
        if equipe:
            # Organiser les données d'équipe par direction
            equipe_par_direction = {}
            for direction, type_, nombre in equipe:
                if nombre > 0:  # Ne considérer que les membres avec un nombre > 0
                    if direction not in equipe_par_direction:
                        equipe_par_direction[direction] = []
                    equipe_par_direction[direction].append((type_, nombre))
            
            # Afficher les équipes par direction
            if equipe_par_direction:
                for direction, membres in equipe_par_direction.items():
                    if membres:  # Ne pas afficher les directions sans membres
                        equipe_text += f"<b>{direction}</b>:\n"
                        for type_, nombre in membres:
                            equipe_text += f"  - {type_}: {nombre}\n"
            else:
                equipe_text += "Aucune info"
        else:
            equipe_text += "Aucune info"
        center_vbox.addWidget(QLabel(equipe_text))
        grid.addLayout(center_vbox, 1, 0, 1, 2)
        # Images en dessous
        img_label = QLabel("<b>Images du projet :</b>")
        main_layout.addWidget(img_label)
        img_hbox = QHBoxLayout()
        for nom, data in images:
            try:
                from PyQt6.QtGui import QPixmap
                pixmap = QPixmap()
                if pixmap.loadFromData(data):
                    img_widget = QLabel()
                    img_widget.setPixmap(pixmap.scaled(150, 150, Qt.AspectRatioMode.KeepAspectRatio))
                    img_widget.setStyleSheet("border: 1px solid gray; margin: 2px;")
                    img_widget.setToolTip(nom)  # Afficher le nom de l'image au survol
                    img_hbox.addWidget(img_widget)
            except Exception as e:
                # En cas d'erreur, afficher un placeholder
                error_widget = QLabel(f"Erreur image:\n{nom}")
                error_widget.setFixedSize(150, 150)
                error_widget.setStyleSheet("border: 1px solid red; background-color: #ffeeee; color: red;")
                error_widget.setAlignment(Qt.AlignmentFlag.AlignCenter)
                img_hbox.addWidget(error_widget)
        main_layout.addLayout(grid)
        main_layout.addLayout(img_hbox)

        # Section actualités
        actualites_label = QLabel("<b>Actualités du projet :</b>")
        main_layout.addWidget(actualites_label)
        self.actualites_list = QListWidget()
        self.actualites_list.setMaximumHeight(100)  # Hauteur max fixée, ajustable selon besoin
        main_layout.addWidget(self.actualites_list)

        # Boutons actualités
        btn_hbox = QHBoxLayout()
        add_btn = QPushButton("Ajouter une actualité")
        edit_btn = QPushButton("Modifier l'actualité sélectionnée")
        del_btn = QPushButton("Supprimer l'actualité sélectionnée")
        btn_hbox.addWidget(add_btn)
        btn_hbox.addWidget(edit_btn)
        btn_hbox.addWidget(del_btn)
        # Bouton import Excel
        import_excel_btn = QPushButton("Importer Excel")
        btn_hbox.addWidget(import_excel_btn)
        # Nouveau bouton "Modifier le budget"
        edit_budget_btn = QPushButton("Modifier le budget")
        btn_hbox.addWidget(edit_budget_btn)
        # Ajout du bouton "Gérer les tâches"
        manage_tasks_btn = QPushButton("Gérer les tâches")
        btn_hbox.addWidget(manage_tasks_btn)
        # Ajout du bouton "Compte de résultat global"
        compte_resultat_btn = QPushButton("Compte de résultat global")
        btn_hbox.addWidget(compte_resultat_btn)
        main_layout.addLayout(btn_hbox)

        self.projet_id = projet_id
        self.load_actualites()

        add_btn.clicked.connect(self.add_actualite)
        edit_btn.clicked.connect(self.edit_actualite)
        del_btn.clicked.connect(self.delete_actualite)
        # Import Excel
        self.df_long = None
        import_excel_btn.clicked.connect(self.handle_import_excel)
        # Connexion du bouton "Modifier le budget"
        edit_budget_btn.clicked.connect(self.edit_budget)
        # Connexion du bouton "Gérer les tâches"
        manage_tasks_btn.clicked.connect(self.open_task_manager)
        # Connexion du bouton "Compte de résultat global"
        compte_resultat_btn.clicked.connect(self.open_compte_resultat)

        # Espace vide en dessous
        main_layout.addStretch()
        self.setLayout(main_layout)

        self.refresh_budget()  # Recalcul initial des coûts

    def handle_import_excel(self):
        """Ouvre l'interface d'import Excel configurable avec le projet pré-sélectionné"""
        try:
            from excel_import import open_excel_import_dialog
            open_excel_import_dialog(self, self.projet_id)
            # Actualiser les coûts et subventions après import Excel
            self.refresh_budget()
        except ImportError as e:
            QMessageBox.critical(
                self, "Erreur", 
                f"Impossible de charger le module d'import Excel:\n{str(e)}\n\n"
                "Assurez-vous que pandas est installé et que le fichier excel_import.py est présent."
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Erreur", 
                f"Erreur lors de l'ouverture de l'import Excel:\n{str(e)}"
            )

    def load_actualites(self):
        self.actualites_list.clear()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT id, message, date FROM actualites WHERE projet_id=? ORDER BY date DESC', (self.projet_id,))
        for id_, msg, date in cursor.fetchall():
            # Création d'un widget personnalisé pour le message
            widget = QLabel(f"[{date}] {msg}")
            widget.setWordWrap(True)
            item = QListWidgetItem()
            item.setSizeHint(widget.sizeHint())
            self.actualites_list.addItem(item)
            self.actualites_list.setItemWidget(item, widget)
            item.setData(Qt.ItemDataRole.UserRole, id_)
        conn.close()

    def get_multiline_text(self, title, label, text=""):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(label))
        edit = QTextEdit()
        edit.setPlainText(text)
        edit.setMinimumSize(400, 150)  # Taille minimum confortable
        layout.addWidget(edit)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        layout.addWidget(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        result = dialog.exec()
        return edit.toPlainText(), result == QDialog.DialogCode.Accepted

    def add_actualite(self):
        text, ok = self.get_multiline_text("Ajouter une actualité", "Message :")
        if ok and text.strip():
            from datetime import datetime
            date = datetime.now().strftime("%Y-%m-%d %H:%M")
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('INSERT INTO actualites (projet_id, message, date) VALUES (?, ?, ?)', (self.projet_id, text, date))
            conn.commit()
            conn.close()
            self.load_actualites()

    def edit_actualite(self):
        item = self.actualites_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Modifier", "Sélectionnez une actualité à modifier.")
            return
        id_ = item.data(Qt.ItemDataRole.UserRole)
        widget = self.actualites_list.itemWidget(item)
        if widget:
            full_text = widget.text()
            old_msg = full_text.split('] ', 1)[-1]
        else:
            old_msg = ""
        text, ok = self.get_multiline_text("Modifier l'actualité", "Message :", old_msg)
        if ok and text.strip():
            from datetime import datetime
            date = datetime.now().strftime("%Y-%m-%d %H:%M")
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('UPDATE actualites SET message=?, date=? WHERE id=?', (text, date, id_))
            conn.commit()
            conn.close()
            self.load_actualites()

    def delete_actualite(self):
        item = self.actualites_list.currentItem()
        if not item:
            QMessageBox.warning(self, "Supprimer", "Sélectionnez une actualité à supprimer.")
            return
        id_ = item.data(Qt.ItemDataRole.UserRole)
        confirm = QMessageBox.question(self, "Confirmation", "Voulez-vous vraiment supprimer cette actualité ?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM actualites WHERE id=?', (id_,))
            conn.commit()
            conn.close()
            self.load_actualites()

    def edit_budget(self):
        from budget_edit_dialog import BudgetEditDialog
        dlg = BudgetEditDialog(self.projet_id, self)
        dlg.exec()
        # Actualiser les coûts et subventions après modification du budget
        self.refresh_budget()

    def open_task_manager(self):
        from task_manager_dialog import TaskManagerDialog
        dlg = TaskManagerDialog(self, self.projet_id)
        dlg.exec()
        # Actualiser les coûts et subventions après gestion des tâches
        self.refresh_budget()

    def open_compte_resultat(self):
        """Ouvre le compte de résultat pour ce projet spécifique"""
        try:
            # Récupérer les dates du projet pour déterminer les années
            import sqlite3
            import datetime
            from compte_resultat_display import show_compte_resultat
            
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('SELECT date_debut, date_fin FROM projets WHERE id = ?', (self.projet_id,))
            date_row = cursor.fetchone()
            conn.close()
            
            if not date_row or not date_row[0] or not date_row[1]:
                QMessageBox.warning(self, "Erreur", "Les dates de début et fin du projet ne sont pas définies.")
                return
            
            try:
                # Convertir les dates MM/yyyy en objets datetime
                debut_projet = datetime.datetime.strptime(date_row[0], '%m/%Y')
                fin_projet = datetime.datetime.strptime(date_row[1], '%m/%Y')
                
                # Générer la liste des années du projet
                years = list(range(debut_projet.year, fin_projet.year + 1))
                
                # Configuration pour le compte de résultat
                config_data = {
                    'project_ids': [self.projet_id],  # Seulement ce projet
                    'years': years,  # Toutes les années du projet
                    'period_type': 'yearly',  # Type de période annuel
                    'granularity': 'yearly',  # Granularité annuelle
                    'cost_type': 'montant_charge'  # Coûts chargés
                }
                
                # Ouvrir le compte de résultat
                show_compte_resultat(self, config_data)
                
            except ValueError as e:
                QMessageBox.warning(self, "Erreur", f"Format de date invalide dans le projet : {str(e)}")
                
        except ImportError as e:
            QMessageBox.critical(
                self, "Erreur", 
                f"Impossible de charger le module de compte de résultat :\n{str(e)}\n\n"
                "Assurez-vous que le fichier compte_resultat_display.py est présent."
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Erreur", 
                f"Erreur lors de l'ouverture du compte de résultat :\n{str(e)}"
            )

    def get_project_data_for_subventions(self):
        """Récupère les données du projet pour calculer les subventions (similaire à SubventionDialog)"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        data = {
            'temps_travail_total': 0,
            'depenses_externes': 0,
            'autres_achats': 0,
            'amortissements': 0
        }
        
        # 1. Récupérer dates de début et fin du projet
        cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (self.projet_id,))
        date_row = cursor.fetchone()
        if not date_row or not date_row[0] or not date_row[1]:
            conn.close()
            return data
        
        import datetime
        
        # Convertir les dates MM/yyyy en objets datetime
        try:
            debut_projet = datetime.datetime.strptime(date_row[0], '%m/%Y')
            fin_projet = datetime.datetime.strptime(date_row[1], '%m/%Y')
        except ValueError:
            conn.close()
            return data
        
        # 2. Calculer le temps de travail et le montant chargé
        cursor.execute("""
            SELECT tt.annee, tt.categorie, tt.mois, tt.jours 
            FROM temps_travail tt 
            WHERE tt.projet_id = ?
        """, (self.projet_id,))
        
        temps_travail_rows = cursor.fetchall()
        cout_total_temps = 0
        
        mapping_categories = {
            "Stagiaire Projet": "STP",
            "Assistante / opérateur": "AOP", 
            "Technicien": "TEP",
            "Junior": "IJP",
            "Senior": "ISP",
            "Expert": "EDP",
            "Collaborateur moyen": "MOY"
        }
        
        for annee, categorie, mois, jours in temps_travail_rows:
            # Convertir la catégorie du temps de travail au format de categorie_cout
            categorie_code = mapping_categories.get(categorie, "")
            
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
        
        data['temps_travail_total'] = cout_total_temps
        
        # 3. Récupérer toutes les dépenses externes
        cursor.execute("""
            SELECT SUM(montant) 
            FROM depenses 
            WHERE projet_id = ?
        """, (self.projet_id,))
        
        depenses_row = cursor.fetchone()
        if depenses_row and depenses_row[0]:
            data['depenses_externes'] = float(depenses_row[0])
        
        # 4. Récupérer toutes les autres dépenses
        cursor.execute("""
            SELECT SUM(montant) 
            FROM autres_depenses 
            WHERE projet_id = ?
        """, (self.projet_id,))
        
        autres_depenses_row = cursor.fetchone()
        if autres_depenses_row and autres_depenses_row[0]:
            data['autres_achats'] = float(autres_depenses_row[0])
        
        # 5. Calculer les dotations aux amortissements
        cursor.execute("""
            SELECT montant, date_achat, duree 
            FROM investissements 
            WHERE projet_id = ?
        """, (self.projet_id,))
        
        amortissements_total = 0
        
        for montant, date_achat, duree in cursor.fetchall():
            try:
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
        
        data['amortissements'] = amortissements_total
        
        conn.close()
        return data

    def has_cir_activated(self):
        """Vérifie si le projet a le CIR activé"""
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            try:
                cursor.execute('SELECT cir FROM projets WHERE id = ?', (self.projet_id,))
                result = cursor.fetchone()
                return result and result[0] == 1
            except sqlite3.OperationalError:
                return False

    def refresh_cir(self, total_subventions):
        """Calcule et affiche le montant du CIR"""
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            
            # Récupérer les coefficients CIR pour les années du projet
            cursor.execute('SELECT date_debut, date_fin FROM projets WHERE id = ?', (self.projet_id,))
            date_row = cursor.fetchone()
            
            if not date_row or not date_row[0] or not date_row[1]:
                return
            
            try:
                import datetime
                debut_projet = datetime.datetime.strptime(date_row[0], '%m/%Y')
                fin_projet = datetime.datetime.strptime(date_row[1], '%m/%Y')
            except ValueError:
                return

            # Récupérer les coefficients CIR (utiliser la première année disponible)
            k1, k2, k3 = None, None, None
            for year in range(debut_projet.year, fin_projet.year + 1):
                cursor.execute('SELECT k1, k2, k3 FROM cir_coeffs WHERE annee = ?', (year,))
                cir_coeffs = cursor.fetchone()
                if cir_coeffs and all(coeff is not None for coeff in cir_coeffs):
                    k1, k2, k3 = cir_coeffs
                    break

            if not all(coeff is not None for coeff in [k1, k2, k3]):
                # Pas de coefficients CIR trouvés
                return

            # Récupérer les données du projet
            projet_data = self.get_project_data_for_subventions()

            # Calculer le montant éligible
            temps_travail_eligible = projet_data['temps_travail_total'] * k1
            amortissements_eligible = projet_data['amortissements'] * k2
            montant_eligible = temps_travail_eligible + amortissements_eligible

            # Soustraire les subventions
            montant_net_eligible = montant_eligible - total_subventions

            # Ajouter un séparateur
            self.budget_vbox.addWidget(QLabel(""))
            
            # Vérifier si le CIR est applicable
            if montant_net_eligible > 0:
                # CIR applicable - afficher l'assiette éligible avec le taux K3
                taux_k3_percent = k3 * 100  # Convertir en pourcentage
                cir_label = QLabel(f"Assiette éligible \"CIR\" : {format_montant(montant_net_eligible)} (taux : {taux_k3_percent:.0f} %)")
                self.budget_vbox.addWidget(cir_label)
            else:
                # CIR non applicable - afficher le message explicatif
                cir_label = QLabel("CIR non applicable (subventions > dépenses éligibles)")
                cir_label.setStyleSheet("color: #e74c3c; font-style: italic;")  # Rouge et italique pour bien voir
                self.budget_vbox.addWidget(cir_label)

    def refresh_budget(self):
        """Recalcule et met à jour les coûts du budget."""
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            # Recalcul des coûts
            cursor.execute('''
                SELECT t.categorie, SUM(t.jours) AS total_jours
                FROM temps_travail t
                WHERE t.projet_id = ?
                GROUP BY t.categorie
            ''', (self.projet_id,))
            categories_jours = cursor.fetchall()

            mapping_categories = {
                "Stagiaire Projet": "STP",
                "Assistante / opérateur": "AOP", 
                "Technicien": "TEP",
                "Junior": "IJP",
                "Senior": "ISP",
                "Expert": "EDP",
                "Collaborateur moyen": "MOY"
            }

            couts = {"charge": 0, "direct": 0, "complet": 0}
            missing_data = False
            for categorie, total_jours in categories_jours:
                code_categorie = mapping_categories.get(categorie, categorie)
                cursor.execute('''
                    SELECT montant_charge, cout_production, cout_complet
                    FROM categorie_cout
                    WHERE categorie = ?
                ''', (code_categorie,))
                res = cursor.fetchone()

                if res:
                    montant_charge, cout_production, cout_complet = res
                    couts["charge"] += (montant_charge or 0) * total_jours
                    couts["direct"] += (cout_production or 0) * total_jours
                    couts["complet"] += (cout_complet or 0) * total_jours
                else:
                    missing_data = True

            # Ajouter les dépenses externes
            cursor.execute('''
                SELECT SUM(montant) 
                FROM depenses 
                WHERE projet_id = ?
            ''', (self.projet_id,))
            depenses_externes = cursor.fetchone()
            if depenses_externes and depenses_externes[0]:
                montant_depenses = float(depenses_externes[0])
                # Les dépenses externes s'ajoutent à tous les types de coûts
                couts["charge"] += montant_depenses
                couts["direct"] += montant_depenses
                couts["complet"] += montant_depenses

            # Ajouter les autres dépenses
            cursor.execute('''
                SELECT SUM(montant) 
                FROM autres_depenses 
                WHERE projet_id = ?
            ''', (self.projet_id,))
            autres_depenses = cursor.fetchone()
            if autres_depenses and autres_depenses[0]:
                montant_autres = float(autres_depenses[0])
                # Les autres dépenses s'ajoutent à tous les types de coûts
                couts["charge"] += montant_autres
                couts["direct"] += montant_autres
                couts["complet"] += montant_autres

            # Mise à jour des labels avec alignement
            cout_charge_label = self.budget_vbox.itemAt(1).widget()
            cout_charge_label.setText(f"Coût chargé     : {format_montant_aligne(couts['charge'])}")
            cout_charge_label.setStyleSheet("font-family: 'Courier New', monospace;")
            
            cout_production_label = self.budget_vbox.itemAt(2).widget()
            cout_production_label.setText(f"Coût production : {format_montant_aligne(couts['direct'])}")
            cout_production_label.setStyleSheet("font-family: 'Courier New', monospace;")
            
            cout_complet_label = self.budget_vbox.itemAt(3).widget()
            cout_complet_label.setText(f"Coût complet    : {format_montant_aligne(couts['complet'])}")
            cout_complet_label.setStyleSheet("font-family: 'Courier New', monospace;")

            if missing_data:
                self.budget_vbox.addWidget(QLabel("<i>Note : Certaines données sont manquantes pour le calcul.</i>"))

            # Calcul et affichage des subventions
            self.refresh_subventions()

    def refresh_subventions(self):
        """Calcule et affiche les montants des subventions"""
        # Supprimer les anciens labels de subventions (s'ils existent)
        while self.budget_vbox.count() > 4:  # Garder seulement les 4 premiers items (titre + 3 coûts)
            item = self.budget_vbox.takeAt(self.budget_vbox.count() - 1)
            if item.widget():
                item.widget().deleteLater()

        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            
            # Récupérer toutes les subventions pour ce projet
            try:
                cursor.execute('''
                    SELECT nom, depenses_temps_travail, coef_temps_travail, 
                           depenses_externes, coef_externes, depenses_autres_achats, coef_autres_achats,
                           depenses_dotation_amortissements, coef_dotation_amortissements, 
                           cd, taux, montant_subvention_max, depenses_eligibles_max, mode_simplifie, montant_forfaitaire
                    FROM subventions 
                    WHERE projet_id = ?
                ''', (self.projet_id,))
                subventions = cursor.fetchall()
            except sqlite3.OperationalError:
                # Les colonnes mode_simplifie et montant_forfaitaire n'existent peut-être pas encore
                try:
                    cursor.execute('''
                        SELECT nom, depenses_temps_travail, coef_temps_travail, 
                               depenses_externes, coef_externes, depenses_autres_achats, coef_autres_achats,
                               depenses_dotation_amortissements, coef_dotation_amortissements, 
                               cd, taux, montant_subvention_max, depenses_eligibles_max
                        FROM subventions 
                        WHERE projet_id = ?
                    ''', (self.projet_id,))
                    subventions = [list(row) + [0, 0] for row in cursor.fetchall()]  # Ajouter 0 pour mode_simplifie et montant_forfaitaire
                except sqlite3.OperationalError:
                    try:
                        cursor.execute('''
                            SELECT nom, depenses_temps_travail, coef_temps_travail, 
                                   depenses_externes, coef_externes, depenses_autres_achats, coef_autres_achats,
                                   depenses_dotation_amortissements, coef_dotation_amortissements, 
                                   cd, taux
                            FROM subventions 
                            WHERE projet_id = ?
                        ''', (self.projet_id,))
                        subventions = [list(row) + [0, 0, 0, 0] for row in cursor.fetchall()]  # Ajouter 0 pour montant_max, depenses_max, mode_simplifie et montant_forfaitaire
                    except sqlite3.OperationalError:
                        subventions = []

            if not subventions:
                return

            # Récupérer les données du projet
            projet_data = self.get_project_data_for_subventions()

            # Ajouter un séparateur
            self.budget_vbox.addWidget(QLabel(""))
            self.budget_vbox.addWidget(QLabel("<b>Subventions :</b>"))

            total_subventions = 0

            for subv in subventions:
                if len(subv) >= 15:
                    nom, depenses_temps, coef_temps, depenses_ext, coef_ext, depenses_autres, coef_autres, depenses_amort, coef_amort, cd, taux, montant_max, depenses_max, mode_simplifie, montant_forfaitaire = subv
                elif len(subv) >= 13:
                    nom, depenses_temps, coef_temps, depenses_ext, coef_ext, depenses_autres, coef_autres, depenses_amort, coef_amort, cd, taux, montant_max, depenses_max = subv[:13]
                    mode_simplifie = subv[13] if len(subv) > 13 else 0
                    montant_forfaitaire = subv[14] if len(subv) > 14 else 0
                else:
                    nom, depenses_temps, coef_temps, depenses_ext, coef_ext, depenses_autres, coef_autres, depenses_amort, coef_amort, cd, taux = subv[:11]
                    montant_max = subv[11] if len(subv) > 11 else 0
                    depenses_max = subv[12] if len(subv) > 12 else 0
                    mode_simplifie = subv[13] if len(subv) > 13 else 0
                    montant_forfaitaire = subv[14] if len(subv) > 14 else 0

                # Vérifier si c'est une subvention en mode simplifié
                if mode_simplifie:
                    # Mode simplifié : calculer l'assiette éligible totale et le taux
                    assiette_totale = (projet_data['temps_travail_total'] + 
                                     projet_data['depenses_externes'] + 
                                     projet_data['autres_achats'] + 
                                     projet_data['amortissements'])
                    
                    # Calculer le taux automatiquement
                    if assiette_totale > 0:
                        taux_calcule = (montant_forfaitaire / assiette_totale) * 100
                    else:
                        taux_calcule = 0
                    
                    # Afficher l'assiette éligible totale avec le taux calculé
                    subv_label = QLabel(f"Assiette éligible \"{nom}\" : {format_montant(assiette_totale)} (taux : {taux_calcule:.2f} %)")
                    self.budget_vbox.addWidget(subv_label)
                    
                    # Ajouter le montant forfaitaire au total
                    total_subventions += montant_forfaitaire
                    
                else:
                    # Mode détaillé : calcul avec les coefficients
                    assiette_eligible = 0

                    # Temps de travail
                    if depenses_temps:
                        temps_travail = projet_data['temps_travail_total'] * (cd or 1)
                        assiette_eligible += (coef_temps or 0) * temps_travail

                    # Dépenses externes
                    if depenses_ext:
                        assiette_eligible += (coef_ext or 0) * projet_data['depenses_externes']

                    # Autres achats
                    if depenses_autres:
                        assiette_eligible += (coef_autres or 0) * projet_data['autres_achats']

                    # Dotation amortissements
                    if depenses_amort:
                        assiette_eligible += (coef_amort or 0) * projet_data['amortissements']

                    # Appliquer le plafond sur l'assiette éligible
                    if depenses_max and depenses_max > 0:
                        assiette_eligible = min(assiette_eligible, depenses_max)

                    # Calculer le montant de la subvention (pour le total)
                    montant = assiette_eligible * ((taux or 0) / 100)

                    # Appliquer le plafond sur le montant final si défini
                    if montant_max and montant_max > 0:
                        montant = min(montant, montant_max)

                    total_subventions += montant

                    # Afficher l'assiette éligible (plafonnée) avec le taux
                    subv_label = QLabel(f"Assiette éligible \"{nom}\" : {format_montant(assiette_eligible)} (taux : {taux:.0f} %)")
                    self.budget_vbox.addWidget(subv_label)

            # Calculer et afficher le CIR si le projet l'a activé
            if self.has_cir_activated():
                # Ajouter un séparateur entre subventions et CIR
                self.budget_vbox.addWidget(QLabel("─" * 40))  # Trait de séparation
                self.refresh_cir(total_subventions)

