from PyQt6.QtWidgets import QDialog, QVBoxLayout, QGridLayout, QLabel, QHBoxLayout, QListWidget, QListWidgetItem, QPushButton, QInputDialog, QMessageBox, QTextEdit, QDialogButtonBox, QTableWidget, QTableWidgetItem, QHeaderView
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
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
                    
        # Réorganisation de la mise en page avec une structure plus claire
        # Partie haute avec informations principales
        top_section = QHBoxLayout()
        
        # Colonne gauche - Informations projet
        left_column = QVBoxLayout()
        left_column.addWidget(QLabel(f"<b>Code projet :</b> {projet[0]}"))
        left_column.addWidget(QLabel(f"<b>Nom projet :</b> {projet[1]}"))
        
        # Champ détails avec retour à la ligne automatique
        details_label = QLabel(f"<b>Détails :</b> {projet[2]}")
        details_label.setWordWrap(True)
        details_label.setMaximumWidth(350)
        left_column.addWidget(details_label)
        
        left_column.addWidget(QLabel(f"<b>Date début :</b> {projet[3]}"))
        left_column.addWidget(QLabel(f"<b>Date fin :</b> {projet[4]}"))
        left_column.addWidget(QLabel(f"<b>Livrables :</b> {projet[5]}"))
        left_column.addWidget(QLabel(f"<b>Chef(fe) de projet :</b> {projet[6]}"))
        if projet[10]:
            left_column.addWidget(QLabel(f"<b>Thème principal :</b> {projet[10]}"))
        if themes:
            left_column.addWidget(QLabel("<b>Thèmes :</b> " + ", ".join(themes)))
        
        # Ajouter la section Équipe dans la colonne gauche (en dessous des livrables)
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
        equipe_label = QLabel(equipe_text)
        equipe_label.setWordWrap(True)  # Activer le retour à la ligne automatique
        equipe_label.setMaximumWidth(450)  # Élargir la largeur maximum
        left_column.addWidget(equipe_label)
        
        # Colonne centrale - État, CIR, Subvention, Investissements (plus serrés)
        center_column = QVBoxLayout()
        center_column.setSpacing(5)  # Réduire l'espacement entre les éléments
        
        center_column.addWidget(QLabel(f"<b>Etat :</b> {projet[7]}"))
        center_column.addWidget(QLabel(f"<b>CIR :</b> {'Oui' if projet[8] else 'Non'}"))
        center_column.addWidget(QLabel(f"<b>Subvention :</b> {'Oui' if projet[9] else 'Non'}"))
        
        # Investissements
        invest_text = "<b>Investissements :</b>\n"
        if investissements:
            for montant, date_achat, duree in investissements:
                invest_text += f"- {format_montant(montant)} | Achat: {date_achat} | Durée: {duree} ans\n"
        else:
            invest_text += "Aucun"
        invest_label = QLabel(invest_text)
        invest_label.setMaximumWidth(250)
        center_column.addWidget(invest_label)
        
        # Colonne droite - Budget et subventions
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
        
        # Ajouter les colonnes au layout horizontal
        top_section.addLayout(left_column)
        top_section.addLayout(center_column)
        top_section.addLayout(self.budget_vbox)
        
        # Ajouter la section haute au layout principal
        main_layout.addLayout(top_section)
        # Section images
        img_label = QLabel("<b>Images du projet :</b>")
        main_layout.addWidget(img_label)
        img_hbox = QHBoxLayout()
        for nom, data in images:
            try:
                from PyQt6.QtGui import QPixmap
                pixmap = QPixmap()
                if pixmap.loadFromData(data):
                    img_widget = QLabel()
                    # Augmenter la taille d'affichage et améliorer la qualité
                    scaled_pixmap = pixmap.scaled(300, 300, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                    img_widget.setPixmap(scaled_pixmap)
                    img_widget.setStyleSheet("border: 2px solid gray; margin: 5px; background-color: white;")
                    img_widget.setToolTip(f"{nom}\nTaille originale: {pixmap.width()}x{pixmap.height()}")  # Afficher le nom et la taille originale
                    # Permettre le clic pour voir en taille réelle
                    img_widget.mousePressEvent = lambda event, p=pixmap, n=nom: self.show_fullsize_image(p, n)
                    img_widget.setCursor(Qt.CursorShape.PointingHandCursor)
                    img_hbox.addWidget(img_widget)
            except Exception as e:
                # En cas d'erreur, afficher un placeholder
                error_widget = QLabel(f"Erreur image:\n{nom}")
                error_widget.setFixedSize(300, 300)
                error_widget.setStyleSheet("border: 2px solid red; background-color: #ffeeee; color: red;")
                error_widget.setAlignment(Qt.AlignmentFlag.AlignCenter)
                img_hbox.addWidget(error_widget)
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
        # Nouveau bouton "Modifier le projet"
        edit_project_btn = QPushButton("Modifier le projet")
        btn_hbox.addWidget(edit_project_btn)
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
        # Connexion du bouton "Modifier le projet"
        edit_project_btn.clicked.connect(self.edit_project)
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

    def show_fullsize_image(self, pixmap, nom):
        """Affiche une image en taille réelle dans une nouvelle fenêtre"""
        from PyQt6.QtWidgets import QScrollArea
        
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Image: {nom}")
        dialog.setModal(True)
        
        # Définir une taille raisonnable pour la fenêtre
        screen = dialog.screen().geometry()
        max_width = int(screen.width() * 0.8)
        max_height = int(screen.height() * 0.8)
        
        # Créer un QScrollArea pour permettre le défilement si l'image est très grande
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        # Créer le label pour l'image
        image_label = QLabel()
        image_label.setPixmap(pixmap)
        image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        image_label.setStyleSheet("background-color: white;")
        
        # Ajouter le label au scroll area
        scroll_area.setWidget(image_label)
        
        # Layout de la dialog
        layout = QVBoxLayout(dialog)
        layout.addWidget(scroll_area)
        
        # Ajuster la taille de la fenêtre
        img_width = pixmap.width()
        img_height = pixmap.height()
        
        # Calculer la taille optimale de la fenêtre
        window_width = min(img_width + 50, max_width)
        window_height = min(img_height + 80, max_height)  # +80 pour la barre de titre
        
        dialog.resize(window_width, window_height)
        
        # Centrer la fenêtre
        dialog.move(
            (screen.width() - window_width) // 2,
            (screen.height() - window_height) // 2
        )
        
        dialog.exec()

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

    def edit_project(self):
        """Ouvre le formulaire de modification du projet"""
        try:
            # Importer la classe ProjectForm depuis main.py
            from main import ProjectForm
            
            # Créer et ouvrir le formulaire de modification
            form = ProjectForm(self, self.projet_id)
            if form.exec():
                # Si le projet a été modifié, rafraîchir les données de cette page
                QMessageBox.information(
                    self, 
                    "Projet modifié", 
                    "Le projet a été modifié avec succès.\n"
                    "Les données vont être actualisées."
                )
                self.refresh_project_data()  # Rafraîchir les données
        except ImportError as e:
            QMessageBox.critical(
                self, "Erreur", 
                f"Impossible de charger le formulaire de projet :\n{str(e)}\n\n"
                "Assurez-vous que le fichier main.py est présent."
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Erreur", 
                f"Erreur lors de l'ouverture du formulaire de projet :\n{str(e)}"
            )

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

    def get_project_data_for_subventions(self):
        """Récupère les données du projet pour calculer les subventions (version simplifiée pour le CIR)"""
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
        
        # 2. Calculer le temps de travail et le montant chargé - FILTRÉ PAR DATES DU PROJET
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
        
        month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                      "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        
        for annee, categorie, mois, jours in temps_travail_rows:
            # NOUVEAU: Vérifier que cette entrée est dans la période du projet
            try:
                # Convertir le mois français en numéro
                mois_num = month_names.index(mois) + 1 if mois in month_names else 1
                entry_date = datetime.datetime(annee, mois_num, 1)
                
                # Vérifier si cette entrée est dans la période du projet
                if not (debut_projet <= entry_date <= fin_projet):
                    continue  # Ignorer cette entrée si hors période
            except (ValueError, IndexError):
                # Si erreur de conversion, ignorer cette entrée
                continue
            
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
        
        # 3. Récupérer les dépenses externes - FILTRÉ PAR DATES DU PROJET
        cursor.execute("""
            SELECT d.annee, d.mois, d.montant
            FROM depenses d
            WHERE d.projet_id = ?
        """, (self.projet_id,))
        
        depenses_rows = cursor.fetchall()
        total_depenses_externes = 0
        
        for annee, mois, montant in depenses_rows:
            # Vérifier que cette dépense est dans la période du projet
            try:
                mois_num = month_names.index(mois) + 1 if mois in month_names else 1
                entry_date = datetime.datetime(annee, mois_num, 1)
                
                if debut_projet <= entry_date <= fin_projet:
                    total_depenses_externes += float(montant)
            except (ValueError, IndexError):
                continue
        
        data['depenses_externes'] = total_depenses_externes
        
        # 4. Récupérer les autres dépenses - FILTRÉ PAR DATES DU PROJET
        cursor.execute("""
            SELECT ad.annee, ad.mois, ad.montant
            FROM autres_depenses ad
            WHERE ad.projet_id = ?
        """, (self.projet_id,))
        
        autres_depenses_rows = cursor.fetchall()
        total_autres_achats = 0
        
        for annee, mois, montant in autres_depenses_rows:
            # Vérifier que cette dépense est dans la période du projet
            try:
                mois_num = month_names.index(mois) + 1 if mois in month_names else 1
                entry_date = datetime.datetime(annee, mois_num, 1)
                
                if debut_projet <= entry_date <= fin_projet:
                    total_autres_achats += float(montant)
            except (ValueError, IndexError):
                continue
        
        data['autres_achats'] = total_autres_achats
        
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

    def refresh_cir(self, total_subventions):
        """Calcule et affiche le montant du CIR sous forme de tableau"""
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

            # Ajouter un titre pour le CIR
            self.budget_vbox.addWidget(QLabel("<b>CIR :</b>"))

            # Créer le tableau du CIR
            cir_table = QTableWidget()
            cir_table.setRowCount(1)  # Une seule ligne pour le CIR
            cir_table.setColumnCount(3)
            
            # Définir les en-têtes
            headers = ["Taux", "Coût éligible courant", "CIR attendue"]
            cir_table.setHorizontalHeaderLabels(headers)
            
            # Ajuster la taille du tableau
            cir_table.setMaximumHeight(80)  # Hauteur fixe pour une seule ligne
            cir_table.setMinimumHeight(80)
            cir_table.setMaximumWidth(350)  # Largeur fixe pour l'alignement
            
            # Configurer l'apparence du tableau
            cir_table.setAlternatingRowColors(True)
            cir_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            cir_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            
            # Style similaire au tableau des subventions
            cir_table.setStyleSheet("""
                QHeaderView::section {
                    font-size: 10px;
                    font-weight: bold;
                    padding: 2px;
                    background-color: #f0f0f0;
                    border: 1px solid #d0d0d0;
                }
                QTableWidget {
                    font-size: 9px;
                }
            """)
            
            # Ajuster automatiquement la largeur des colonnes
            header = cir_table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)  # Taux
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)  # Coût éligible courant
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)  # CIR attendue
            
            # Définir les largeurs de colonnes pour le CIR
            cir_table.setColumnWidth(0, 60)   # Taux
            cir_table.setColumnWidth(1, 150)  # Coût éligible courant
            cir_table.setColumnWidth(2, 120)  # CIR attendue

            # Remplir les données du tableau
            taux_k3_percent = k3 * 100  # Convertir en pourcentage
            
            # Vérifier si le CIR est applicable
            if montant_net_eligible > 0:
                # CIR applicable
                cir_attendu = montant_net_eligible * k3
                
                # Taux (avec virgule française)
                cir_table.setItem(0, 0, QTableWidgetItem(f"{taux_k3_percent:.1f}%".replace('.', ',')))
                
                # Coût éligible courant
                cir_table.setItem(0, 1, QTableWidgetItem(format_montant(montant_net_eligible)))
                
                # Subvention attendue
                cir_table.setItem(0, 2, QTableWidgetItem(format_montant(cir_attendu)))
                
            else:
                # CIR non applicable
                cir_table.setItem(0, 0, QTableWidgetItem(f"{taux_k3_percent:.1f}%".replace('.', ',')))
                cir_table.setItem(0, 1, QTableWidgetItem("Non applicable"))
                cir_table.setItem(0, 2, QTableWidgetItem("0 €"))
                
                # Colorer la ligne en rouge pour indiquer que le CIR n'est pas applicable
                for col in range(3):
                    item = cir_table.item(0, col)
                    if item:
                        item.setBackground(QColor(255, 235, 235))  # Fond rouge clair

            # Ajouter le tableau directement au layout principal (il sera aligné automatiquement)
            self.budget_vbox.addWidget(cir_table)

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
            
            # Afficher le CIR même s'il n'y a pas de subventions
            if self.has_cir_activated():
                # Vérifier si refresh_cir n'a pas déjà été appelé dans refresh_subventions
                has_subventions = False
                with sqlite3.connect(DB_PATH) as conn:
                    cursor = conn.cursor()
                    cursor.execute('SELECT COUNT(*) FROM subventions WHERE projet_id = ?', (self.projet_id,))
                    has_subventions = cursor.fetchone()[0] > 0
                
                if not has_subventions:
                    # Pas de subventions, mais CIR activé : afficher le CIR avec total_subventions = 0
                    self.refresh_cir(0)

    def refresh_subventions(self):
        """Affiche les montants des subventions sous forme de tableau (recalcule avec la logique de répartition)"""
        # Supprimer les anciens labels de subventions (s'ils existent)
        while self.budget_vbox.count() > 4:  # Garder seulement les 4 premiers items (titre + 3 coûts)
            item = self.budget_vbox.takeAt(self.budget_vbox.count() - 1)
            if item.widget():
                item.widget().deleteLater()

        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            
            # Récupérer toutes les subventions pour ce projet avec leurs paramètres
            try:
                cursor.execute('''
                    SELECT nom, mode_simplifie, montant_forfaitaire, depenses_temps_travail, coef_temps_travail, 
                           depenses_externes, coef_externes, depenses_autres_achats, coef_autres_achats, 
                           depenses_dotation_amortissements, coef_dotation_amortissements, cd, taux,
                           date_debut_subvention, date_fin_subvention, montant_subvention_max, depenses_eligibles_max
                    FROM subventions 
                    WHERE projet_id = ?
                ''', (self.projet_id,))
                subventions = cursor.fetchall()
            except sqlite3.OperationalError:
                # Fallback pour les anciennes bases de données
                try:
                    cursor.execute('''
                        SELECT nom, mode_simplifie, montant_forfaitaire, depenses_temps_travail, coef_temps_travail, 
                               depenses_externes, coef_externes, depenses_autres_achats, coef_autres_achats, 
                               depenses_dotation_amortissements, coef_dotation_amortissements, cd, taux
                        FROM subventions 
                        WHERE projet_id = ?
                    ''', (self.projet_id,))
                    # Ajouter des valeurs par défaut pour les colonnes manquantes
                    subventions = [list(row) + [None, None, 0, 0] for row in cursor.fetchall()]
                except sqlite3.OperationalError:
                    subventions = []

            if not subventions:
                return

            # Récupérer les dates du projet pour déterminer les années à calculer
            cursor.execute('SELECT date_debut, date_fin FROM projets WHERE id = ?', (self.projet_id,))
            date_row = cursor.fetchone()
            
            if not date_row or not date_row[0] or not date_row[1]:
                return
            
            try:
                import datetime
                debut_projet = datetime.datetime.strptime(date_row[0], '%m/%Y')
                fin_projet = datetime.datetime.strptime(date_row[1], '%m/%Y')
                # Calculer les années du projet
                annees_projet = list(range(debut_projet.year, fin_projet.year + 1))
            except ValueError:
                return

            # Ajouter un séparateur et titre
            self.budget_vbox.addWidget(QLabel(""))
            self.budget_vbox.addWidget(QLabel("<b>Subventions :</b>"))

            # Créer le tableau des subventions
            subv_table = QTableWidget()
            subv_table.setRowCount(len(subventions))
            subv_table.setColumnCount(6)
            
            # Définir les en-têtes avec sauts de ligne pour réduire la largeur
            headers = ["Nom", "Coût éligible\nmax", "Aide\nmax", "Taux", "Coût éligible\ncourant", "Subvention\nattendue"]
            subv_table.setHorizontalHeaderLabels(headers)
            
            # Ajuster la taille du tableau - largeur réduite
            subv_table.setMaximumHeight(120 + len(subventions) * 25)  # Hauteur adaptative
            subv_table.setMinimumHeight(60 + len(subventions) * 25)
            subv_table.setMinimumWidth(500)  # Largeur réduite
            subv_table.setMaximumWidth(550)  # Largeur maximum pour rester compact
            
            # Configurer l'apparence du tableau
            subv_table.setAlternatingRowColors(True)
            subv_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            subv_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            
            # Style du tableau
            subv_table.setStyleSheet("""
                QHeaderView::section {
                    font-size: 9px;
                    font-weight: bold;
                    padding: 2px;
                    background-color: #f0f0f0;
                    border: 1px solid #d0d0d0;
                    text-align: center;
                }
                QTableWidget {
                    font-size: 9px;
                }
            """)
            
            # Ajuster automatiquement la largeur des colonnes
            header = subv_table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)  # Nom
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)  # Coût éligible max
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)  # Aide max
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)  # Taux
            header.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)  # Coût éligible courant
            header.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)  # Subvention attendue
            
            # Définir les largeurs de colonnes
            subv_table.setColumnWidth(0, 80)   # Nom
            subv_table.setColumnWidth(1, 85)   # Coût éligible max
            subv_table.setColumnWidth(2, 75)   # Aide max
            subv_table.setColumnWidth(3, 50)   # Taux
            subv_table.setColumnWidth(4, 100)  # Coût éligible courant
            subv_table.setColumnWidth(5, 100)  # Subvention attendue

            total_subventions = 0

            # Importer la nouvelle méthode de calcul depuis SubventionDialog
            from subvention_dialog import SubventionDialog

            for row, subv in enumerate(subventions):
                (nom, mode_simplifie, montant_forfaitaire, dep_temps, coef_temps, dep_ext, coef_ext, 
                 dep_autres, coef_autres, dep_amort, coef_amort, cd, taux, 
                 date_debut_subv, date_fin_subv, montant_max, depenses_max) = subv
                
                # Construire le dictionnaire de données de subvention pour le calcul
                subvention_data = {
                    'nom': nom,
                    'mode_simplifie': mode_simplifie or 0,
                    'montant_forfaitaire': montant_forfaitaire or 0,
                    'depenses_temps_travail': dep_temps or 0,
                    'coef_temps_travail': coef_temps or 1,
                    'depenses_externes': dep_ext or 0,
                    'coef_externes': coef_ext or 1,
                    'depenses_autres_achats': dep_autres or 0,
                    'coef_autres_achats': coef_autres or 1,
                    'depenses_dotation_amortissements': dep_amort or 0,
                    'coef_dotation_amortissements': coef_amort or 1,
                    'cd': cd or 1,
                    'taux': taux or 100,
                    'date_debut_subvention': date_debut_subv,
                    'date_fin_subvention': date_fin_subv
                }
                
                # Calculer le montant total estimé avec la nouvelle logique (somme sur toutes les années du projet)
                montant_total_estime = 0
                assiette_totale_courante = 0
                
                for annee in annees_projet:
                    # Calculer l'assiette éligible pour cette année avec la nouvelle logique de redistribution
                    assiette_annee = self.calculate_period_eligible_expenses_with_redistribution(
                        cursor, self.projet_id, subvention_data, annee, None
                    )
                    assiette_totale_courante += assiette_annee
                    
                    # Calculer la subvention pour cette année
                    if mode_simplifie:
                        # Mode forfaitaire
                        montant_annee = montant_forfaitaire or 0
                    else:
                        # Mode détaillé - appliquer le taux sur l'assiette éligible
                        montant_avant_plafond = assiette_annee * (taux / 100.0)
                        
                        # Appliquer les plafonds si définis
                        if montant_max and montant_max > 0:
                            montant_annee = min(montant_avant_plafond, montant_max)
                        else:
                            montant_annee = montant_avant_plafond
                    
                    montant_total_estime += montant_annee
                
                # Nom de la subvention
                subv_table.setItem(row, 0, QTableWidgetItem(nom or ""))
                
                # Coût éligible max - afficher "---" pour les subventions forfaitaires
                if mode_simplifie:
                    subv_table.setItem(row, 1, QTableWidgetItem("---"))
                else:
                    subv_table.setItem(row, 1, QTableWidgetItem(format_montant(depenses_max) if depenses_max and depenses_max > 0 else "Illimité"))
                
                # Aide max
                subv_table.setItem(row, 2, QTableWidgetItem(format_montant(montant_max) if montant_max and montant_max > 0 else "Illimité"))

                # Taux (avec virgule française)
                subv_table.setItem(row, 3, QTableWidgetItem(f"{taux:.1f}%".replace('.', ',') if taux else "0,0%"))
                
                # Coût éligible courant (recalculé avec la nouvelle logique)
                if mode_simplifie:
                    # Pour le mode forfaitaire, afficher "---" 
                    subv_table.setItem(row, 4, QTableWidgetItem("---"))
                else:
                    # Pour le mode détaillé, afficher l'assiette éligible recalculée
                    subv_table.setItem(row, 4, QTableWidgetItem(format_montant(assiette_totale_courante)))
                
                # Subvention attendue (recalculée avec la nouvelle logique)
                subv_table.setItem(row, 5, QTableWidgetItem(format_montant(montant_total_estime)))
                
                # Ajouter au total
                total_subventions += montant_total_estime

            # Ajouter le tableau au layout
            self.budget_vbox.addWidget(subv_table)

            # Calculer et afficher le CIR si le projet l'a activé
            if self.has_cir_activated():
                self.refresh_cir(total_subventions)
            subv_table.setMinimumHeight(60 + len(subventions) * 25)
            subv_table.setMinimumWidth(500)  # Largeur réduite
            subv_table.setMaximumWidth(550)  # Largeur maximum pour rester compact
            
            # Configurer l'apparence du tableau
            subv_table.setAlternatingRowColors(True)
            subv_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            subv_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            
            # Style du tableau
            subv_table.setStyleSheet("""
                QHeaderView::section {
                    font-size: 9px;
                    font-weight: bold;
                    padding: 2px;
                    background-color: #f0f0f0;
                    border: 1px solid #d0d0d0;
                    text-align: center;
                }
                QTableWidget {
                    font-size: 9px;
                }
            """)
            
            # Ajuster automatiquement la largeur des colonnes
            header = subv_table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)  # Nom
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)  # Coût éligible max
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)  # Aide max
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)  # Taux
            header.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)  # Coût éligible courant
            header.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)  # Subvention attendue
            
            # Définir les largeurs de colonnes
            subv_table.setColumnWidth(0, 80)   # Nom
            subv_table.setColumnWidth(1, 85)   # Coût éligible max
            subv_table.setColumnWidth(2, 75)   # Aide max
            subv_table.setColumnWidth(3, 50)   # Taux
            subv_table.setColumnWidth(4, 100)  # Coût éligible courant
            subv_table.setColumnWidth(5, 100)  # Subvention attendue

            total_subventions = 0

            # La boucle pour remplir le tableau est déjà gérée plus haut dans la fonction

            # Ajouter le tableau au layout
            self.budget_vbox.addWidget(subv_table)

            # Calculer et afficher le CIR si le projet l'a activé
            if self.has_cir_activated():
                self.refresh_cir(total_subventions)

    def refresh_project_data(self):
        """Rafraîchit toutes les données du projet dans la page de détails"""
        try:
            # Fermer et rouvrir la fenêtre avec les nouvelles données
            # C'est la méthode la plus simple pour tout rafraîchir
            projet_id = self.projet_id
            parent = self.parent()
            
            # Fermer cette fenêtre
            self.close()
            
            # Rouvrir immédiatement une nouvelle fenêtre avec les données actualisées
            from project_details_dialog import ProjectDetailsDialog
            new_dialog = ProjectDetailsDialog(parent, projet_id)
            new_dialog.show()
            
        except Exception as e:
            QMessageBox.critical(
                self, "Erreur de rafraîchissement", 
                f"Impossible de rafraîchir les données :\n{str(e)}"
            )

    def calculate_period_eligible_expenses_with_redistribution(self, cursor, project_id, subvention_data, year, month):
        """Calcule les dépenses éligibles pour une période donnée avec redistribution"""
        # Extraire les paramètres de la subvention
        depenses_temps_travail = subvention_data.get('depenses_temps_travail', 0)
        coef_temps_travail = subvention_data.get('coef_temps_travail', 1)
        depenses_externes = subvention_data.get('depenses_externes', 0)
        coef_externes = subvention_data.get('coef_externes', 1)
        depenses_autres_achats = subvention_data.get('depenses_autres_achats', 0)
        coef_autres_achats = subvention_data.get('coef_autres_achats', 1)
        depenses_dotation_amortissements = subvention_data.get('depenses_dotation_amortissements', 0)
        coef_dotation_amortissements = subvention_data.get('coef_dotation_amortissements', 1)
        
        depenses_eligibles = 0
        
        # Noms des mois en français
        month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                      "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        
        # Temps de travail éligible
        if depenses_temps_travail:
            if month is not None:
                # Mois spécifique
                temps_travail_cout = self.calculate_temps_travail_for_period_with_redistribution(
                    cursor, project_id, year, month
                )
            else:
                # Année complète
                temps_travail_cout = 0
                for m in range(1, 13):
                    temps_travail_cout += self.calculate_temps_travail_for_period_with_redistribution(
                        cursor, project_id, year, m
                    )
            
            temps_travail_eligible = temps_travail_cout * coef_temps_travail
            depenses_eligibles += temps_travail_eligible
        
        # Dépenses externes éligibles avec redistribution
        if depenses_externes:
            depenses_ext = self.calculate_redistributed_expenses(
                cursor, project_id, year, month, 'depenses'
            )
            depenses_ext_eligible = depenses_ext * coef_externes
            depenses_eligibles += depenses_ext_eligible
        
        # Autres dépenses éligibles avec redistribution
        if depenses_autres_achats:
            autres_dep = self.calculate_redistributed_expenses(
                cursor, project_id, year, month, 'autres_depenses'
            )
            autres_dep_eligible = autres_dep * coef_autres_achats
            depenses_eligibles += autres_dep_eligible
        
        # Dotations aux amortissements (déjà réparties par nature)
        if depenses_dotation_amortissements:
            amortissements = self.calculate_amortissements_for_period(
                cursor, project_id, year, month
            )
            amortissements_eligible = amortissements * coef_dotation_amortissements
            depenses_eligibles += amortissements_eligible
        
        return depenses_eligibles

    def calculate_temps_travail_for_period_with_redistribution(self, cursor, project_id, year, month):
        """Calcule le coût du temps de travail avec redistribution pour une période"""
        # Récupérer les données du projet pour la redistribution
        cursor.execute('SELECT date_debut, date_fin FROM projets WHERE id = ?', (project_id,))
        date_row = cursor.fetchone()
        
        if not date_row or not date_row[0] or not date_row[1]:
            return 0
        
        try:
            import datetime
            debut_projet = datetime.datetime.strptime(date_row[0], '%m/%Y')
            fin_projet = datetime.datetime.strptime(date_row[1], '%m/%Y')
        except ValueError:
            return 0
        
        # Calculer le nombre total de mois du projet
        total_mois = (fin_projet.year - debut_projet.year) * 12 + fin_projet.month - debut_projet.month + 1
        
        # Vérifier si le mois demandé est dans la période du projet
        try:
            periode_demandee = datetime.datetime(year, month, 1)
            if not (debut_projet <= periode_demandee <= fin_projet):
                return 0
        except ValueError:
            return 0
        
        # Récupérer tout le temps de travail du projet
        cursor.execute("""
            SELECT tt.categorie, tt.jours
            FROM temps_travail tt
            WHERE tt.projet_id = ? AND tt.annee = ?
        """, (project_id, year))
        
        temps_travail_rows = cursor.fetchall()
        cout_total = 0
        
        mapping_categories = {
            "Stagiaire Projet": "STP",
            "Assistante / opérateur": "AOP", 
            "Technicien": "TEP",
            "Junior": "IJP",
            "Senior": "ISP",
            "Expert": "EDP",
            "Collaborateur moyen": "MOY"
        }
        
        for categorie, jours in temps_travail_rows:
            categorie_code = mapping_categories.get(categorie, "")
            if not categorie_code:
                continue
            
            # Récupérer le coût pour cette catégorie
            cursor.execute("""
                SELECT montant_charge 
                FROM categorie_cout 
                WHERE categorie = ? AND annee = ?
            """, (categorie_code, year))
            
            cout_row = cursor.fetchone()
            if cout_row and cout_row[0]:
                montant_charge = float(cout_row[0])
                # Redistribuer les jours sur tous les mois du projet
                jours_par_mois = jours / total_mois
                cout_total += jours_par_mois * montant_charge
        
        return cout_total

    def calculate_redistributed_expenses(self, cursor, project_id, year, month, table_name):
        """Calcule les dépenses redistribuées pour une période"""
        # Récupérer les données du projet pour la redistribution
        cursor.execute('SELECT date_debut, date_fin FROM projets WHERE id = ?', (project_id,))
        date_row = cursor.fetchone()
        
        if not date_row or not date_row[0] or not date_row[1]:
            return 0
        
        try:
            import datetime
            debut_projet = datetime.datetime.strptime(date_row[0], '%m/%Y')
            fin_projet = datetime.datetime.strptime(date_row[1], '%m/%Y')
        except ValueError:
            return 0
        
        # Calculer le nombre total de mois du projet
        total_mois = (fin_projet.year - debut_projet.year) * 12 + fin_projet.month - debut_projet.month + 1
        
        # Vérifier si le mois demandé est dans la période du projet
        if month is not None:
            try:
                periode_demandee = datetime.datetime(year, month, 1)
                if not (debut_projet <= periode_demandee <= fin_projet):
                    return 0
            except ValueError:
                return 0
        
        # Récupérer toutes les dépenses de cette année pour ce projet
        cursor.execute(f"""
            SELECT SUM(montant)
            FROM {table_name}
            WHERE projet_id = ? AND annee = ?
        """, (project_id, year))
        
        total_annee = cursor.fetchone()[0] or 0
        
        if month is not None:
            # Pour un mois spécifique, redistribuer le total annuel
            return float(total_annee) / total_mois
        else:
            # Pour toute l'année
            return float(total_annee)

    def calculate_amortissements_for_period(self, cursor, project_id, year, month):
        """Calcule les amortissements pour une période"""
        cursor.execute("""
            SELECT montant, date_achat, duree 
            FROM investissements 
            WHERE projet_id = ?
        """, (project_id,))
        
        amortissements_total = 0
        
        for montant, date_achat, duree in cursor.fetchall():
            try:
                import datetime
                # Convertir la date d'achat en datetime
                achat_date = datetime.datetime.strptime(date_achat, '%m/%Y')
                
                # La dotation commence le mois suivant l'achat
                debut_amort = achat_date.replace(day=1)
                debut_amort = datetime.datetime(debut_amort.year, debut_amort.month, 1) + datetime.timedelta(days=32)
                debut_amort = debut_amort.replace(day=1)
                
                # Calculer la dotation mensuelle
                dotation_mensuelle = float(montant) / (int(duree) * 12)
                
                if month is not None:
                    # Pour un mois spécifique
                    periode_demandee = datetime.datetime(year, month, 1)
                    
                    # Vérifier si ce mois est dans la période d'amortissement
                    fin_amort = datetime.datetime(debut_amort.year + int(duree), debut_amort.month, 1)
                    
                    if debut_amort <= periode_demandee < fin_amort:
                        amortissements_total += dotation_mensuelle
                else:
                    # Pour toute l'année
                    for m in range(1, 13):
                        periode_m = datetime.datetime(year, m, 1)
                        fin_amort = datetime.datetime(debut_amort.year + int(duree), debut_amort.month, 1)
                        
                        if debut_amort <= periode_m < fin_amort:
                            amortissements_total += dotation_mensuelle
                            
            except Exception:
                continue
        
        return amortissements_total

