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
        
        # Champ détails avec un QTextEdit en lecture seule pour une meilleure gestion du texte long
        details_container = QVBoxLayout()
        details_title = QLabel("<b>Détails :</b>")
        details_container.addWidget(details_title)
        
        details_text = QTextEdit()
        details_text.setPlainText(projet[2] if projet[2] else "Aucun détail")
        details_text.setReadOnly(True)
        details_text.setMaximumHeight(100)  # Hauteur fixe mais scrollable
        details_text.setMaximumWidth(400)   # Largeur légèrement augmentée
        details_text.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        details_text.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        # Style pour que ça ressemble plus à un affichage qu'à un champ d'édition
        details_text.setStyleSheet("""
            QTextEdit {
                background-color: #f8f8f8;
                border: 1px solid #ddd;
                font-family: inherit;
                font-size: inherit;
            }
        """)
        details_container.addWidget(details_text)
        left_column.addLayout(details_container)
        
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
        # Ajout du bouton "Imprimer la page"
        print_page_btn = QPushButton("Imprimer la page")
        btn_hbox.addWidget(print_page_btn)
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
        # Connexion du bouton "Imprimer la page"
        print_page_btn.clicked.connect(self.print_page)

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

    def get_project_data_for_subventions(self, date_debut_subv=None, date_fin_subv=None):
        """Récupère les données du projet pour calculer les subventions AVEC REDISTRIBUTION AUTOMATIQUE
        Si les dates de subvention sont fournies, calcule sur cette période, sinon sur tout le projet"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        data = {
            'temps_travail_total': 0,
            'depenses_externes': 0,
            'autres_achats': 0,
            'amortissements': 0
        }
        
        import datetime
        
        # Déterminer les dates de calcul
        if date_debut_subv and date_fin_subv:
            # Utiliser les dates de subvention fournies
            try:
                debut_calcul = datetime.datetime.strptime(date_debut_subv, '%m/%Y')
                fin_calcul = datetime.datetime.strptime(date_fin_subv, '%m/%Y')
            except ValueError:
                conn.close()
                return data
        else:
            # Utiliser les dates du projet
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (self.projet_id,))
            date_row = cursor.fetchone()
            if not date_row or not date_row[0] or not date_row[1]:
                conn.close()
                return data
            
            try:
                debut_calcul = datetime.datetime.strptime(date_row[0], '%m/%Y')
                fin_calcul = datetime.datetime.strptime(date_row[1], '%m/%Y')
            except ValueError:
                conn.close()
                return data
        
        # UTILISER LA MÊME LOGIQUE QUE SubventionDialog - Redistribution sur la période spécifiée
        # Créer une classe helper pour accéder aux méthodes de redistribution (copie de SubventionDialog)
        class RedistributionHelper:
            def __init__(self):
                pass
                
            def calculate_redistributed_temps_travail(self, cursor, project_id, year, month, cost_type):
                """Méthode simplifiée de redistribution du temps de travail basée sur compte_resultat_display"""
                # Récupérer les dates du projet
                cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id = ?", (project_id,))
                projet_info = cursor.fetchone()
                if not projet_info or not projet_info[0] or not projet_info[1]:
                    return 0
                    
                debut_projet = datetime.datetime.strptime(projet_info[0], '%m/%Y')
                fin_projet = datetime.datetime.strptime(projet_info[1], '%m/%Y')
                
                # Vérifier si l'année est dans la période du projet
                if year < debut_projet.year or year > fin_projet.year:
                    return 0
                
                month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                              "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
                
                # Vérifier si tous les couples (membre_id, categorie) n'ont qu'une seule entrée dans l'année
                cursor.execute("""
                    SELECT membre_id, categorie, COUNT(*) as nb_entries
                    FROM temps_travail 
                    WHERE projet_id = ? AND annee = ?
                    GROUP BY membre_id, categorie
                """, (project_id, year))
                
                couples_entries = cursor.fetchall()
                if not couples_entries:
                    return 0
                
                # Vérifier si TOUS les couples n'ont qu'une seule entrée
                all_single_entry = all(nb_entries == 1 for _, _, nb_entries in couples_entries)
                
                if not all_single_entry:
                    # PAS DE REDISTRIBUTION - utiliser les données réelles
                    mois_nom = month_names[month - 1]
                    cursor.execute("""
                        SELECT SUM(tt.jours * cc.montant_charge)
                        FROM temps_travail tt
                        JOIN categorie_cout cc ON cc.libelle = tt.categorie AND cc.annee = tt.annee
                        WHERE tt.projet_id = ? AND tt.annee = ? AND tt.mois = ?
                    """, (project_id, year, mois_nom))
                    
                    result = cursor.fetchone()
                    return float(result[0]) if result and result[0] else 0
                
                else:
                    # REDISTRIBUTION - calculer le montant redistributé pour ce mois
                    total_mois_projet = (fin_projet.year - debut_projet.year) * 12 + (fin_projet.month - debut_projet.month) + 1
                    
                    # Récupérer le coût total de l'année et le redistribuer
                    cursor.execute("""
                        SELECT SUM(tt.jours * cc.montant_charge)
                        FROM temps_travail tt
                        JOIN categorie_cout cc ON cc.libelle = tt.categorie AND cc.annee = tt.annee
                        WHERE tt.projet_id = ? AND tt.annee = ?
                    """, (project_id, year))
                    
                    result = cursor.fetchone()
                    total_annee = float(result[0]) if result and result[0] else 0
                    
                    # Redistribuer sur tous les mois du projet dans cette année
                    if total_annee > 0:
                        debut_annee_projet = max(debut_projet, datetime.datetime(year, 1, 1))
                        fin_annee_projet = min(fin_projet, datetime.datetime(year, 12, 31))
                        
                        mois_dans_annee = (fin_annee_projet.year - debut_annee_projet.year) * 12 + (fin_annee_projet.month - debut_annee_projet.month) + 1
                        
                        # Vérifier si le mois demandé est dans la période du projet
                        mois_demande = datetime.datetime(year, month, 1)
                        if debut_annee_projet <= mois_demande <= fin_annee_projet:
                            return total_annee / mois_dans_annee
                    
                    return 0
            
            def calculate_redistributed_expenses(self, cursor, project_id, year, month, table_name):
                """Calcule les dépenses redistributées pour une période"""
                # Récupérer les dates du projet
                cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id = ?", (project_id,))
                projet_info = cursor.fetchone()
                if not projet_info or not projet_info[0] or not projet_info[1]:
                    return 0
                    
                debut_projet = datetime.datetime.strptime(projet_info[0], '%m/%Y')
                fin_projet = datetime.datetime.strptime(projet_info[1], '%m/%Y')
                
                # Vérifier si l'année est dans la période du projet
                if year < debut_projet.year or year > fin_projet.year:
                    return 0
                
                month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                              "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
                
                # Vérifier s'il y a plusieurs entrées dans l'année (nécessite redistribution)
                cursor.execute(f"""
                    SELECT COUNT(DISTINCT mois) 
                    FROM {table_name}
                    WHERE projet_id = ? AND annee = ?
                """, (project_id, year))
                
                nb_mois_entries = cursor.fetchone()[0] or 0
                
                if nb_mois_entries <= 1:
                    # PAS DE REDISTRIBUTION - utiliser les données réelles
                    mois_nom = month_names[month - 1]
                    cursor.execute(f"""
                        SELECT SUM(montant)
                        FROM {table_name}
                        WHERE projet_id = ? AND annee = ? AND mois = ?
                    """, (project_id, year, mois_nom))
                    
                    result = cursor.fetchone()
                    return float(result[0]) if result and result[0] else 0
                
                else:
                    # REDISTRIBUTION - calculer le montant redistributé pour ce mois
                    # Récupérer le total de l'année
                    cursor.execute(f"""
                        SELECT SUM(montant)
                        FROM {table_name}
                        WHERE projet_id = ? AND annee = ?
                    """, (project_id, year))
                    
                    result = cursor.fetchone()
                    total_annee = float(result[0]) if result and result[0] else 0
                    
                    if total_annee > 0:
                        # Calculer les mois du projet dans cette année
                        debut_annee_projet = max(debut_projet, datetime.datetime(year, 1, 1))
                        fin_annee_projet = min(fin_projet, datetime.datetime(year, 12, 31))
                        
                        mois_dans_annee = (fin_annee_projet.year - debut_annee_projet.year) * 12 + (fin_annee_projet.month - debut_annee_projet.month) + 1
                        
                        # Vérifier si le mois demandé est dans la période du projet
                        mois_demande = datetime.datetime(year, month, 1)
                        if debut_annee_projet <= mois_demande <= fin_annee_projet:
                            return total_annee / mois_dans_annee
                    
                    return 0
            
            def calculate_amortissement_for_period(self, cursor, project_id, year, month):
                """Calcule les amortissements pour une période"""
                cursor.execute("""
                    SELECT montant, date_achat, duree 
                    FROM investissements 
                    WHERE projet_id = ?
                """, (project_id,))
                
                amortissements_total = 0
                
                for montant, date_achat, duree in cursor.fetchall():
                    try:
                        # Parser la date d'achat
                        date_parts = date_achat.split('/')
                        if len(date_parts) == 2:  # Format MM/yyyy
                            mois_achat = int(date_parts[0])
                            annee_achat = int(date_parts[1])
                        elif len(date_parts) == 3:  # Format dd/MM/yyyy
                            annee_achat = int(date_parts[2])
                            mois_achat = int(date_parts[1])
                        else:
                            continue
                            
                        debut_amort = datetime.datetime(annee_achat, mois_achat, 1)
                        
                        # Calculer la fin d'amortissement
                        fin_month = mois_achat + int(duree) - 1
                        fin_year = annee_achat + (fin_month - 1) // 12
                        fin_month = ((fin_month - 1) % 12) + 1
                        fin_amort = datetime.datetime(fin_year, fin_month, 1)
                        
                        # Vérifier si le mois demandé est dans la période d'amortissement
                        mois_demande = datetime.datetime(year, month, 1)
                        if debut_amort <= mois_demande <= fin_amort:
                            # Calculer la dotation mensuelle
                            amortissement_mensuel = float(montant) / int(duree)
                            amortissements_total += amortissement_mensuel
                            
                    except (ValueError, TypeError):
                        continue
                
                return amortissements_total
        
        # Créer l'instance helper
        helper = RedistributionHelper()
        
        temps_travail_total = 0
        depenses_externes = 0
        autres_achats = 0
        amortissements = 0
        
        # Parcourir tous les mois de la période de calcul
        current_date = debut_calcul.replace(day=1)
        while current_date <= fin_calcul:
            year = current_date.year
            month = current_date.month
            
            # 1. TEMPS DE TRAVAIL - avec redistribution
            cout_temps_travail = helper.calculate_redistributed_temps_travail(
                cursor, self.projet_id, year, month, 'montant_charge'
            )
            temps_travail_total += cout_temps_travail
            
            # 2. DÉPENSES EXTERNES - avec redistribution
            depenses_externes += helper.calculate_redistributed_expenses(
                cursor, self.projet_id, year, month, 'depenses'
            )
            
            # 3. AUTRES ACHATS - avec redistribution
            autres_achats += helper.calculate_redistributed_expenses(
                cursor, self.projet_id, year, month, 'autres_depenses'
            )
            
            # 4. AMORTISSEMENTS - avec redistribution
            amortissements += helper.calculate_amortissement_for_period(
                cursor, self.projet_id, year, month
            )
            
            # Passer au mois suivant
            if current_date.month == 12:
                current_date = current_date.replace(year=current_date.year + 1, month=1)
            else:
                current_date = current_date.replace(month=current_date.month + 1)
        
        data['temps_travail_total'] = temps_travail_total
        data['depenses_externes'] = depenses_externes
        data['autres_achats'] = autres_achats
        data['amortissements'] = amortissements
        
        conn.close()
        return data

    def refresh_cir(self, total_subventions):
        """Calcule et affiche le montant du CIR sous forme de tableau avec répartition mensuelle"""
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            
            # Récupérer les coefficients CIR et dates du projet
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

            # Calculer le CIR mois par mois avec la bonne répartition des subventions
            montant_net_eligible_total = 0
            cir_total = 0
            
            # Récupérer toutes les subventions pour ce projet
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

            # Préparer les données des subventions
            subventions_data = []
            if subventions:
                from subvention_dialog import SubventionDialog
                
                for subv in subventions:
                    (nom, mode_simplifie, montant_forfaitaire, dep_temps, coef_temps, dep_ext, coef_ext, 
                     dep_autres, coef_autres, dep_amort, coef_amort, cd, taux, 
                     date_debut_subv, date_fin_subv, montant_max, depenses_max) = subv
                    
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
                        'date_fin_subvention': date_fin_subv,
                        'montant_subvention_max': montant_max or 0,
                        'depenses_eligibles_max': depenses_max or 0
                    }
                    subventions_data.append(subvention_data)
            
            # UTILISER EXACTEMENT LA MÊME LOGIQUE QUE LE COMPTE DE RÉSULTAT
            # Importer la classe CompteResultatDisplay pour utiliser ses méthodes
            from compte_resultat_display import CompteResultatDisplay
            temp_compte_resultat = CompteResultatDisplay(self, {'project_ids': [self.projet_id], 'years': list(range(debut_projet.year, fin_projet.year + 1))})
            
            # Calculer le CIR mois par mois en utilisant la même méthode que le compte de résultat
            current_date = debut_projet
            while current_date <= fin_projet:
                year = current_date.year
                month = current_date.month
                
                # Calculer les éléments pour ce mois EXACTEMENT comme dans le compte de résultat
                temps_travail_cout_mois = temp_compte_resultat.calculate_redistributed_temps_travail(
                    cursor, self.projet_id, year, month, 'montant_charge'
                )
                amortissements_mois = temp_compte_resultat.calculate_amortissement_for_period(
                    cursor, self.projet_id, year, month
                )
                projet_info = (date_row[0], date_row[1])
                total_subventions_mois = temp_compte_resultat.calculate_smart_distributed_subvention(
                    cursor, self.projet_id, year, month, projet_info
                )
                
                # Calculer l'assiette éligible pour ce mois
                assiette_mensuelle = (temps_travail_cout_mois * k1) + (amortissements_mois * k2) - total_subventions_mois
                
                # Calculer le CIR pour ce mois en utilisant EXACTEMENT la même méthode que le compte de résultat
                cir_mensuel = temp_compte_resultat.calculate_distributed_cir(cursor, year, month)
                
                if assiette_mensuelle > 0 and cir_mensuel > 0:
                    montant_net_eligible_total += assiette_mensuelle
                    cir_total += cir_mensuel
                
                # Passer au mois suivant
                if current_date.month == 12:
                    current_date = datetime.datetime(current_date.year + 1, 1, 1)
                else:
                    current_date = datetime.datetime(current_date.year, current_date.month + 1, 1)

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

            # Remplir les données du tableau avec les valeurs calculées mois par mois
            taux_k3_percent = k3 * 100  # Convertir en pourcentage
            
            # Vérifier si le CIR est applicable
            if montant_net_eligible_total > 0:
                # CIR applicable
                # Taux (avec virgule française)
                cir_table.setItem(0, 0, QTableWidgetItem(f"{taux_k3_percent:.1f}%".replace('.', ',')))
                
                # Coût éligible courant (total calculé mois par mois)
                cir_table.setItem(0, 1, QTableWidgetItem(format_montant(montant_net_eligible_total)))
                
                # CIR attendu (total calculé mois par mois)
                cir_table.setItem(0, 2, QTableWidgetItem(format_montant(cir_total)))
                
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

    def calculate_total_subventions_for_cir(self):
        """Calcule le total des subventions avec la même logique que refresh_subventions"""
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            
            # Récupérer toutes les subventions pour ce projet
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
                    return 0

            if not subventions:
                return 0

            # Récupérer les dates du projet pour déterminer les années à calculer
            cursor.execute('SELECT date_debut, date_fin FROM projets WHERE id = ?', (self.projet_id,))
            date_row = cursor.fetchone()
            
            if not date_row or not date_row[0] or not date_row[1]:
                return 0
            
            try:
                import datetime
                debut_projet = datetime.datetime.strptime(date_row[0], '%m/%Y')
                fin_projet = datetime.datetime.strptime(date_row[1], '%m/%Y')
                # Calculer les années du projet
                annees_projet = list(range(debut_projet.year, fin_projet.year + 1))
            except ValueError:
                return 0

            total_subventions = 0

            # Importer la méthode de calcul depuis SubventionDialog
            from subvention_dialog import SubventionDialog

            for subv in subventions:
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
                    'date_fin_subvention': date_fin_subv,
                    'montant_subvention_max': montant_max or 0,
                    'depenses_eligibles_max': depenses_max or 0
                }
                
                # Calculer le montant total estimé avec la logique de SubventionDialog
                montant_total_estime = 0
                
                for annee in annees_projet:
                    # Utiliser la méthode de SubventionDialog pour calculer la subvention
                    montant_annee = SubventionDialog.calculate_distributed_subvention(
                        self.projet_id, subvention_data, annee, None
                    )
                    montant_total_estime += montant_annee
                
                # Ajouter au total
                total_subventions += montant_total_estime

            return total_subventions

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
                
                # Calculer le montant total estimé SEULEMENT sur la période de subvention
                montant_total_estime = 0
                assiette_totale_courante = 0
                
                # Calculer les années de la période de subvention (pas du projet)
                if date_debut_subv and date_fin_subv:
                    try:
                        debut_subv = datetime.datetime.strptime(date_debut_subv, '%m/%Y')
                        fin_subv = datetime.datetime.strptime(date_fin_subv, '%m/%Y')
                        annees_subvention = list(range(debut_subv.year, fin_subv.year + 1))
                    except ValueError:
                        annees_subvention = annees_projet  # Fallback sur années projet
                else:
                    annees_subvention = annees_projet  # Fallback sur années projet
                
                for annee in annees_subvention:
                    # Calculer l'assiette éligible pour cette année avec la nouvelle logique de redistribution
                    assiette_annee = self.calculate_period_eligible_expenses_with_redistribution(
                        cursor, self.projet_id, subvention_data, annee, None
                    )
                    assiette_totale_courante += assiette_annee
                
                # Appliquer le plafond du coût éligible max à l'assiette totale
                assiette_plafonnee = assiette_totale_courante
                if not mode_simplifie and depenses_max and depenses_max > 0:
                    assiette_plafonnee = min(assiette_totale_courante, depenses_max)
                
                # Calculer la subvention attendue
                if mode_simplifie:
                    # Mode forfaitaire
                    montant_total_estime = montant_forfaitaire or 0
                else:
                    # Mode détaillé - appliquer le taux sur l'assiette éligible plafonnée
                    montant_avant_plafond = assiette_plafonnee * (taux / 100.0)
                    
                    # Appliquer le plafond du montant max si défini
                    if montant_max and montant_max > 0:
                        montant_total_estime = min(montant_avant_plafond, montant_max)
                    else:
                        montant_total_estime = montant_avant_plafond
                
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
                    # Pour le mode détaillé, afficher l'assiette éligible plafonnée
                    subv_table.setItem(row, 4, QTableWidgetItem(format_montant(assiette_plafonnee)))
                
                # Subvention attendue (recalculée avec la nouvelle logique)
                subv_table.setItem(row, 5, QTableWidgetItem(format_montant(montant_total_estime)))
                
                # Ajouter au total
                total_subventions += montant_total_estime

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
        """Calcule les dépenses éligibles pour une période donnée avec redistribution SIMPLIFIÉE"""
        # Utiliser la même logique que SubventionDialog : récupérer les données totales et redistribuer
        
        # Récupérer les dates de subvention et utiliser get_project_data_for_subventions avec ces dates
        date_debut_subv = subvention_data.get('date_debut_subvention')
        date_fin_subv = subvention_data.get('date_fin_subvention')
        data = self.get_project_data_for_subventions(date_debut_subv, date_fin_subv)
        
        # Extraire les paramètres de la subvention
        depenses_temps_travail = subvention_data.get('depenses_temps_travail', 0)
        coef_temps_travail = subvention_data.get('coef_temps_travail', 1)
        depenses_externes = subvention_data.get('depenses_externes', 0)
        coef_externes = subvention_data.get('coef_externes', 1)
        depenses_autres_achats = subvention_data.get('depenses_autres_achats', 0)
        coef_autres_achats = subvention_data.get('coef_autres_achats', 1)
        depenses_dotation_amortissements = subvention_data.get('depenses_dotation_amortissements', 0)
        coef_dotation_amortissements = subvention_data.get('coef_dotation_amortissements', 1)
        cd = subvention_data.get('cd', 1)
        
        # Récupérer les dates de subvention pour calculer la proportion
        date_debut_subv = subvention_data.get('date_debut_subvention')
        date_fin_subv = subvention_data.get('date_fin_subvention')
        
        if not date_debut_subv or not date_fin_subv:
            return 0
        
        try:
            import datetime
            debut_subv = datetime.datetime.strptime(date_debut_subv, '%m/%Y')
            fin_subv = datetime.datetime.strptime(date_fin_subv, '%m/%Y')
            duree_subv_mois = (fin_subv.year - debut_subv.year) * 12 + (fin_subv.month - debut_subv.month) + 1
        except ValueError:
            return 0
        
        # Calculer l'assiette éligible totale
        assiette_totale = 0
        
        if depenses_temps_travail:
            # Temps de travail avec coefficient de charge et coefficient d'éligibilité
            assiette_totale += data['temps_travail_total'] * cd * coef_temps_travail
        
        if depenses_externes:
            assiette_totale += data['depenses_externes'] * coef_externes
        
        if depenses_autres_achats:
            assiette_totale += data['autres_achats'] * coef_autres_achats
        
        if depenses_dotation_amortissements:
            assiette_totale += data['amortissements'] * coef_dotation_amortissements
        
        # Si on demande pour une année complète (month = None), calculer la part de cette année
        if month is None:
            # Vérifier si l'année demandée est dans la période de subvention
            if not date_debut_subv or not date_fin_subv:
                return assiette_totale  # Si pas de dates spécifiques, retourner l'assiette totale
                
            try:
                debut_subv = datetime.datetime.strptime(date_debut_subv, '%m/%Y')
                fin_subv = datetime.datetime.strptime(date_fin_subv, '%m/%Y')
                
                # Calculer les bornes de l'année demandée dans la période de subvention
                debut_annee_subv = max(debut_subv, datetime.datetime(year, 1, 1))
                fin_annee_subv = min(fin_subv, datetime.datetime(year, 12, 31))
                
                # Si l'année n'intersecte pas avec la période de subvention
                if debut_annee_subv > fin_annee_subv:
                    return 0
                
                # Calculer le nombre de mois de l'année dans la période de subvention
                mois_annee_dans_subv = (fin_annee_subv.year - debut_annee_subv.year) * 12 + (fin_annee_subv.month - debut_annee_subv.month) + 1
                
                # Calculer le nombre total de mois de la période de subvention
                total_mois_subv = (fin_subv.year - debut_subv.year) * 12 + (fin_subv.month - debut_subv.month) + 1
                
                # Retourner la part proportionnelle de l'assiette pour cette année
                return assiette_totale * (mois_annee_dans_subv / total_mois_subv)
                
            except ValueError:
                return assiette_totale  # En cas d'erreur, retourner l'assiette totale
        
        # Si on demande pour un mois spécifique, vérifier s'il est dans la période de subvention
        try:
            mois_demande = datetime.datetime(year, month, 1)
            if debut_subv <= mois_demande <= fin_subv:
                # Redistribuer l'assiette sur la durée de la subvention
                return assiette_totale / duree_subv_mois
            else:
                return 0
        except ValueError:
            return 0

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

    def calculate_temps_travail_real_for_month(self, cursor, project_id, year, month):
        """Calcule le coût du temps de travail RÉEL pour un mois spécifique (sans redistribution)"""
        # Noms des mois en français
        month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                      "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        
        if month < 1 or month > 12:
            return 0
            
        mois_francais = month_names[month - 1]
        
        # Récupérer le temps de travail réel pour ce mois précis
        cursor.execute("""
            SELECT tt.categorie, tt.jours
            FROM temps_travail tt
            WHERE tt.projet_id = ? AND tt.annee = ? AND tt.mois = ?
        """, (project_id, year, mois_francais))
        
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
            
            # Récupérer le coût pour cette catégorie et cette année
            cursor.execute("""
                SELECT montant_charge 
                FROM categorie_cout 
                WHERE categorie = ? AND annee = ?
            """, (categorie_code, year))
            
            cout_row = cursor.fetchone()
            if cout_row and cout_row[0]:
                montant_charge = float(cout_row[0])
                cout_total += jours * montant_charge
            else:
                # Valeur par défaut si pas de coût défini
                cout_total += jours * 500
        
        return cout_total

    def print_page(self):
        """Ouvre l'aperçu d'impression de la page des détails du projet"""
        try:
            from PyQt6.QtPrintSupport import QPrintPreviewDialog, QPrinter
            from PyQt6.QtGui import QTextDocument
            from PyQt6.QtCore import QMarginsF
            import datetime
            
            # Créer le document HTML
            html_content = self.generate_print_html()
            
            # Créer l'objet printer avec configuration minimale
            printer = QPrinter()
            
            # Configuration basique qui fonctionne partout
            try:
                printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            except:
                pass
            
            # Créer la dialog d'aperçu d'impression
            preview_dialog = QPrintPreviewDialog(printer, self)
            preview_dialog.setWindowTitle("Aperçu d'impression - Détails du projet")
            
            # Connecter le signal pour générer le contenu
            preview_dialog.paintRequested.connect(lambda: self.print_document(printer, html_content))
            
            # Redimensionner la fenêtre d'aperçu
            screen = preview_dialog.screen().geometry()
            preview_dialog.resize(int(screen.width() * 0.8), int(screen.height() * 0.8))
            
            # Afficher l'aperçu
            preview_dialog.exec()
            
        except ImportError as e:
            QMessageBox.critical(
                self, 
                "Erreur", 
                f"Impossible de charger les modules d'impression :\n{str(e)}\n\n"
                "Assurez-vous que PyQt6.QtPrintSupport est installé."
            )
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Erreur", 
                f"Erreur lors de l'ouverture de l'aperçu d'impression :\n{str(e)}"
            )

    def print_document(self, printer, html_content):
        """Imprime le document HTML sur l'imprimante/PDF"""
        from PyQt6.QtGui import QTextDocument
        
        document = QTextDocument()
        document.setHtml(html_content)
        document.print(printer)

    def generate_print_html(self):
        """Génère le contenu HTML formaté pour l'impression"""
        # Récupérer toutes les données du projet
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            
            # Informations du projet
            cursor.execute('''
                SELECT p.code, p.nom, p.details, p.date_debut, p.date_fin, p.livrables, 
                       c.nom || ' ' || c.prenom || ' - ' || c.direction AS chef_complet, 
                       p.etat, p.cir, p.subvention, p.theme_principal
                FROM projets p
                LEFT JOIN chefs_projet c ON p.chef = c.id
                WHERE p.id=?
            ''', (self.projet_id,))
            projet = cursor.fetchone()
            
            # Thèmes
            cursor.execute('SELECT t.nom FROM themes t JOIN projet_themes pt ON t.id=pt.theme_id WHERE pt.projet_id=?', (self.projet_id,))
            themes = [nom for (nom,) in cursor.fetchall()]
            
            # Équipe
            cursor.execute('SELECT direction, type, nombre FROM equipe WHERE projet_id=?', (self.projet_id,))
            equipe = cursor.fetchall()
            
            # Investissements
            cursor.execute('SELECT montant, date_achat, duree FROM investissements WHERE projet_id=?', (self.projet_id,))
            investissements = cursor.fetchall()
            
            # Actualités
            cursor.execute('SELECT message, date FROM actualites WHERE projet_id=? ORDER BY date DESC', (self.projet_id,))
            actualites = cursor.fetchall()
            
            # Images
            cursor.execute('SELECT nom, data FROM images WHERE projet_id=?', (self.projet_id,))
            images = cursor.fetchall()
            
            # Calcul des coûts (réutiliser la logique existante)
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
                    couts["charge"] += float(total_jours * montant_charge) if montant_charge else 0
                    couts["direct"] += float(total_jours * cout_production) if cout_production else 0
                    couts["complet"] += float(total_jours * cout_complet) if cout_complet else 0

            # Ajouter les dépenses externes et autres
            cursor.execute('SELECT SUM(montant) FROM depenses WHERE projet_id = ?', (self.projet_id,))
            depenses_externes = cursor.fetchone()
            if depenses_externes and depenses_externes[0]:
                montant_depenses = float(depenses_externes[0])
                couts["charge"] += montant_depenses
                couts["direct"] += montant_depenses
                couts["complet"] += montant_depenses

            cursor.execute('SELECT SUM(montant) FROM autres_depenses WHERE projet_id = ?', (self.projet_id,))
            autres_depenses = cursor.fetchone()
            if autres_depenses and autres_depenses[0]:
                montant_autres = float(autres_depenses[0])
                couts["charge"] += montant_autres
                couts["direct"] += montant_autres
                couts["complet"] += montant_autres

        # Génération du HTML
        from datetime import datetime
        from utils import format_montant
        import html as html_module
        
        # Fonction pour échapper les caractères HTML
        def escape_html(text):
            return html_module.escape(str(text)) if text else ""
        
        date_impression = datetime.now().strftime("%d/%m/%Y à %H:%M")
        
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Détails du projet - {projet[1] if projet[1] else 'Sans nom'}</title>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    font-size: 12pt;
                    line-height: 1.4;
                    margin: 20px;
                    padding: 0;
                    max-width: 100%;
                }}
                .header {{
                    text-align: center;
                    border-bottom: 2px solid #333;
                    padding-bottom: 10px;
                    margin-bottom: 20px;
                }}
                .project-title {{
                    font-size: 18pt;
                    font-weight: bold;
                    color: #333;
                    margin-bottom: 5px;
                }}
                .project-code {{
                    font-size: 14pt;
                    color: #666;
                    margin-bottom: 5px;
                }}
                .print-date {{
                    font-size: 9pt;
                    color: #888;
                    font-style: italic;
                }}
                .section {{
                    margin-bottom: 20px;
                    break-inside: avoid;
                }}
                .section-title {{
                    font-size: 14pt;
                    font-weight: bold;
                    color: #333;
                    border-bottom: 1px solid #ccc;
                    padding-bottom: 5px;
                    margin-bottom: 10px;
                }}
                .info-grid {{
                    display: table;
                    width: 100%;
                    margin-bottom: 15px;
                }}
                .info-grid > div {{
                    display: table-cell;
                    width: 50%;
                    vertical-align: top;
                    padding-right: 20px;
                }}
                .info-item {{
                    margin-bottom: 8px;
                }}
                .info-label {{
                    font-weight: bold;
                    color: #333;
                }}
                .details-box {{
                    background-color: #f9f9f9;
                    border: 1px solid #ddd;
                    padding: 10px;
                    border-radius: 4px;
                    margin-bottom: 10px;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 15px;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: left;
                }}
                th {{
                    background-color: #f2f2f2;
                    font-weight: bold;
                }}
                .budget-table {{
                    width: 50%;
                }}
                .budget-table td:last-child {{
                    text-align: right;
                    font-family: 'Courier New', monospace;
                }}
                .actualites-table {{
                    margin-top: 10px;
                }}
                .actualites-table td:first-child {{
                    width: 120px;
                    text-align: center;
                    font-size: 9pt;
                }}
                .footer {{
                    margin-top: 30px;
                    text-align: center;
                    font-size: 9pt;
                    color: #666;
                    border-top: 1px solid #ccc;
                    padding-top: 10px;
                }}
            </style>
        </head>
        <body>
            <div class="header">
                <div class="project-title">{escape_html(projet[1]) if projet[1] else 'Projet sans nom'}</div>
                <div class="project-code">Code: {escape_html(projet[0]) if projet[0] else 'Non défini'}</div>
                <div class="print-date">Imprimé le {date_impression}</div>
            </div>

            <div class="section">
                <div class="section-title">Informations générales</div>
                <div class="info-grid">
                    <div>
                        <div class="info-item">
                            <span class="info-label">Date de début:</span> {escape_html(projet[3]) if projet[3] else 'Non définie'}
                        </div>
                        <div class="info-item">
                            <span class="info-label">Date de fin:</span> {escape_html(projet[4]) if projet[4] else 'Non définie'}
                        </div>
                        <div class="info-item">
                            <span class="info-label">État:</span> {escape_html(projet[7]) if projet[7] else 'Non défini'}
                        </div>
                        <div class="info-item">
                            <span class="info-label">Chef de projet:</span> {escape_html(projet[6]) if projet[6] else 'Non défini'}
                        </div>
                    </div>
                    <div>
                        <div class="info-item">
                            <span class="info-label">CIR:</span> {'Oui' if projet[8] else 'Non'}
                        </div>
                        <div class="info-item">
                            <span class="info-label">Subvention:</span> {'Oui' if projet[9] else 'Non'}
                        </div>
                        {f"<div class='info-item'><span class='info-label'>Thème principal:</span> {escape_html(projet[10])}</div>" if projet[10] else ""}
                        {f"<div class='info-item'><span class='info-label'>Thèmes:</span> {escape_html(', '.join(themes))}</div>" if themes else ""}
                    </div>
                </div>
                
                {f'''<div class="details-box">
                    <div class="info-label">Détails du projet:</div>
                    <div>{escape_html(projet[2]) if projet[2] else "Aucun détail"}</div>
                </div>''' if projet[2] else ''}
                
                {f'''<div class="details-box">
                    <div class="info-label">Livrables:</div>
                    <div>{escape_html(projet[5]) if projet[5] else "Aucun livrable défini"}</div>
                </div>''' if projet[5] else ''}
            </div>
        """
        
        # Section Équipe
        if equipe:
            equipe_par_direction = {}
            for direction, type_, nombre in equipe:
                if nombre > 0:
                    if direction not in equipe_par_direction:
                        equipe_par_direction[direction] = []
                    equipe_par_direction[direction].append(f"{type_}: {nombre}")
            
            if equipe_par_direction:
                html += '''
                <div class="section">
                    <div class="section-title">Équipe</div>
                    <table>
                        <thead>
                            <tr>
                                <th>Direction</th>
                                <th>Composition</th>
                            </tr>
                        </thead>
                        <tbody>
                '''
                for direction, membres in equipe_par_direction.items():
                    html += f'''
                            <tr>
                                <td>{escape_html(direction)}</td>
                                <td>{escape_html(", ".join(membres))}</td>
                            </tr>
                    '''
                html += '''
                        </tbody>
                    </table>
                </div>
                '''

        # Section Investissements
        if investissements:
            html += '''
            <div class="section">
                <div class="section-title">Investissements</div>
                <table>
                    <thead>
                        <tr>
                            <th>Montant</th>
                            <th>Date d'achat</th>
                            <th>Durée (années)</th>
                        </tr>
                    </thead>
                    <tbody>
            '''
            for montant, date_achat, duree in investissements:
                html += f'''
                        <tr>
                            <td>{format_montant(montant)}</td>
                            <td>{escape_html(date_achat)}</td>
                            <td>{duree}</td>
                        </tr>
                '''
            html += '''
                    </tbody>
                </table>
            </div>
            '''

        # Section Budget
        html += f'''
        <div class="section">
            <div class="section-title">Budget</div>
            <table class="budget-table">
                <thead>
                    <tr>
                        <th>Type de coût</th>
                        <th>Montant</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>Coût chargé</td>
                        <td>{format_montant(couts["charge"])}</td>
                    </tr>
                    <tr>
                        <td>Coût production</td>
                        <td>{format_montant(couts["direct"])}</td>
                    </tr>
                    <tr>
                        <td>Coût complet</td>
                        <td>{format_montant(couts["complet"])}</td>
                    </tr>
                </tbody>
            </table>
        </div>
        '''

        # Section Subventions
        if projet[9]:  # Si le projet a des subventions
            cursor.execute('''
                SELECT nom, mode_simplifie, montant_forfaitaire, depenses_temps_travail, coef_temps_travail, 
                       depenses_externes, coef_externes, depenses_autres_achats, coef_autres_achats, 
                       depenses_dotation_amortissements, coef_dotation_amortissements, cd, taux,
                       date_debut_subvention, date_fin_subvention, montant_subvention_max, depenses_eligibles_max
                FROM subventions 
                WHERE projet_id = ?
            ''', (self.projet_id,))
            subventions_data = cursor.fetchall()
            
            if subventions_data:
                html += '''
                <div class="section">
                    <div class="section-title">Subventions</div>
                    <table>
                        <thead>
                            <tr>
                                <th>Nom</th>
                                <th>Type</th>
                                <th>Période</th>
                                <th>Taux</th>
                                <th>Montant max</th>
                                <th>Dépenses éligibles max</th>
                            </tr>
                        </thead>
                        <tbody>
                '''
                
                for subv in subventions_data:
                    nom, mode_simplifie, montant_forfaitaire, dep_temps, coef_temps, dep_ext, coef_ext, dep_autres, coef_autres, dep_amort, coef_amort, cd, taux, date_debut, date_fin, montant_max, depenses_max = subv
                    
                    type_subv = "Forfaitaire" if mode_simplifie else "Au réel"
                    periode = f"{date_debut} - {date_fin}" if date_debut and date_fin else "Non définie"
                    taux_display = f"{taux}%" if taux else "Non défini"
                    montant_max_display = format_montant(montant_max) if montant_max else "Non défini"
                    depenses_max_display = format_montant(depenses_max) if depenses_max else "Non défini"
                    
                    html += f'''
                            <tr>
                                <td>{escape_html(nom)}</td>
                                <td>{type_subv}</td>
                                <td>{periode}</td>
                                <td>{taux_display}</td>
                                <td>{montant_max_display}</td>
                                <td>{depenses_max_display}</td>
                            </tr>
                    '''
                
                html += '''
                        </tbody>
                    </table>
                </div>
                '''

        # Section CIR
        if projet[8]:  # Si le projet a le CIR activé
            # Récupérer les coefficients CIR
            cursor.execute('SELECT date_debut, date_fin FROM projets WHERE id = ?', (self.projet_id,))
            date_row = cursor.fetchone()
            
            if date_row and date_row[0] and date_row[1]:
                try:
                    import datetime
                    debut_projet = datetime.datetime.strptime(date_row[0], '%m/%Y')
                    fin_projet = datetime.datetime.strptime(date_row[1], '%m/%Y')
                    
                    # Récupérer les coefficients CIR pour la première année
                    cursor.execute('SELECT k1, k2, k3 FROM cir_coeffs WHERE annee = ?', (debut_projet.year,))
                    cir_coeffs = cursor.fetchone()
                    
                    if cir_coeffs and cir_coeffs[2]:  # k3 existe
                        k3 = cir_coeffs[2]
                        taux_cir = k3 * 100
                        
                        html += f'''
                        <div class="section">
                            <div class="section-title">Crédit d'Impôt Recherche (CIR)</div>
                            <table class="budget-table">
                                <thead>
                                    <tr>
                                        <th>Information</th>
                                        <th>Valeur</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td>Statut</td>
                                        <td>Activé</td>
                                    </tr>
                                    <tr>
                                        <td>Taux applicable</td>
                                        <td>{taux_cir:.1f}%</td>
                                    </tr>
                                    <tr>
                                        <td>Période d'application</td>
                                        <td>{date_row[0]} - {date_row[1]}</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        '''
                except:
                    html += '''
                    <div class="section">
                        <div class="section-title">Crédit d'Impôt Recherche (CIR)</div>
                        <p>CIR activé - Informations détaillées non disponibles</p>
                    </div>
                    '''

        # Section Images
        if images:
            html += '''
            <div class="section">
                <div class="section-title">Images du projet</div>
                <div style="text-align: center;">
            '''
            
            for nom, data in images:
                try:
                    # Convertir les données binaires en base64 pour l'inclusion dans le HTML
                    import base64
                    image_base64 = base64.b64encode(data).decode('utf-8')
                    
                    # Détecter le type d'image (simple détection basée sur les premiers bytes)
                    if data.startswith(b'\xFF\xD8\xFF'):
                        mime_type = 'image/jpeg'
                    elif data.startswith(b'\x89PNG'):
                        mime_type = 'image/png'
                    elif data.startswith(b'GIF'):
                        mime_type = 'image/gif'
                    else:
                        mime_type = 'image/jpeg'  # Par défaut
                    
                    html += f'''
                        <div style="margin: 10px; display: inline-block;">
                            <p style="font-weight: bold; margin-bottom: 5px;">{escape_html(nom)}</p>
                            <img src="data:{mime_type};base64,{image_base64}" 
                                 style="max-width: 300px; max-height: 200px; border: 1px solid #ddd; border-radius: 4px;" 
                                 alt="{escape_html(nom)}" />
                        </div>
                    '''
                except Exception as e:
                    # En cas d'erreur, afficher juste le nom de l'image
                    html += f'''
                        <div style="margin: 10px; display: inline-block;">
                            <p style="font-weight: bold;">{escape_html(nom)}</p>
                            <p style="color: #666; font-style: italic;">Image non disponible</p>
                        </div>
                    '''
            
            html += '''
                </div>
            </div>
            '''

        # Section Actualités
        if actualites:
            html += '''
            <div class="section">
                <div class="section-title">Actualités du projet</div>
                <table class="actualites-table">
                    <thead>
                        <tr>
                            <th>Date</th>
                            <th>Message</th>
                        </tr>
                    </thead>
                    <tbody>
            '''
            for message, date in actualites:
                html += f'''
                        <tr>
                            <td>{escape_html(date)}</td>
                            <td>{escape_html(message)}</td>
                        </tr>
                '''
            html += '''
                    </tbody>
                </table>
            </div>
            '''

        html += '''
            <div class="footer">
                Document généré automatiquement par le système de gestion de budget
            </div>
        </body>
        </html>
        '''
        
        return html

