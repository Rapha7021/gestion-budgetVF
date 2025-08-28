from PyQt6.QtWidgets import QDialog, QVBoxLayout, QGridLayout, QLabel, QHBoxLayout, QListWidget, QListWidgetItem, QPushButton, QInputDialog, QMessageBox, QTextEdit, QDialogButtonBox
from PyQt6.QtCore import Qt
import sqlite3
import os
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
                    
        # Affichage des coûts
        budget_vbox = QVBoxLayout()
        budget_vbox.addWidget(QLabel(f"<b>Budget Total :</b>"))
        budget_vbox.addWidget(QLabel(f"Coût chargé : {couts['charge']:.2f} €"))
        budget_vbox.addWidget(QLabel(f"Coût direct : {couts['direct']:.2f} €"))
        budget_vbox.addWidget(QLabel(f"Coût complet : {couts['complet']:.2f} €"))
        grid.addLayout(budget_vbox, 0, 2)
        
        # En haut à gauche
        left_vbox = QVBoxLayout()
        left_vbox.addWidget(QLabel(f"<b>Code projet :</b> {projet[0]}"))
        h_nom = QHBoxLayout()
        h_nom.addWidget(QLabel(f"<b>Nom projet :</b> {projet[1]}"))
        left_vbox.addLayout(h_nom)
        left_vbox.addWidget(QLabel(f"<b>Détails :</b> {projet[2]}"))
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
                invest_text += f"- {montant} € | Achat: {date_achat} | Durée: {duree} ans\n"
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
                pixmap.loadFromData(data)
                img_widget = QLabel()
                img_widget.setPixmap(pixmap.scaled(150, 150, Qt.AspectRatioMode.KeepAspectRatio))
                img_hbox.addWidget(img_widget)
            except Exception:
                pass
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

        # Espace vide en dessous
        main_layout.addStretch()
        self.setLayout(main_layout)

    def handle_import_excel(self):
        from import_manager_dialog import ImportManagerDialog
        dlg = ImportManagerDialog(self, self.projet_id)
        dlg.exec()

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

    def open_task_manager(self):
        from task_manager_dialog import TaskManagerDialog
        dlg = TaskManagerDialog(self, self.projet_id)
        dlg.exec()

