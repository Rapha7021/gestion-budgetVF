from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QListWidget, QPushButton, QHBoxLayout,
    QMessageBox, QInputDialog,
    QDialog, QLabel, QLineEdit, QTextEdit, QDateEdit, QComboBox, QCheckBox, QSpinBox,
    QFileDialog, QFormLayout, QGroupBox, QMenu, QGridLayout, QScrollArea, QTableWidget, QTableWidgetItem
)
from PyQt6.QtCore import Qt
import sqlite3
import sys
import datetime
import re
import shutil
import os
import pandas as pd  # Ajout pour lecture Excel

DB_PATH = 'gestion_budget.db'

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS themes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nom TEXT UNIQUE NOT NULL
    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS projets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT NOT NULL,
        nom TEXT NOT NULL,
        details TEXT,
        date_debut TEXT,
        date_fin TEXT,
        livrables TEXT,
        chef TEXT,
        etat TEXT,
        cir INTEGER,
        subvention INTEGER,
        theme_principal TEXT
    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS projet_themes (
        projet_id INTEGER,
        theme_id INTEGER,
        FOREIGN KEY(projet_id) REFERENCES projets(id),
        FOREIGN KEY(theme_id) REFERENCES themes(id),
        PRIMARY KEY (projet_id, theme_id)
    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS images (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        projet_id INTEGER,
        nom TEXT,
        data BLOB,
        FOREIGN KEY(projet_id) REFERENCES projets(id)
    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS investissements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        projet_id INTEGER,
        nom TEXT,
        montant REAL,
        date_achat TEXT,
        duree INTEGER,
        FOREIGN KEY(projet_id) REFERENCES projets(id)
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS subventions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        projet_id INTEGER,
        nom TEXT,
        depenses_temps_travail INTEGER,
        coef_temps_travail REAL,
        depenses_externes INTEGER,
        coef_externes REAL,
        depenses_autres_achats INTEGER,
        coef_autres_achats REAL,
        depenses_dotation_amortissements INTEGER,
        coef_dotation_amortissements REAL,
        cd REAL,
        taux REAL,
        FOREIGN KEY(projet_id) REFERENCES projets(id)
    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS equipe (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        projet_id INTEGER,
        type TEXT,
        nombre INTEGER,
        direction TEXT,
        FOREIGN KEY(projet_id) REFERENCES projets(id)
    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS actualites (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        projet_id INTEGER,
        message TEXT NOT NULL,
        date TEXT NOT NULL,
        FOREIGN KEY(projet_id) REFERENCES projets(id)
    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS directions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nom TEXT UNIQUE NOT NULL
    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS chefs_projet (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nom TEXT NOT NULL,
        prenom TEXT NOT NULL,
        direction TEXT NOT NULL,
        FOREIGN KEY(direction) REFERENCES directions(nom)
    )''')
    conn.commit()
    conn.close()

class MainWindow(QWidget):
    def edit_project(self):
        rows = self.project_table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.warning(self, 'Modifier', 'Sélectionnez un projet à modifier.')
            return
        row = rows[0].row()
        code = self.project_table.item(row, 0).text()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM projets WHERE code=?', (code,))
        res = cursor.fetchone()
        conn.close()
        if not res:
            QMessageBox.warning(self, 'Erreur', 'Projet introuvable.')
            return
        projet_id = res[0]
        form = ProjectForm(self, projet_id)
        if form.exec():
            self.load_projects()

    def delete_project(self):
        rows = self.project_table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.warning(self, 'Supprimer', 'Sélectionnez un projet à supprimer.')
            return
        row = rows[0].row()
        code = self.project_table.item(row, 0).text()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM projets WHERE code=?', (code,))
        res = cursor.fetchone()
        if not res:
            conn.close()
            QMessageBox.warning(self, 'Erreur', 'Projet introuvable.')
            return
        pid = res[0]
        confirm = QMessageBox.question(
            self, 'Confirmation', f'Voulez-vous vraiment supprimer le projet {code} ?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm == QMessageBox.StandardButton.Yes:
            cursor.execute('DELETE FROM projets WHERE id=?', (pid,))
            conn.commit()
            conn.close()
            self.load_projects()
        else:
            conn.close()
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Gestion de Budget - Menu Principal')
        screen = QApplication.primaryScreen().geometry()
        self.setGeometry(
            int(screen.x() + screen.width() * 0.05),
            int(screen.y() + screen.height() * 0.05),
            int(screen.width() * 0.9),
            int(screen.height() * 0.85)
        )
        init_db()
        self.setup_ui()
        self.load_projects()

    def setup_ui(self):
        layout = QVBoxLayout()
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem
        self.project_table = QTableWidget()
        self.project_table.setColumnCount(4)
        self.project_table.setHorizontalHeaderLabels(['Code projet', 'Nom projet', 'Chef de projet', 'Etat'])
        self.project_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.project_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        layout.addWidget(self.project_table)
        btn_layout = QHBoxLayout()
        self.btn_new = QPushButton('Nouveau projet')
        self.btn_edit = QPushButton('Modifier le projet sélectionné')
        self.btn_delete = QPushButton('Supprimer')
        self.btn_themes = QPushButton('Gérer les thèmes')
        self.btn_couts_categorie = QPushButton('Coûts par catégorie')
        self.btn_cir = QPushButton('CIR')
        self.btn_directions = QPushButton('Gérer les directions')
        self.btn_project_managers = QPushButton('Gérer les chefs de projet')
        self.btn_couts_categorie.setToolTip(
            "Source :\nMagic S\nRevue de projet\nHypothèse LLH"
        )
        btn_layout.addWidget(self.btn_new)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addWidget(self.btn_themes)
        btn_layout.addWidget(self.btn_couts_categorie, alignment=Qt.AlignmentFlag.AlignRight)
        btn_layout.addWidget(self.btn_cir, alignment=Qt.AlignmentFlag.AlignRight)
        btn_layout.addWidget(self.btn_directions)
        btn_layout.addWidget(self.btn_project_managers)
        layout.addLayout(btn_layout)
        self.setLayout(layout)
        self.btn_new.clicked.connect(self.open_project_form)
        self.btn_edit.clicked.connect(self.edit_project)
        self.btn_delete.clicked.connect(self.delete_project)
        self.btn_themes.clicked.connect(self.open_theme_manager)
        self.project_table.cellDoubleClicked.connect(self.show_project_details)
        self.btn_couts_categorie.clicked.connect(self.open_categorie_cout_dialog)
        self.btn_cir.clicked.connect(self.open_cir_dialog)
        self.btn_directions.clicked.connect(self.open_direction_manager)
        self.btn_project_managers.clicked.connect(self.open_project_manager_dialog)

    def open_categorie_cout_dialog(self):
        from categorie_cout_dialog import CategorieCoutDialog
        dialog = CategorieCoutDialog(self)
        dialog.exec()

    def load_projects(self):
        self.project_table.setRowCount(0)
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT id, code, nom, chef, etat FROM projets ORDER BY id DESC')
        for row_idx, (pid, code, nom, chef, etat) in enumerate(cursor.fetchall()):
            self.project_table.insertRow(row_idx)
            self.project_table.setItem(row_idx, 0, QTableWidgetItem(str(code)))
            self.project_table.setItem(row_idx, 1, QTableWidgetItem(str(nom)))
            self.project_table.setItem(row_idx, 2, QTableWidgetItem(str(chef)))
            self.project_table.setItem(row_idx, 3, QTableWidgetItem(str(etat)))
        conn.close()

    def open_project_form(self):
        form = ProjectForm(self)
        if form.exec():
            self.load_projects()

    def open_theme_manager(self):
        dialog = ThemeManager(self)
        dialog.exec()

    def open_direction_manager(self):
        dialog = DirectionManager(self)
        dialog.exec()

    def open_cir_dialog(self):
        from cir_dialog import CIRDialog
        dialog = CIRDialog(self)
        dialog.exec()

    def open_project_manager_dialog(self):
        from project_manager_dialog import ProjectManagerDialog
        dialog = ProjectManagerDialog()
        dialog.exec()

    def show_project_details(self, row, column):
        code = self.project_table.item(row, 0).text()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM projets WHERE code=?', (code,))
        res = cursor.fetchone()
        conn.close()
        if not res:
            QMessageBox.warning(self, 'Erreur', 'Projet introuvable.')
            return
        projet_id = res[0]
        # Correction : import ici
        from project_details_dialog import ProjectDetailsDialog
        dialog = ProjectDetailsDialog(self, projet_id)
        dialog.exec()

class ProjectForm(QDialog):
    def __init__(self, parent=None, projet_id=None):
        super().__init__(parent)
        self.projet_id = projet_id
        self.setWindowTitle('Créer un projet')
        self.setMinimumWidth(1200)
        self.layout = QVBoxLayout()
        grid = QGridLayout()
        row = 0
        # Code projet et Nom projet côte à côte
        grid.addWidget(QLabel('Code projet:'), row, 0)
        self.code_edit = QLineEdit()
        grid.addWidget(self.code_edit, row, 1)
        grid.addWidget(QLabel('Nom projet:'), row, 2)
        self.nom_edit = QLineEdit()
        grid.addWidget(self.nom_edit, row, 3)
        row += 1
        # Dates côte à côte
        grid.addWidget(QLabel('Début:'), row, 0)
        self.date_debut = QDateEdit()
        self.date_debut.setDisplayFormat('MM/yyyy')
        self.date_debut.setDate(datetime.date.today())
        grid.addWidget(self.date_debut, row, 1)
        grid.addWidget(QLabel('Fin:'), row, 2)
        self.date_fin = QDateEdit()
        self.date_fin.setDisplayFormat('MM/yyyy')
        self.date_fin.setDate(datetime.date.today())
        grid.addWidget(self.date_fin, row, 3)
        row += 1
        # Livrables et Chef de projet côte à côte
        grid.addWidget(QLabel('Livrables principaux:'), row, 0)
        self.livrables_edit = QLineEdit()
        grid.addWidget(self.livrables_edit, row, 1)
        grid.addWidget(QLabel('Nom chef de projet:'), row, 2)
        self.chef_edit = QLineEdit()
        grid.addWidget(self.chef_edit, row, 3)
        row += 1
        # Etat projet
        grid.addWidget(QLabel('Etat projet:'), row, 0)
        self.etat_combo = QComboBox()
        self.etat_combo.addItems(['Terminé', 'En cours', 'Futur'])
        grid.addWidget(self.etat_combo, row, 1)
        row += 1
        # Détails projet (sur toute la largeur) - version compacte
        grid.addWidget(QLabel('Détails projet:'), row, 0)
        self.details_edit = QTextEdit()
        self.details_edit.setMaximumHeight(60)  # Hauteur limitée
        grid.addWidget(self.details_edit, row, 1, 1, 3)
        row += 1
        # Thèmes (recherche + tags) - version très compacte
        theme_group = QGroupBox('Thèmes')
        theme_vbox = QVBoxLayout()
        self.theme_search = QLineEdit()
        self.theme_search.setPlaceholderText('Rechercher un thème...')
        self.theme_listwidget = QListWidget()
        self.theme_listwidget.setMaximumHeight(50)  # Encore plus petit
        theme_vbox.addWidget(self.theme_search)
        theme_vbox.addWidget(self.theme_listwidget)
        self.tag_area = QScrollArea()
        self.tag_area.setWidgetResizable(True)
        self.tag_area.setMaximumHeight(30)  # Encore plus petit pour les tags
        self.tag_widget = QWidget()
        self.tag_layout = QHBoxLayout()
        self.tag_layout.setContentsMargins(0,0,0,0)
        self.tag_widget.setLayout(self.tag_layout)
        self.tag_area.setWidget(self.tag_widget)
        theme_vbox.addWidget(self.tag_area)
        
        # Ajout du QComboBox pour le thème principal dans le groupe thèmes
        theme_principal_layout = QHBoxLayout()
        theme_principal_layout.addWidget(QLabel('Thème principal:'))
        self.theme_principal_combo = QComboBox()
        theme_principal_layout.addWidget(self.theme_principal_combo)
        theme_principal_layout.addStretch()  # Pour aligner à gauche
        theme_vbox.addLayout(theme_principal_layout)
        
        theme_group.setLayout(theme_vbox)
        grid.addWidget(theme_group, row, 0, 1, 2)  # Réduit à 1 rangée au lieu de 2
        self.selected_themes = []
        self.theme_search.textChanged.connect(self.filter_themes)
        self.theme_listwidget.itemClicked.connect(self.add_theme_tag)
        self.load_themes()
        # Images (groupe à part) - version très compacte
        img_group = QGroupBox('Images')
        img_vbox = QVBoxLayout()
        self.btn_add_image = QPushButton('Ajouter image')
        self.btn_add_image.clicked.connect(self.add_image)
        img_vbox.addWidget(self.btn_add_image)
        self.images_list = QListWidget()
        self.images_list.setMaximumHeight(50)  # Encore plus petit
        img_vbox.addWidget(self.images_list)
        img_group.setLayout(img_vbox)
        grid.addWidget(img_group, row, 2, 1, 2)  # Réduit à 1 rangée au lieu de 2
        # Ajout du menu contextuel pour suppression d'image
        self.images_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.images_list.customContextMenuRequested.connect(self.image_context_menu)
        row += 1  # Incrément réduit
        # Investissements et Subventions empilés à gauche, Équipe à droite
        invest_group = QGroupBox('Investissements')
        invest_vbox = QVBoxLayout()
        self.btn_add_invest = QPushButton('Ajouter investissement')
        self.btn_add_invest.clicked.connect(self.add_invest)
        invest_vbox.addWidget(self.btn_add_invest)
        self.invest_list = QListWidget()
        self.invest_list.setMaximumHeight(40)  # Encore plus petit
        self.invest_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.invest_list.customContextMenuRequested.connect(self.invest_context_menu)
        self.invest_list.itemDoubleClicked.connect(self.edit_invest)
        invest_vbox.addWidget(self.invest_list)
        invest_group.setLayout(invest_vbox)
        grid.addWidget(invest_group, row, 0, 1, 2)  # Colonne gauche
        
        # Subventions juste en dessous des investissements
        subv_group = QGroupBox('Subventions')
        subv_vbox = QVBoxLayout()
        self.btn_add_subv = QPushButton('Ajouter subvention')
        self.btn_add_subv.clicked.connect(self.add_subvention)
        subv_vbox.addWidget(self.btn_add_subv)
        self.subv_list = QListWidget()
        self.subv_list.setMaximumHeight(40)  # Encore plus petit
        self.subv_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.subv_list.customContextMenuRequested.connect(self.subv_context_menu)
        self.subv_list.itemDoubleClicked.connect(self.edit_subvention)
        subv_vbox.addWidget(self.subv_list)
        subv_group.setLayout(subv_vbox)
        grid.addWidget(subv_group, row + 1, 0, 1, 2)  # Colonne gauche, ligne suivante
        
        # Checkbox CIR juste en dessous des subventions
        self.cir_check = QCheckBox('Ajouter un CIR')
        grid.addWidget(self.cir_check, row + 2, 0, 1, 2)  # En dessous des subventions
        
        # Équipe par direction (côté droit, sur 2 rangées)
        equipe_group = QGroupBox('Équipe par direction')
        equipe_vbox = QVBoxLayout()
        self.direction_combo = QComboBox()
        self.directions = []
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT nom FROM directions ORDER BY nom')
        self.directions = [nom for (nom,) in cursor.fetchall()]
        conn.close()
        self.direction_combo.addItems(self.directions)
        equipe_vbox.addWidget(QLabel('Direction :'))
        equipe_vbox.addWidget(self.direction_combo)
        self.equipe_types_labels = [
            'Stagiaire Projet',
            'Assistante / opérateur',
            'Technicien',
            'Junior',
            'Senior',
            'Expert',
            'Collaborateur moyen'
        ]
        self.equipe_spins = {}
        equipe_form = QFormLayout()
        for label in self.equipe_types_labels:
            spin = QSpinBox()
            spin.setRange(0, 99)
            self.equipe_spins[label] = spin
            equipe_form.addRow(label, spin)
        equipe_vbox.addLayout(equipe_form)
        equipe_group.setLayout(equipe_vbox)
        grid.addWidget(equipe_group, row, 2, 3, 2)  # Côté droit, sur 3 rangées
        row += 3
        row += 1
            # --- Ajout : gestion des effectifs par direction ---
        self.equipe_data = {dir_: {label: 0 for label in self.equipe_types_labels} for dir_ in self.directions}
        self.direction_combo.currentTextChanged.connect(self.on_direction_changed)
        
        # Connexion des spins avec une fonction spécifique pour chaque label
        def make_callback(label_name):
            return lambda value: self.on_equipe_spin_changed(label_name, value)
            
        for label, spin in self.equipe_spins.items():
            callback = make_callback(label)
            spin.valueChanged.connect(callback)
            
        self._current_direction = self.direction_combo.currentText()  # <-- Ajout ici
        # Pas besoin d'appeler on_direction_changed explicitement ici
        
        # Initialisation des données de subventions
        self.subventions_data = []
        
        self.layout.insertLayout(0, grid)
         # Boutons valider/annuler
        btns = QHBoxLayout()
        self.btn_ok = QPushButton('Valider')
        self.btn_cancel = QPushButton('Annuler')
        self.btn_ok.setEnabled(False)
        self.btn_ok.clicked.connect(self.save_project)
        self.btn_cancel.clicked.connect(self.reject)
        btns.addWidget(self.btn_ok)
        btns.addWidget(self.btn_cancel)
        self.layout.addLayout(btns)
        self.setLayout(self.layout)
        # Contrôles de saisie
        self.code_edit.textChanged.connect(self.check_form_valid)
        self.nom_edit.textChanged.connect(self.check_form_valid)
        self.details_edit.textChanged.connect(self.check_form_valid)
        self.livrables_edit.textChanged.connect(self.check_form_valid)
        self.chef_edit.textChanged.connect(self.check_form_valid)
        self.theme_listwidget.itemSelectionChanged.connect(self.check_form_valid)
        self.check_form_valid()
        # Charger les données du projet si modification
        if self.projet_id:
            self.load_project_data()
        # Bouton Import Excel
        self.btn_import_excel = QPushButton('Importer Excel')
        self.btn_import_excel.setToolTip(
            "Source Agresso\n"
            "Requête\n"
            "1-Gestion de projet\n"
            "4-Requetes Edition\n"
            "6-Liste des opérations\n"
            "Liste des opérations actives\n"
            "Choix de la date de début\n"
            "Menu Editions\n"
            "Format XLSX"
        )
        self.btn_import_excel.clicked.connect(self.import_excel_dialog)
        self.layout.insertWidget(0, self.btn_import_excel)
        
        # Mettre à jour le combo box avec les thèmes déjà sélectionnés
        self.update_theme_tags()

    def on_direction_changed(self, direction):
        # Sauvegarde les valeurs courantes dans la direction précédente
        if hasattr(self, '_current_direction') and self._current_direction is not None:
            # Sauvegarde directement toutes les valeurs actuelles
            for label in self.equipe_types_labels:
                self.equipe_data[self._current_direction][label] = self.equipe_spins[label].value()
                
        # Charge les valeurs pour la direction sélectionnée
        for label in self.equipe_types_labels:
            # Bloquer temporairement les signaux pour éviter des appels récursifs
            self.equipe_spins[label].blockSignals(True)
            self.equipe_spins[label].setValue(self.equipe_data[direction][label])
            self.equipe_spins[label].blockSignals(False)
            
        self._current_direction = direction

    def on_equipe_spin_changed(self, label, value):
        # Met à jour le dictionnaire pour la direction courante et le label modifié
        if hasattr(self, '_current_direction') and self._current_direction is not None:
            # Débogage pour vérifier les valeurs
            self.equipe_data[self._current_direction][label] = value

    def import_excel_dialog(self):
        # Importer la fonction depuis le nouveau module
        from excel_importer_projet import import_project_from_excel
        import_project_from_excel(self)

    def check_form_valid(self):
        code_ok = bool(self.code_edit.text().strip())
        nom_ok = bool(self.nom_edit.text().strip())
        debut_ok = bool(self.date_debut.text().strip())
        fin_ok = bool(self.date_fin.text().strip())
        self.btn_ok.setEnabled(code_ok and nom_ok and debut_ok and fin_ok)

    def invest_context_menu(self, pos):
        item = self.invest_list.itemAt(pos)
        if item:
            menu = QMenu()
            edit_action = menu.addAction('Modifier')
            delete_action = menu.addAction('Supprimer')
            action = menu.exec(self.invest_list.mapToGlobal(pos))
            if action == edit_action:
                self.edit_invest(item)
            elif action == delete_action:
                confirm = QMessageBox.question(
                    self,
                    'Confirmation',
                    'Voulez-vous vraiment supprimer cet investissement ?',
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if confirm == QMessageBox.StandardButton.Yes:
                    self.invest_list.takeItem(self.invest_list.row(item))
        # Si clic droit sur une zone vide, rien ne se passe

    def edit_invest(self, item):
        # Parse l'investissement existant
        txt = item.text()
        try:
            if 'Nom: ' in txt:
                nom = txt.split('Nom: ')[1].split(',')[0].strip()
                montant = txt.split('Montant: ')[1].split(' €')[0]
                date_achat = txt.split('Date achat: ')[1].split(',')[0]
                duree = txt.split('Durée amort.: ')[1].split(' ans')[0]
            else:
                # Format ancien sans nom
                nom = ""
                montant = txt.split('Montant: ')[1].split(' €')[0]
                date_achat = txt.split('Date achat: ')[1].split(',')[0]
                duree = txt.split('Durée amort.: ')[1].split(' ans')[0]
        except Exception:
            nom, montant, date_achat, duree = '', '', '', ''
        dialog = InvestDialog(self, nom, montant, date_achat, duree)
        if dialog.exec():
            invest_str = f"Nom: {dialog.nom_edit.text()}, Montant: {dialog.montant_edit.text()} €, Date achat: {dialog.date_achat.text()}, Durée amort.: {dialog.duree_edit.text()} ans"
            item.setText(invest_str)

    def add_theme_tag(self, item):
        nom = item.text()
        if nom not in self.selected_themes:
            self.selected_themes.append(nom)
            self.update_theme_tags()
            self.theme_search.clear()
            self.filter_themes("")

    def update_theme_tags(self):
        # Supprime tous les widgets existants
        while self.tag_layout.count():
            child = self.tag_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        # Ajoute un tag pour chaque thème sélectionné
        for nom in self.selected_themes:
            tag = QWidget()
            tag_layout = QHBoxLayout()
            tag_layout.setContentsMargins(5, 2, 5, 2)
            lbl = QLabel(nom)
            btn = QPushButton('✕')
            btn.setFixedSize(20, 20)
            btn.setStyleSheet('QPushButton { border: none; background: #eee; }')
            btn.clicked.connect(lambda _, n=nom: self.remove_theme_tag(n))
            tag_layout.addWidget(lbl)
            tag_layout.addWidget(btn)
            tag.setLayout(tag_layout)
            self.tag_layout.addWidget(tag)
        self.tag_layout.addStretch()

        # Met à jour le QComboBox pour le thème principal (si il existe)
        if hasattr(self, 'theme_principal_combo'):
            self.theme_principal_combo.clear()
            self.theme_principal_combo.addItems(self.selected_themes)

    def remove_theme_tag(self, nom):
        if nom in self.selected_themes:
            self.selected_themes.remove(nom)
            self.update_theme_tags()
            self.filter_themes(self.theme_search.text())

    def load_themes(self):
        self.theme_listwidget.clear()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT nom FROM themes ORDER BY nom')
        self.all_themes = [nom for (nom,) in cursor.fetchall()]
        conn.close()
        for nom in self.all_themes:
            self.theme_listwidget.addItem(nom)
        self.update_theme_tags()

    def filter_themes(self, text):
        self.theme_listwidget.clear()
        for nom in self.all_themes:
            if text.lower() in nom.lower() and nom not in self.selected_themes:
                self.theme_listwidget.addItem(nom)

    def add_image(self):
        files, _ = QFileDialog.getOpenFileNames(self, 'Sélectionner des images', '', 'Images (*.png *.jpg *.jpeg *.bmp *.gif)')
        images_dir = os.path.join(os.path.dirname(__file__), 'images')
        for f in files:
            if os.path.isfile(f):
                dest = os.path.join(images_dir, os.path.basename(f))
                if not os.path.exists(dest):
                    shutil.copy2(f, dest)
                self.images_list.addItem(os.path.relpath(dest, os.path.dirname(__file__)))

    def add_invest(self):
        dialog = InvestDialog(self)
        if dialog.exec():
            invest_str = f"Nom: {dialog.nom_edit.text()}, Montant: {dialog.montant_edit.text()} €, Date achat: {dialog.date_achat.text()}, Durée amort.: {dialog.duree_edit.text()} ans"
            self.invest_list.addItem(invest_str)

    def add_subvention(self):
        from subvention_dialog import SubventionDialog
        dialog = SubventionDialog(self)
        # Passer l'ID du projet au dialogue
        dialog.projet_id = self.projet_id
        if dialog.exec():
            data = dialog.get_data()
            self.subventions_data.append(data)
            cats = []
            if data['depenses_temps_travail']: cats.append(f"Temps travail (coef {data['coef_temps_travail']})")
            if data['depenses_externes']: cats.append(f"Externes (coef {data['coef_externes']})")
            if data['depenses_autres_achats']: cats.append(f"Autres achats (coef {data['coef_autres_achats']})")
            if data['depenses_dotation_amortissements']: cats.append(f"Dotation (coef {data['coef_dotation_amortissements']})")
            subv_str = f"{data['nom']} | {', '.join(cats)} | Cd: {data['cd']} | Taux: {data['taux']}%"
            self.subv_list.addItem(subv_str)

    def subv_context_menu(self, pos):
        item = self.subv_list.itemAt(pos)
        if item:
            menu = QMenu()
            edit_action = menu.addAction('Modifier')
            delete_action = menu.addAction('Supprimer')
            action = menu.exec(self.subv_list.mapToGlobal(pos))
            if action == edit_action:
                self.edit_subvention(item)
            elif action == delete_action:
                confirm = QMessageBox.question(
                    self,
                    'Confirmation',
                    'Voulez-vous vraiment supprimer cette subvention ?',
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if confirm == QMessageBox.StandardButton.Yes:
                    row = self.subv_list.row(item)
                    self.subv_list.takeItem(row)
                    if row < len(self.subventions_data):
                        self.subventions_data.pop(row)

    def edit_subvention(self, item):
        row = self.subv_list.row(item)
        if row >= len(self.subventions_data):
            return
        from subvention_dialog import SubventionDialog
        existing_data = self.subventions_data[row]
        dialog = SubventionDialog(self, existing_data)
        # Passer l'ID du projet au dialogue
        dialog.projet_id = self.projet_id
        if dialog.exec():
            data = dialog.get_data()
            self.subventions_data[row] = data
            cats = []
            if data['depenses_temps_travail']: cats.append(f"Temps travail (coef {data['coef_temps_travail']})")
            if data['depenses_externes']: cats.append(f"Externes (coef {data['coef_externes']})")
            if data['depenses_autres_achats']: cats.append(f"Autres achats (coef {data['coef_autres_achats']})")
            if data['depenses_dotation_amortissements']: cats.append(f"Dotation (coef {data['coef_dotation_amortissements']})")
            subv_str = f"{data['nom']} | {', '.join(cats)} | Cd: {data['cd']} | Taux: {data['taux']}%"
            item.setText(subv_str)

    def save_project(self):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        theme_principal = self.theme_principal_combo.currentText()  # Récupère le thème principal
        if self.projet_id:
            # Mise à jour du projet existant
            cursor.execute('''UPDATE projets SET code=?, nom=?, details=?, date_debut=?, date_fin=?, livrables=?, chef=?, etat=?, cir=?, subvention=?, theme_principal=? WHERE id=?''', (
                self.code_edit.text().strip(),
                self.nom_edit.text().strip(),
                self.details_edit.toPlainText().strip(),
                self.date_debut.text(),
                self.date_fin.text(),
                self.livrables_edit.text().strip(),
                self.chef_edit.text().strip(),
                self.etat_combo.currentText(),
                int(self.cir_check.isChecked()),
                1 if len(self.subventions_data) > 0 else 0,  # Automatique selon la liste
                theme_principal,
                self.projet_id
            ))
            projet_id = self.projet_id
            # Met à jour les thèmes liés
            cursor.execute('DELETE FROM projet_themes WHERE projet_id=?', (projet_id,))
            for nom in self.selected_themes:
                cursor.execute('SELECT id FROM themes WHERE nom=?', (nom,))
                res = cursor.fetchone()
                if res:
                    theme_id = res[0]
                    cursor.execute('INSERT INTO projet_themes (projet_id, theme_id) VALUES (?, ?)', (projet_id, theme_id))
        else:
            # Création d'un nouveau projet
            cursor.execute('''INSERT INTO projets (code, nom, details, date_debut, date_fin, livrables, chef, etat, cir, subvention, theme_principal)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', (
                self.code_edit.text().strip(),
                self.nom_edit.text().strip(),
                self.details_edit.toPlainText().strip(),
                self.date_debut.text(),
                self.date_fin.text(),
                self.livrables_edit.text().strip(),
                self.chef_edit.text().strip(),
                self.etat_combo.currentText(),
                int(self.cir_check.isChecked()),
                1 if len(self.subventions_data) > 0 else 0,  # Automatique selon la liste
                theme_principal
            ))
            projet_id = cursor.lastrowid
        conn.commit()
        conn.close()
        self.accept()

    def load_project_data(self):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Chargement des données de base du projet
        cursor.execute('SELECT code, nom, details, date_debut, date_fin, livrables, chef, etat, cir, subvention, theme_principal FROM projets WHERE id=?', (self.projet_id,))
        res = cursor.fetchone()
        if res:
            self.code_edit.setText(str(res[0]))
            self.nom_edit.setText(str(res[1]))
            self.details_edit.setPlainText(str(res[2]))
            # Dates : conversion MM/yyyy vers QDate
            try:
                debut = res[3]
                fin = res[4]
                if debut:
                    m, y = debut.split('/')
                    self.date_debut.setDate(datetime.date(int(y), int(m), 1))
                if fin:
                    m, y = fin.split('/')
                    self.date_fin.setDate(datetime.date(int(y), int(m), 1))
            except Exception:
                pass
            self.livrables_edit.setText(str(res[5]))
            self.chef_edit.setText(str(res[6]))
            idx = self.etat_combo.findText(str(res[7]))
            if idx >= 0:
                self.etat_combo.setCurrentIndex(idx)
            self.cir_check.setChecked(bool(res[8]))
            # Note: subvention est maintenant gérée automatiquement via la liste
            
            # Thème principal
            theme_principal = res[10] if len(res) > 10 and res[10] else ""
            if theme_principal and hasattr(self, 'theme_principal_combo'):
                idx = self.theme_principal_combo.findText(theme_principal)
                if idx >= 0:
                    self.theme_principal_combo.setCurrentIndex(idx)
            
        # Thèmes liés
        cursor.execute('SELECT t.nom FROM projet_themes pt JOIN themes t ON pt.theme_id = t.id WHERE pt.projet_id=?', (self.projet_id,))
        self.selected_themes = [nom for (nom,) in cursor.fetchall()]
        self.update_theme_tags()
        
        # Images liées
        cursor.execute('SELECT nom FROM images WHERE projet_id=?', (self.projet_id,))
        self.images_list.clear()
        for (img_name,) in cursor.fetchall():
            self.images_list.addItem(os.path.join('images', img_name))
            
        # Investissements liés
        self.invest_list.clear()
        cursor.execute('SELECT nom, montant, date_achat, duree FROM investissements WHERE projet_id=?', (self.projet_id,))
        for nom, montant, date_achat, duree in cursor.fetchall():
            nom_display = nom if nom else "Sans nom"  # Gestion des anciens investissements sans nom
            invest_str = f"Nom: {nom_display}, Montant: {montant} €, Date achat: {date_achat}, Durée amort.: {duree} ans"
            self.invest_list.addItem(invest_str)
            
        # Subventions liées
        self.subv_list.clear()
        self.subventions_data = []
        cursor.execute('SELECT depenses_temps_travail, coef_temps_travail, depenses_externes, coef_externes, depenses_autres_achats, coef_autres_achats, depenses_dotation_amortissements, coef_dotation_amortissements, cd, taux, nom FROM subventions WHERE projet_id=?', (self.projet_id,))
        for subv in cursor.fetchall():
            data = {
                'depenses_temps_travail': subv[0],
                'coef_temps_travail': subv[1],
                'depenses_externes': subv[2],
                'coef_externes': subv[3],
                'depenses_autres_achats': subv[4],
                'coef_autres_achats': subv[5],
                'depenses_dotation_amortissements': subv[6],
                'coef_dotation_amortissements': subv[7],
                'cd': subv[8],
                'taux': subv[9],
                'nom': subv[10]
            }
            self.subventions_data.append(data)
            cats = []
            if subv[0]: cats.append(f"Temps travail (coef {subv[1]})")
            if subv[2]: cats.append(f"Externes (coef {subv[3]})")
            if subv[4]: cats.append(f"Autres achats (coef {subv[5]})")
            if subv[6]: cats.append(f"Dotation (coef {subv[7]})")
            nom = subv[10] if subv[10] else "Sans nom"
            subv_str = f"{nom} | {', '.join(cats)} | Cd: {subv[8]} | Taux: {subv[9]}%"
            self.subv_list.addItem(subv_str)
            
        # Initialiser self.equipe_data avec des valeurs par défaut pour toutes les directions
        self.equipe_data = {dir_: {label: 0 for label in self.equipe_types_labels} for dir_ in self.directions}
        
        # Équipe liée - Charger toutes les données d'équipe de la base de données
        cursor.execute('SELECT direction, type, nombre FROM equipe WHERE projet_id=?', (self.projet_id,))
        for direction, type_, nombre in cursor.fetchall():
            if direction in self.equipe_data and type_ in self.equipe_data[direction]:
                self.equipe_data[direction][type_] = nombre
            
        
        # Désactiver temporairement le signal currentTextChanged pour éviter un appel automatique à on_direction_changed
        self.direction_combo.blockSignals(True)
        
        # Parcourir chaque direction pour mettre à jour les spinbox
        for direction in self.directions:
            # Sélectionner cette direction temporairement
            self.direction_combo.setCurrentText(direction)
        
            
            # Mettre à jour les spinbox pour cette direction
            for label in self.equipe_types_labels:
                # Bloquer temporairement les signaux pour éviter des appels récursifs
                self.equipe_spins[label].blockSignals(True)
                self.equipe_spins[label].setValue(self.equipe_data[direction][label])
                self.equipe_spins[label].blockSignals(False)
        
        # Restaurer la direction initiale (celle qui était sélectionnée avant)
        current_direction = self.direction_combo.currentText()
        self._current_direction = current_direction
        
        # Réactiver le signal
        self.direction_combo.blockSignals(False)
    
        
        conn.close()

    def image_context_menu(self, pos):
        item = self.images_list.itemAt(pos)
        if item:
            menu = QMenu()
            delete_action = menu.addAction('Supprimer')
            action = menu.exec(self.images_list.mapToGlobal(pos))
            if action == delete_action:
                confirm = QMessageBox.question(
                    self,
                    'Confirmation',
                    f'Voulez-vous vraiment supprimer cette image ?',
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if confirm == QMessageBox.StandardButton.Yes:
                    # Vérifie si l'image est utilisée dans d'autres projets avant suppression du fichier
                    img_rel_path = item.text()
                    img_name = os.path.basename(img_rel_path)
                    conn = sqlite3.connect(DB_PATH)
                    cursor = conn.cursor()
                    cursor.execute('SELECT COUNT(*) FROM images WHERE nom=?', (img_name,))
                    count = cursor.fetchone()[0]
                    conn.close()
                    # Si cette image n'est plus utilisée (après suppression de la ligne courante)
                    if count <= 1:
                        img_path = os.path.join(os.path.dirname(__file__), img_rel_path)
                        try:
                            if os.path.isfile(img_path):
                                os.remove(img_path)
                        except Exception:
                            pass  # Suppression silencieuse
                    self.images_list.takeItem(self.images_list.row(item))
        # Si clic droit sur une zone vide, rien ne se passe

class InvestDialog(QDialog):
    def __init__(self, parent=None, nom='', montant='', date_achat='', duree=''):
        super().__init__(parent)
        self.setWindowTitle('Ajouter investissement')
        layout = QFormLayout()
        self.nom_edit = QLineEdit(nom)
        self.montant_edit = QLineEdit(montant)
        self.date_achat = QLineEdit(date_achat)
        self.duree_edit = QLineEdit(duree)
        layout.addRow('Nom:', self.nom_edit)
        layout.addRow('Montant (€):', self.montant_edit)
        layout.addRow('Date achat (MM/AAAA):', self.date_achat)
        layout.addRow('Durée amortissement (ans):', self.duree_edit)
        btns = QHBoxLayout()
        btn_ok = QPushButton('Valider')
        btn_cancel = QPushButton('Annuler')
        btn_ok.clicked.connect(self.validate_and_accept)
        btn_cancel.clicked.connect(self.reject)
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)
        layout.addRow(btns)
        self.setLayout(layout)

    def validate_and_accept(self):
        nom = self.nom_edit.text().strip()
        montant = self.montant_edit.text().replace(',', '.').strip()
        date_achat = self.date_achat.text().strip()
        duree = self.duree_edit.text().strip()
        # Contrôle nom
        if not nom:
            QMessageBox.warning(self, 'Erreur', 'Le nom de l\'investissement est obligatoire.')
            return
        # Contrôle montant
        try:
            if float(montant) <= 0:
                raise ValueError
        except Exception:
            QMessageBox.warning(self, 'Erreur', 'Le montant doit être un nombre positif.')
            return
        # Contrôle durée
        try:
            if int(duree) <= 0:
                raise ValueError
        except Exception:
            QMessageBox.warning(self, 'Erreur', 'La durée doit être un entier positif.')
            return
        # Contrôle date MM/AAAA
        if not re.match(r'^(0[1-9]|1[0-2])/\d{4}$', date_achat):
            QMessageBox.warning(self, 'Erreur', 'La date doit être au format MM/AAAA.')
            return
        self.accept()

class ThemeManager(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Gestion des thèmes')
        self.setMinimumWidth(400)
        layout = QVBoxLayout()
        self.theme_list = QListWidget()
        layout.addWidget(self.theme_list)
        btns = QHBoxLayout()
        self.theme_input = QLineEdit()
        btn_add = QPushButton('Ajouter')
        btn_del = QPushButton('Supprimer')
        btns.addWidget(self.theme_input)
        btns.addWidget(btn_add)
        btns.addWidget(btn_del)
        layout.addLayout(btns)
        self.setLayout(layout)
        btn_add.clicked.connect(self.add_theme)
        btn_del.clicked.connect(self.delete_theme)
        self.load_themes()

    def load_themes(self):
        self.theme_list.clear()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT nom FROM themes ORDER BY nom')
        for (nom,) in cursor.fetchall():
            self.theme_list.addItem(nom)
        conn.close()

    def add_theme(self):
        nom = self.theme_input.text().strip()
        if not nom:
            return
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            cursor.execute('INSERT INTO themes (nom) VALUES (?)', (nom,))
            conn.commit()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, 'Erreur', 'Ce thème existe déjà.')
        conn.close()
        self.theme_input.clear()
        self.load_themes()

    def delete_theme(self):
        item = self.theme_list.currentItem()
        if not item:
            return
        confirm = QMessageBox.question(
            self,
            'Confirmation',
            f'Voulez-vous vraiment supprimer le thème "{item.text()}" ?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm == QMessageBox.StandardButton.Yes:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM themes WHERE nom=?', (item.text(),))
            conn.commit()
            conn.close()
            self.load_themes()

    def edit_theme(self, item):
        old_theme = item.text()
        new_theme, ok = QInputDialog.getText(self, 'Modifier le thème', 'Nouveau nom du thème:', text=old_theme)
        if ok and new_theme and new_theme != old_theme:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            try:
                cursor.execute('UPDATE themes SET nom=? WHERE nom=?', (new_theme, old_theme))
                conn.commit()
            except sqlite3.IntegrityError:  # Correction ici
                QMessageBox.warning(self, 'Erreur', 'Ce thème existe déjà.')
            conn.close()
            self.load_themes()

class DirectionManager(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Gestion des directions')
        self.setMinimumWidth(400)
        layout = QVBoxLayout()
        self.direction_list = QListWidget()
        layout.addWidget(self.direction_list)
        btns = QHBoxLayout()
        self.direction_input = QLineEdit()
        btn_add = QPushButton('Ajouter')
        btn_del = QPushButton('Supprimer')
        btns.addWidget(self.direction_input)
        btns.addWidget(btn_add)
        btns.addWidget(btn_del)
        layout.addLayout(btns)
        self.setLayout(layout)
        btn_add.clicked.connect(self.add_direction)
        btn_del.clicked.connect(self.delete_direction)
        self.load_directions()

    def load_directions(self):
        self.direction_list.clear()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('SELECT nom FROM directions ORDER BY nom')
        for (nom,) in cursor.fetchall():
            self.direction_list.addItem(nom)
        conn.close()

    def add_direction(self):
        nom = self.direction_input.text().strip()
        if not nom:
            return
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            cursor.execute('INSERT INTO directions (nom) VALUES (?)', (nom,))
            conn.commit()
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, 'Erreur', 'Cette direction existe déjà.')
        conn.close()
        self.direction_input.clear()
        self.load_directions()

    def delete_direction(self):
        item = self.direction_list.currentItem()
        if not item:
            return
        confirm = QMessageBox.question(
            self,
            'Confirmation',
            f'Voulez-vous vraiment supprimer la direction "{item.text()}" ?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm == QMessageBox.StandardButton.Yes:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute('DELETE FROM directions WHERE nom=?', (item.text(),))
            conn.commit()
            conn.close()
            self.load_directions()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())