
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QPushButton, QSpinBox, QMessageBox, QInputDialog
import sqlite3

CATEGORIES = [
    ("STP", "Stagiaire Projet"),
    ("AOP", "Assistante / opérateur"),
    ("TEP", "Technicien"),
    ("IJP", "Junior"),
    ("ISP", "Senior"),
    ("EDP", "Expert"),
    ("MOY", "Collaborateur moyen")
]
DB_PATH = 'gestion_budget.db'

class CategorieCoutDialog(QDialog):
    def eventFilter(self, obj, event):
        from PyQt6.QtCore import QEvent
        if obj == self.table and event.type() == QEvent.Type.KeyPress:
            from PyQt6.QtGui import QKeySequence
            if event.matches(QKeySequence.StandardKey.Paste):
                self.paste_from_clipboard()
                return True
        return super().eventFilter(obj, event)

    def paste_from_clipboard(self):
        from PyQt6.QtWidgets import QApplication
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return
        rows = text.split('\n')
        start_row = self.table.currentRow()
        start_col = self.table.currentColumn()
        for i, row_text in enumerate(rows):
            if not row_text.strip():
                continue
            cols = row_text.split('\t')
            for j, value in enumerate(cols):
                r = start_row + i
                c = start_col + j
                if r < self.table.rowCount() and c < self.table.columnCount():
                    self.table.setItem(r, c, QTableWidgetItem(value))
        self._dirty = True
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Gestion des coûts par catégorie')
        self.resize(700, 350)
        
        # Initialiser les attributs avant tout
        self.custom_categories = []  # Pour stocker les catégories personnalisées
        self._dirty = False
        self._loading = False
        self.brouillons = {}  # année -> valeurs du tableau
        
        self.ensure_table_exists()
        self.load_custom_categories()
        
        main_layout = QVBoxLayout()

        # Sélection de l'année
        year_layout = QHBoxLayout()
        year_layout.addWidget(QLabel('Année :'))
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2000, 2100)
        self.year_spin.setValue(self.get_default_year())
        self.year_spin.valueChanged.connect(self.change_year)
        year_layout.addWidget(self.year_spin)
        main_layout.addLayout(year_layout)

        # Tableau des catégories
        self.table = QTableWidget(len(self.get_categories()), 5)
        self.table.setHorizontalHeaderLabels(['Catégorie', 'Libellé', 'Montant chargé', 'Coût de production', 'Coût complet'])
        self.table.setColumnWidth(3, 140)  # Coût de production
        self.table.setColumnWidth(1, 140)
        self.populate_table()
        main_layout.addWidget(self.table)

        # Boutons
        btn_layout = QHBoxLayout()
        self.add_row_btn = QPushButton('Ajouter ligne')
        self.add_row_btn.clicked.connect(self.add_new_category)
        btn_layout.addWidget(self.add_row_btn)
        
        self.delete_row_btn = QPushButton('Supprimer ligne')
        self.delete_row_btn.clicked.connect(self.delete_selected_category)
        btn_layout.addWidget(self.delete_row_btn)
        
        btn_layout.addStretch()
        
        self.save_btn = QPushButton('Enregistrer')
        self.save_btn.clicked.connect(self.confirm_save)
        btn_layout.addWidget(self.save_btn)
        main_layout.addLayout(btn_layout)

        self.table.cellChanged.connect(self.mark_dirty)
        self.table.installEventFilter(self)

        self.setLayout(main_layout)
        self.current_year = self.year_spin.value()
        self.load_data()
    
    def get_categories(self):
        """Retourne la liste complète des catégories (prédéfinies + personnalisées)"""
        return CATEGORIES + self.custom_categories
    
    def populate_table(self):
        """Remplit le tableau avec toutes les catégories"""
        categories = self.get_categories()
        for i, (code, libelle) in enumerate(categories):
            self.table.setItem(i, 0, QTableWidgetItem(code))
            self.table.setItem(i, 1, QTableWidgetItem(libelle))
    
    def load_custom_categories(self):
        """Charge les catégories personnalisées depuis la base de données"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        # Récupérer toutes les catégories uniques qui ne sont pas dans les catégories prédéfinies
        predefined_codes = [code for code, _ in CATEGORIES]
        placeholders = ','.join(['?' for _ in predefined_codes])
        cursor.execute(f'''SELECT DISTINCT categorie FROM categorie_cout 
                          WHERE categorie NOT IN ({placeholders})
                          ORDER BY categorie''', predefined_codes)
        
        for row in cursor.fetchall():
            code = row[0]
            # Pour les catégories personnalisées, utiliser le code comme libellé par défaut
            self.custom_categories.append((code, code))
        conn.close()
    
    def add_new_category(self):
        """Ajoute une nouvelle catégorie personnalisée"""
        code, ok = QInputDialog.getText(self, 'Nouvelle catégorie', 
                                       'Code de la catégorie (3 lettres max):')
        if ok and code:
            code = code.strip().upper()[:3]  # Limiter à 3 caractères
            if not code:
                QMessageBox.warning(self, 'Erreur', 'Le code ne peut pas être vide.')
                return
            
            # Vérifier si la catégorie existe déjà
            all_categories = self.get_categories()
            existing_codes = [cat[0] for cat in all_categories]
            if code in existing_codes:
                QMessageBox.warning(self, 'Erreur', 'Cette catégorie existe déjà.')
                return
            
            libelle, ok = QInputDialog.getText(self, 'Nouvelle catégorie', 
                                             'Libellé de la catégorie:', text=code)
            if ok:
                libelle = libelle.strip() if libelle.strip() else code
                
                # Demander confirmation pour toutes les années
                reply = QMessageBox.question(self, 'Confirmation', 
                                           f'Voulez-vous ajouter la catégorie "{code}" pour toutes les années existantes dans la base de données ?',
                                           QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                
                if reply == QMessageBox.StandardButton.Yes:
                    # Ajouter la catégorie pour toutes les années existantes
                    self.add_category_for_all_years(code, libelle)
                
                # Ajouter la catégorie personnalisée
                self.custom_categories.append((code, libelle))
                
                # Sauvegarder les brouillons actuels
                self.brouillons[self.current_year] = self.get_table_values()
                
                # Recréer le tableau avec la nouvelle ligne
                self.table.setRowCount(len(self.get_categories()))
                self.populate_table()
                
                # Recharger les données
                self.set_table_values(self.brouillons.get(self.current_year, None))
                
                # Marquer comme modifié
                self._dirty = True
                
                QMessageBox.information(self, 'Succès', 
                                      f'La catégorie "{code}" a été ajoutée.')
    
    def add_category_for_all_years(self, code, libelle):
        """Ajoute une nouvelle catégorie vide pour toutes les années existantes"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Récupérer toutes les années existantes
        cursor.execute('SELECT DISTINCT annee FROM categorie_cout ORDER BY annee')
        years = [row[0] for row in cursor.fetchall()]
        
        # Si aucune année n'existe, utiliser l'année courante
        if not years:
            years = [self.current_year]
        
        # Ajouter la catégorie pour chaque année
        for year in years:
            cursor.execute('''INSERT OR IGNORE INTO categorie_cout (annee, categorie) VALUES (?, ?)''', 
                          (year, code))
        
        conn.commit()
        conn.close()
    
    def delete_selected_category(self):
        """Supprime la catégorie sélectionnée"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, 'Erreur', 'Veuillez sélectionner une ligne à supprimer.')
            return
        
        categories = self.get_categories()
        if current_row >= len(categories):
            QMessageBox.warning(self, 'Erreur', 'Ligne invalide sélectionnée.')
            return
        
        code, libelle = categories[current_row]
        
        # Vérifier si c'est une catégorie prédéfinie
        predefined_codes = [cat[0] for cat in CATEGORIES]
        if code in predefined_codes:
            QMessageBox.warning(self, 'Erreur', 'Impossible de supprimer une catégorie prédéfinie.')
            return
        
        # Demander confirmation
        reply = QMessageBox.question(self, 'Confirmation', 
                                   f'Voulez-vous vraiment supprimer la catégorie "{code}" ?\n'
                                   'Cette action supprimera toutes les données associées pour toutes les années.',
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            # Supprimer de la base de données
            self.delete_category_from_db(code)
            
            # Supprimer des catégories personnalisées
            self.custom_categories = [(c, l) for c, l in self.custom_categories if c != code]
            
            # Sauvegarder les brouillons actuels
            self.brouillons[self.current_year] = self.get_table_values()
            
            # Recréer le tableau sans la ligne supprimée
            self.table.setRowCount(len(self.get_categories()))
            self.populate_table()
            
            # Recharger les données
            self.set_table_values(self.brouillons.get(self.current_year, None))
            
            # Marquer comme modifié
            self._dirty = True
            
            QMessageBox.information(self, 'Succès', 
                                  f'La catégorie "{code}" a été supprimée.')
    
    def delete_category_from_db(self, code):
        """Supprime une catégorie de la base de données pour toutes les années"""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM categorie_cout WHERE categorie = ?', (code,))
        conn.commit()
        conn.close()

    def change_year(self):
        # Sauvegarde le brouillon courant
        self.brouillons[self.current_year] = self.get_table_values()
        # Change l'année affichée
        self.current_year = self.year_spin.value()
        # Charge le brouillon de l'année sélectionnée ou les valeurs de la base
        self.set_table_values(self.brouillons.get(self.current_year, None))
    def get_table_values(self):
        # Retourne les valeurs du tableau sous forme de dict {categorie: [montant_charge, cout_production, cout_complet]}
        values = {}
        categories = self.get_categories()
        for i, (code, libelle) in enumerate(categories):
            if i < self.table.rowCount():  # S'assurer que la ligne existe
                row = []
                for j in range(2, 5):
                    item = self.table.item(i, j)
                    row.append(item.text() if item and item.text() else '')
                values[code] = row
        return values

    def set_table_values(self, values):
        self._loading = True
        categories = self.get_categories()
        if values:
            for i, (code, libelle) in enumerate(categories):
                if i < self.table.rowCount():  # S'assurer que la ligne existe
                    for j in range(2, 5):
                        self.table.setItem(i, j, QTableWidgetItem(values.get(code, ['', '', ''])[j-2]))
        else:
            # Si pas de brouillon, charger depuis la base
            year = self.current_year
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            for i, (code, libelle) in enumerate(categories):
                if i < self.table.rowCount():  # S'assurer que la ligne existe
                    cursor.execute('''SELECT montant_charge, cout_production, cout_complet FROM categorie_cout WHERE annee=? AND categorie=?''', (year, code))
                    res = cursor.fetchone()
                    for j in range(2, 5):
                        self.table.setItem(i, j, QTableWidgetItem(''))
                    if res:
                        self.table.setItem(i, 2, QTableWidgetItem('' if res[0] is None else str(res[0])))
                        self.table.setItem(i, 3, QTableWidgetItem('' if res[1] is None else str(res[1])))
                        self.table.setItem(i, 4, QTableWidgetItem('' if res[2] is None else str(res[2])))
            conn.close()
        self._loading = False

    def mark_dirty(self, row, column):
        if not getattr(self, '_loading', False) and column in [2, 3, 4]:
            self._dirty = True
    def confirm_save(self):
        reply = QMessageBox.question(self, 'Confirmation', 'Voulez-vous enregistrer les modifications ?',
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.save_data(show_message=True)
            self._dirty = False
            self.accept()

    def save_data(self, show_message=True):
        year = self.year_spin.value()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        categories = self.get_categories()
        for i, (code, libelle) in enumerate(categories):
            if i < self.table.rowCount():  # S'assurer que la ligne existe
                # Récupérer les valeurs existantes
                cursor.execute('''SELECT id, montant_charge, cout_production, cout_complet FROM categorie_cout WHERE annee=? AND categorie=?''', (year, code))
                res = cursor.fetchone()
                update_fields = {}
                # Montant chargé
                val = self.table.item(i, 2)
                if val and val.text().strip():
                    try:
                        update_fields['montant_charge'] = float(val.text().replace(',', '.'))
                    except Exception:
                        pass  # Ignore les erreurs de conversion
                # Coût de production
                val = self.table.item(i, 3)
                if val and val.text().strip():
                    try:
                        update_fields['cout_production'] = float(val.text().replace(',', '.'))
                    except Exception:
                        pass
                # Coût complet
                val = self.table.item(i, 4)
                if val and val.text().strip():
                    try:
                        update_fields['cout_complet'] = float(val.text().replace(',', '.'))
                    except Exception:
                        pass
                if res:
                    # Mise à jour partielle : ne modifie que les champs renseignés
                    set_clause = ', '.join([f"{k}=?" for k in update_fields.keys()])
                    if set_clause:
                        sql = f"UPDATE categorie_cout SET {set_clause} WHERE id=?"
                        cursor.execute(sql, list(update_fields.values()) + [res[0]])
                elif update_fields:
                    # Insertion uniquement si au moins un champ est renseigné
                    fields = ', '.join(update_fields.keys())
                    placeholders = ', '.join(['?'] * len(update_fields))
                    sql = f"INSERT INTO categorie_cout (annee, categorie, {fields}) VALUES (?, ?, {placeholders})"
                    cursor.execute(sql, [year, code] + list(update_fields.values()))
        conn.commit()
        conn.close()
        if show_message:
            QMessageBox.information(self, 'Sauvegarde', 'Les coûts ont été enregistrés avec succès.')

    def ensure_table_exists(self):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS categorie_cout (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            annee INTEGER,
            categorie TEXT,
            montant_charge REAL,
            cout_production REAL,
            cout_complet REAL
        )''')
        conn.commit()
        conn.close()

    def get_default_year(self):
        import datetime
        return datetime.datetime.now().year

    def load_data(self):
        self.set_table_values(None)
    def closeEvent(self, event):
        if self._dirty:
            reply = QMessageBox.question(self, 'Attention',
                'Des modifications non enregistrées vont être perdues. Voulez-vous enregistrer avant de quitter ?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel)
            if reply == QMessageBox.StandardButton.Yes:
                self.save_data(show_message=True)
                self._dirty = False
                event.accept()
            elif reply == QMessageBox.StandardButton.No:
                self._dirty = False
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

