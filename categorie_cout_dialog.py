
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QPushButton, QSpinBox, QMessageBox, QInputDialog

from database import get_connection
from category_utils import DEFAULT_CATEGORIES, invalidate_category_cache
class CategorieCoutDialog(QDialog):
    def eventFilter(self, obj, event):
        from PyQt6.QtCore import QEvent
        if obj == self.table and event.type() == QEvent.Type.KeyPress:
            from PyQt6.QtGui import QKeySequence
            if event.matches(QKeySequence.StandardKey.Copy):
                self.copy_to_clipboard()
                return True
            elif event.matches(QKeySequence.StandardKey.Paste):
                self.paste_from_clipboard()
                return True
        return super().eventFilter(obj, event)

    def copy_to_clipboard(self):
        from PyQt6.QtWidgets import QApplication
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges:
            return
        
        # Prendre la première sélection
        selection = selected_ranges[0]
        rows = []
        for row in range(selection.topRow(), selection.bottomRow() + 1):
            cols = []
            for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                item = self.table.item(row, col)
                cols.append(item.text() if item else '')
            rows.append('\t'.join(cols))
        
        text = '\n'.join(rows)
        clipboard = QApplication.clipboard()
        clipboard.setText(text)

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
        self.base_categories = list(DEFAULT_CATEGORIES)
        
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
        self.table.cellChanged.connect(self.validate_category_code)
        self.table.cellChanged.connect(self.auto_save)  # Sauvegarde automatique
        self.table.installEventFilter(self)

        self.setLayout(main_layout)
        self.current_year = self.year_spin.value()
        self.load_data()
    
    def get_categories(self):
        """Retourne la liste complète des catégories (prédéfinies + personnalisées)"""
        return self.base_categories + self.custom_categories
    
    def populate_table(self):
        """Remplit le tableau avec toutes les catégories"""
        categories = self.get_categories()
        for i, (code, libelle) in enumerate(categories):
            self.table.setItem(i, 0, QTableWidgetItem(code))
            self.table.setItem(i, 1, QTableWidgetItem(libelle))
    
    def load_custom_categories(self):
        """Charge les catégories personnalisées depuis la base de données"""
        conn = get_connection()
        cursor = conn.cursor()
        predefined_codes = [code for code, _ in self.base_categories]
        self.custom_categories = []

        if predefined_codes:
            placeholders = ','.join(['?'] * len(predefined_codes))
            cursor.execute(
                f'''SELECT DISTINCT categorie, libelle FROM categorie_cout 
                    WHERE categorie NOT IN ({placeholders})
                    ORDER BY categorie''',
                predefined_codes,
            )
        else:
            cursor.execute(
                '''SELECT DISTINCT categorie, libelle FROM categorie_cout 
                   ORDER BY categorie'''
            )

        for code, libelle in cursor.fetchall():
            code = (code or '').strip()
            if not code:
                continue
            label = (libelle or '').strip() or code
            self.custom_categories.append((code, label))
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
            
            # Aussi vérifier les codes actuellement dans le tableau
            for i in range(self.table.rowCount()):
                item = self.table.item(i, 0)
                if item and item.text().strip().upper() == code:
                    existing_codes.append(code)
                    break
            
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
        conn = get_connection()
        cursor = conn.cursor()
        
        # Récupérer toutes les années existantes
        cursor.execute('SELECT DISTINCT annee FROM categorie_cout ORDER BY annee')
        years = [row[0] for row in cursor.fetchall()]
        
        # Si aucune année n'existe, utiliser l'année courante
        if not years:
            years = [self.current_year]
        
        # Ajouter la catégorie pour chaque année
        for year in years:
            cursor.execute('''INSERT OR IGNORE INTO categorie_cout (annee, categorie, libelle) VALUES (?, ?, ?)''', 
                          (year, code, libelle))
        
        conn.commit()
        conn.close()
        invalidate_category_cache()
    
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
        
        # Demander confirmation
        reply = QMessageBox.question(self, 'Confirmation', 
                                   f'Voulez-vous vraiment supprimer la catégorie "{code}" ?\n'
                                   'Cette action supprimera toutes les données associées pour toutes les années.',
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            # Supprimer de la base de données
            self.delete_category_from_db(code)
            
            # Supprimer des catégories personnalisées si c'en est une
            self.custom_categories = [(c, l) for c, l in self.custom_categories if c != code]
            
            # Si c'est une catégorie prédéfinie, la retirer temporairement de la liste de base
            self.base_categories = [(c, l) for c, l in self.base_categories if c != code]
            
            # Sauvegarder les brouillons actuels
            self.brouillons[self.current_year] = self.get_table_values()
            
            # Recréer le tableau sans la ligne supprimée
            self.table.setRowCount(len(self.get_categories()))
            self.populate_table()
            
            # Recharger les données
            self.set_table_values(self.brouillons.get(self.current_year, None))
            
            # Marquer comme modifié
            self._dirty = True
            invalidate_category_cache()
            
            QMessageBox.information(self, 'Succès', 
                                  f'La catégorie "{code}" a été supprimée.')

    def delete_category_from_db(self, code):
        """Supprime une catégorie de la base de données pour toutes les années"""
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM categorie_cout WHERE categorie = ?', (code,))
        conn.commit()
        conn.close()
        invalidate_category_cache()

    def update_category_code_in_lists(self, old_code, new_code):
        """Met à jour le code de catégorie dans les listes internes"""
        # Mettre à jour dans la liste de base si c'est une catégorie prédéfinie
        self.base_categories = [
            (new_code if code == old_code else code, libelle)
            for code, libelle in self.base_categories
        ]
        
        # Mettre à jour dans custom_categories si c'est une catégorie personnalisée
        self.custom_categories = [(new_code if code == old_code else code, libelle) 
                                 for code, libelle in self.custom_categories]
        invalidate_category_cache()

    def change_year(self):
        # Sauvegarde le brouillon courant
        self.brouillons[self.current_year] = self.get_table_values()
        # Change l'année affichée
        self.current_year = self.year_spin.value()
        # Charge le brouillon de l'année sélectionnée ou les valeurs de la base
        self.set_table_values(self.brouillons.get(self.current_year, None))
    def get_table_values(self):
        # Retourne les valeurs du tableau sous forme de dict {original_code: [current_code, libelle, montant_charge, cout_production, cout_complet]}
        values = {}
        categories = self.get_categories()
        for i, (original_code, default_libelle) in enumerate(categories):
            if i < self.table.rowCount():  # S'assurer que la ligne existe
                row = []
                # Code de catégorie
                code_item = self.table.item(i, 0)
                row.append(code_item.text() if code_item else original_code)
                # Libellé
                libelle_item = self.table.item(i, 1)
                row.append(libelle_item.text() if libelle_item else default_libelle)
                # Montants
                for j in range(2, 5):
                    item = self.table.item(i, j)
                    row.append(item.text() if item and item.text() else '')
                values[original_code] = row
        return values

    def set_table_values(self, values):
        self._loading = True
        categories = self.get_categories()
        if values:
            for i, (original_code, default_libelle) in enumerate(categories):
                if i < self.table.rowCount():  # S'assurer que la ligne existe
                    if original_code in values:
                        saved_values = values[original_code]
                        # Code de catégorie (index 0)
                        self.table.setItem(i, 0, QTableWidgetItem(saved_values[0] if saved_values[0] else original_code))
                        # Libellé (index 1)
                        self.table.setItem(i, 1, QTableWidgetItem(saved_values[1] if len(saved_values) > 1 and saved_values[1] else default_libelle))
                        # Montants (indices 2, 3, 4)
                        for j in range(2, 5):
                            self.table.setItem(i, j, QTableWidgetItem(saved_values[j] if len(saved_values) > j else ''))
        else:
            # Si pas de brouillon, charger depuis la base
            year = self.current_year
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute(
                '''SELECT categorie, libelle, montant_charge, cout_production, cout_complet 
                   FROM categorie_cout WHERE annee=?''',
                (year,),
            )
            rows = cursor.fetchall()
            conn.close()

            # Préparer un accès direct par code pour éviter une requête par ligne
            data_by_code = {}
            for db_code, libelle, montant_charge, cout_production, cout_complet in rows:
                normalized_code = (db_code or '').strip().upper()
                if not normalized_code:
                    continue
                data_by_code[normalized_code] = (
                    libelle,
                    montant_charge,
                    cout_production,
                    cout_complet,
                )

            for i, (code, default_libelle) in enumerate(categories):
                if i >= self.table.rowCount():
                    continue

                normalized_code = (code or '').strip().upper()
                db_values = data_by_code.get(normalized_code)

                # Charger le libellé depuis la DB si disponible, sinon utiliser le défaut
                libelle_to_use = db_values[0] if db_values and db_values[0] else default_libelle
                self.table.setItem(i, 1, QTableWidgetItem(libelle_to_use))

                montant_charge = db_values[1] if db_values else None
                cout_production = db_values[2] if db_values else None
                cout_complet = db_values[3] if db_values else None

                self.table.setItem(i, 2, QTableWidgetItem('' if montant_charge is None else str(montant_charge)))
                self.table.setItem(i, 3, QTableWidgetItem('' if cout_production is None else str(cout_production)))
                self.table.setItem(i, 4, QTableWidgetItem('' if cout_complet is None else str(cout_complet)))
        self._loading = False

    def mark_dirty(self, row, column):
        if not getattr(self, '_loading', False) and column in [0, 1, 2, 3, 4]:
            self._dirty = True
    
    def auto_save(self, row, column):
        """Sauvegarde automatique après modification d'une cellule"""
        if not getattr(self, '_loading', False) and column in [2, 3, 4]:  # Seulement pour les montants
            # Délai court pour permettre à l'utilisateur de finir sa saisie
            from PyQt6.QtCore import QTimer
            if not hasattr(self, '_auto_save_timer'):
                self._auto_save_timer = QTimer()
                self._auto_save_timer.setSingleShot(True)
                self._auto_save_timer.timeout.connect(lambda: self.save_data(show_message=False))
            self._auto_save_timer.stop()
            self._auto_save_timer.start(1000)  # Sauvegarde après 1 seconde d'inactivité
    
    def validate_category_code(self, row, column):
        """Valide et formate le code de catégorie (colonne 0)"""
        if column == 0 and not getattr(self, '_loading', False):
            item = self.table.item(row, column)
            if item:
                text = item.text()
                # Limiter à 3 caractères et convertir en majuscules
                formatted_text = text.upper()[:3]
                if text != formatted_text:
                    # Désactiver temporairement les signaux pour éviter la récursion
                    self.table.blockSignals(True)
                    item.setText(formatted_text)
                    self.table.blockSignals(False)
    def confirm_save(self):
        reply = QMessageBox.question(self, 'Confirmation', 'Voulez-vous enregistrer les modifications ?',
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.save_data(show_message=True)
            self._dirty = False
            self.accept()

    def save_data(self, show_message=True):
        year = self.year_spin.value()
        conn = get_connection()
        cursor = conn.cursor()
        categories = self.get_categories()

        cursor.execute(
            '''SELECT id, categorie FROM categorie_cout WHERE annee=?''',
            (year,),
        )
        existing_ids_by_code = {}
        for row_id, db_code in cursor.fetchall():
            normalized_code = (db_code or '').strip().upper()
            if normalized_code:
                existing_ids_by_code[normalized_code] = row_id

        def parse_float(item):
            if item and item.text() and item.text().strip():
                try:
                    return float(item.text().replace(',', '.'))
                except Exception:
                    return None
            return None
        
        for i, (original_code, original_libelle) in enumerate(categories):
            if i < self.table.rowCount():  # S'assurer que la ligne existe
                # Récupérer les valeurs modifiées depuis le tableau
                code_item = self.table.item(i, 0)
                current_code = ''
                if code_item and code_item.text():
                    current_code = code_item.text().strip().upper()
                if not current_code:
                    current_code = (original_code or '').strip().upper()

                libelle_item = self.table.item(i, 1)
                current_libelle = libelle_item.text() if libelle_item and libelle_item.text() else original_libelle

                original_code_key = (original_code or '').strip().upper()
                code_changed = current_code != original_code_key
                original_id = existing_ids_by_code.get(original_code_key)
                target_id = existing_ids_by_code.get(current_code)

                if code_changed:
                    if target_id is not None and target_id != original_id:
                        if show_message:
                            QMessageBox.warning(self, 'Erreur', f'Le code "{current_code}" existe déjà pour cette année.')
                        continue

                    reply = QMessageBox.question(
                        self,
                        'Changement de code',
                        f'Voulez-vous changer le code "{original_code}" en "{current_code}" pour toutes les années ?\n'
                        'Si vous choisissez "Non", seule l\'année courante sera modifiée.',
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    )

                    if reply == QMessageBox.StandardButton.Yes:
                        cursor.execute(
                            '''UPDATE categorie_cout SET categorie=? WHERE categorie=?''',
                            (current_code, original_code),
                        )
                        self.update_category_code_in_lists(original_code, current_code)

                    if original_id is not None:
                        existing_ids_by_code.pop(original_code_key, None)
                        existing_ids_by_code[current_code] = original_id
                else:
                    if original_id is not None:
                        existing_ids_by_code[original_code_key] = original_id

                update_fields = {'libelle': current_libelle}
                if code_changed:
                    update_fields['categorie'] = current_code

                update_fields['montant_charge'] = parse_float(self.table.item(i, 2))
                update_fields['cout_production'] = parse_float(self.table.item(i, 3))
                update_fields['cout_complet'] = parse_float(self.table.item(i, 4))

                record_key = current_code if code_changed else original_code_key
                record_id = existing_ids_by_code.get(record_key)

                if record_id:
                    set_clause = ', '.join([f"{k}=?" for k in update_fields.keys()])
                    sql = f"UPDATE categorie_cout SET {set_clause} WHERE id=?"
                    cursor.execute(sql, list(update_fields.values()) + [record_id])
                else:
                    update_fields['annee'] = year
                    update_fields['categorie'] = current_code

                    fields = ', '.join(update_fields.keys())
                    placeholders = ', '.join(['?'] * len(update_fields))
                    sql = f"INSERT INTO categorie_cout ({fields}) VALUES ({placeholders})"
                    cursor.execute(sql, list(update_fields.values()))
                    existing_ids_by_code[current_code] = cursor.lastrowid
        
        conn.commit()
        conn.close()
        invalidate_category_cache()
        if show_message:
            QMessageBox.information(self, 'Sauvegarde', 'Les coûts ont été enregistrés avec succès.')

    def ensure_table_exists(self):
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS categorie_cout (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            annee INTEGER,
            categorie TEXT,
            libelle TEXT,
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

