
from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QPushButton, QSpinBox, QMessageBox
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
        self.ensure_table_exists()
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
        self.table = QTableWidget(len(CATEGORIES), 5)
        self.table.setHorizontalHeaderLabels(['Catégorie', 'Libellé', 'Montant chargé', 'Coût de production', 'Coût complet'])
        self.table.setColumnWidth(3, 140)  # Coût de production
        self.table.setColumnWidth(1, 140)
        for i, (code, libelle) in enumerate(CATEGORIES):
            self.table.setItem(i, 0, QTableWidgetItem(code))
            self.table.setItem(i, 1, QTableWidgetItem(libelle))
        main_layout.addWidget(self.table)

        # Bouton enregistrer
        btn_layout = QHBoxLayout()
        self.save_btn = QPushButton('Enregistrer')
        self.save_btn.clicked.connect(self.confirm_save)
        btn_layout.addStretch()
        btn_layout.addWidget(self.save_btn)
        main_layout.addLayout(btn_layout)

        self.table.cellChanged.connect(self.mark_dirty)
        self.table.installEventFilter(self)

        self.setLayout(main_layout)
        self._dirty = False
        self._loading = False
        self.brouillons = {}  # année -> valeurs du tableau
        self.current_year = self.year_spin.value()
        self.load_data()
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
        for i, (code, libelle) in enumerate(CATEGORIES):
            row = []
            for j in range(2, 5):
                item = self.table.item(i, j)
                row.append(item.text() if item and item.text() else '')
            values[code] = row
        return values

    def set_table_values(self, values):
        self._loading = True
        if values:
            for i, (code, libelle) in enumerate(CATEGORIES):
                for j in range(2, 5):
                    self.table.setItem(i, j, QTableWidgetItem(values.get(code, ['', '', ''])[j-2]))
        else:
            # Si pas de brouillon, charger depuis la base
            year = self.current_year
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            for i, (code, libelle) in enumerate(CATEGORIES):
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
        for i, (code, libelle) in enumerate(CATEGORIES):
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

