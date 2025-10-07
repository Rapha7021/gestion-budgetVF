from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QPushButton, QSpinBox, QMessageBox
from PyQt6.QtCore import Qt

from database import get_connection

class CIRDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('CIR - Coefficients par année')
        self.resize(400, 100)
        self._dirty = False
        self._loading = False
        self.brouillons = {}
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

        # Tableau CIR
        self.table = QTableWidget(1, 3)
        self.table.setHorizontalHeaderLabels(['K1', 'K2', 'K3'])
        self.table.setColumnWidth(0, 120)
        self.table.setColumnWidth(1, 120)
        self.table.setColumnWidth(2, 120)
        main_layout.addWidget(self.table)
        
        # Initialiser les cellules vides pour pouvoir mettre la validation
        for col in range(3):
            self.table.setItem(0, col, QTableWidgetItem(""))

        # Ajouter des infobulles pour les colonnes du tableau CIR
        self.table.horizontalHeaderItem(0).setToolTip("Coefficient pour les charges directes associées aux temps de travail (≥ 1)")
        self.table.horizontalHeaderItem(1).setToolTip("Coefficient pour les charges directes associées aux amortissements (≥ 1)")
        self.table.horizontalHeaderItem(2).setToolTip("Taux de \"subvention\" CIR (entre 0 et 1)")

        # Boutons
        btn_layout = QHBoxLayout()
        self.save_btn = QPushButton('Enregistrer')
        self.save_btn.clicked.connect(self.confirm_save)
        btn_layout.addWidget(self.save_btn)
        main_layout.addLayout(btn_layout)

        # Restreindre les valeurs des coefficients K1, K2 et K3
        for col in range(3):
            item = self.table.item(0, col)
            if item:
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight)

        self.table.cellChanged.connect(self.validate_coefficients)
        self.table.cellChanged.connect(self.validate_and_mark_dirty)
        self.setLayout(main_layout)
        self.current_year = self.year_spin.value()
        self.load_data()
        
    def validate_coefficients(self, row, column):
        if column in [0, 1, 2]:
            item = self.table.item(row, column)
            if item and item.text():
                try:
                    value = float(item.text().replace(',', '.'))
                    
                    # Validation spécifique selon la colonne
                    if column in [0, 1]:  # K1 et K2
                        if value < 1:
                            raise ValueError("K1 et K2 doivent être >= 1")
                    elif column == 2:  # K3
                        if not (0 <= value <= 1):
                            raise ValueError("K3 doit être entre 0 et 1")
                    
                    # Limiter à deux chiffres après la virgule
                    item.setText(f"{value:.2f}")
                except ValueError as e:
                    if "K1 et K2" in str(e):
                        QMessageBox.warning(self, "Erreur", "K1 et K2 doivent être des nombres supérieurs ou égaux à 1.")
                    elif "K3 doit être" in str(e):
                        QMessageBox.warning(self, "Erreur", "K3 doit être un nombre entre 0 et 1.")
                    else:
                        QMessageBox.warning(self, "Erreur", "Veuillez entrer un nombre valide.")
                    self.table.blockSignals(True)
                    item.setText("")
                    self.table.blockSignals(False)

    def validate_and_mark_dirty(self, row, column):
        # Ne pas traiter les changements pendant le chargement des données
        if getattr(self, '_loading', False):
            return
            
        # Vérifier que la valeur entrée est un nombre
        item = self.table.item(row, column)
        if item and item.text():
            try:
                float(item.text().replace(',', '.'))
                # Si c'est un nombre valide, marquer comme modifié
                self._dirty = True
            except ValueError:
                # Si ce n'est pas un nombre valide, rejeter la valeur
                QMessageBox.warning(self, "Erreur", "Veuillez entrer uniquement des nombres pour les coefficients.")
                self._loading = True  # Éviter récursion
                item.setText("")
                self._loading = False

    def ensure_table_exists(self):
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS cir_coeffs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            annee INTEGER,
            k1 REAL,
            k2 REAL,
            k3 REAL
        )''')
        conn.commit()
        conn.close()

    def get_default_year(self):
        import datetime
        return datetime.datetime.now().year

    def change_year(self):
        self.brouillons[self.current_year] = self.get_table_values()
        self.current_year = self.year_spin.value()
        self.set_table_values(self.brouillons.get(self.current_year, None))

    def get_table_values(self):
        values = {}
        for j in range(3):
            item = self.table.item(0, j)
            values[j] = item.text() if item and item.text() else ''
        return values

    def set_table_values(self, values):
        self._loading = True
        if values:
            for j in range(3):
                self.table.setItem(0, j, QTableWidgetItem(values.get(j, '')))
        else:
            year = self.current_year
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('''SELECT k1, k2, k3 FROM cir_coeffs WHERE annee=?''', (year,))
            res = cursor.fetchone()
            for j in range(3):
                self.table.setItem(0, j, QTableWidgetItem(''))
            if res:
                self.table.setItem(0, 0, QTableWidgetItem('' if res[0] is None else str(res[0])))
                self.table.setItem(0, 1, QTableWidgetItem('' if res[1] is None else str(res[1])))
                self.table.setItem(0, 2, QTableWidgetItem('' if res[2] is None else str(res[2])))
            conn.close()
        self._loading = False

    def mark_dirty(self, row, column):
        if not getattr(self, '_loading', False) and column in [0, 1, 2]:
            self._dirty = True

    def confirm_save(self):
        reply = QMessageBox.question(self, 'Confirmation', 'Voulez-vous enregistrer les coefficients CIR ?',
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.save_data(show_message=True)
            self._dirty = False
            self.accept()

    def save_data(self, show_message=True):
        # Sauvegarder les valeurs actuelles dans les brouillons avant de les enregistrer
        self.brouillons[self.current_year] = self.get_table_values()

        conn = get_connection()
        cursor = conn.cursor()
        # Sauvegarder tous les brouillons (toutes les années modifiées)
        for year, values in self.brouillons.items():
            cursor.execute('SELECT id FROM cir_coeffs WHERE annee=?', (year,))
            res = cursor.fetchone()
            k1_val = None
            k2_val = None
            k3_val = None
            try:
                k1_val = float(values.get(0, '').replace(',', '.')) if values.get(0, '').strip() else None
            except Exception:
                pass
            try:
                k2_val = float(values.get(1, '').replace(',', '.')) if values.get(1, '').strip() else None
            except Exception:
                pass
            try:
                k3_val = float(values.get(2, '').replace(',', '.')) if values.get(2, '').strip() else None
            except Exception:
                pass
            if res:
                sql = "UPDATE cir_coeffs SET k1=?, k2=?, k3=? WHERE id=?"
                cursor.execute(sql, [k1_val, k2_val, k3_val, res[0]])
            else:
                sql = "INSERT INTO cir_coeffs (annee, k1, k2, k3) VALUES (?, ?, ?, ?)"
                cursor.execute(sql, [year, k1_val, k2_val, k3_val])

        conn.commit()
        conn.close()
        if show_message:
            QMessageBox.information(self, 'Sauvegarde', 'Les coefficients CIR ont été enregistrés avec succès.')

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
