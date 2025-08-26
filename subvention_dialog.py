from PyQt6.QtWidgets import QDialog, QFormLayout, QCheckBox, QDoubleSpinBox, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QLineEdit

class SubventionDialog(QDialog):
    def __init__(self, parent=None, data=None):
        super().__init__(parent)
        self.setWindowTitle('Ajouter/Modifier une subvention')
        layout = QFormLayout()
        # Nom de la subvention
        self.nom_edit = QLineEdit()
        self.nom_edit.setPlaceholderText('Ex: ADEME, Région, Europe...')
        layout.addRow('Nom de la subvention:', self.nom_edit)
        # Dépenses éligibles
        self.cb_temps = QCheckBox('Temps de travail')
        self.spin_temps = QDoubleSpinBox()
        self.spin_temps.setValue(1)
        self.spin_temps.setDecimals(2)
        self.spin_temps.setRange(0, 1)  # Coefficient entre 0 et 1
        self.spin_temps.setSingleStep(0.1)  # Pas de 0.1
        h_temps = QHBoxLayout()
        h_temps.addWidget(self.cb_temps)
        h_temps.addWidget(QLabel('Coef:'))
        h_temps.addWidget(self.spin_temps)
        layout.addRow(h_temps)
        self.cb_externes = QCheckBox('Dépenses externes')
        self.spin_externes = QDoubleSpinBox()
        self.spin_externes.setValue(1)
        self.spin_externes.setDecimals(2)
        self.spin_externes.setRange(0, 1)  # Coefficient entre 0 et 1
        self.spin_externes.setSingleStep(0.1)  # Pas de 0.1
        h_externes = QHBoxLayout()
        h_externes.addWidget(self.cb_externes)
        h_externes.addWidget(QLabel('Coef:'))
        h_externes.addWidget(self.spin_externes)
        layout.addRow(h_externes)
        self.cb_autres = QCheckBox('Autres dépenses')
        self.spin_autres = QDoubleSpinBox()
        self.spin_autres.setValue(1)
        self.spin_autres.setDecimals(2)
        self.spin_autres.setRange(0, 1)  # Coefficient entre 0 et 1
        self.spin_autres.setSingleStep(0.1)  # Pas de 0.1
        h_autres = QHBoxLayout()
        h_autres.addWidget(self.cb_autres)
        h_autres.addWidget(QLabel('Coef:'))
        h_autres.addWidget(self.spin_autres)
        layout.addRow(h_autres)
        self.cb_dotation = QCheckBox('Dotation amortissements')
        self.spin_dotation = QDoubleSpinBox()
        self.spin_dotation.setValue(1)
        self.spin_dotation.setDecimals(2)
        self.spin_dotation.setRange(0, 1)  # Coefficient entre 0 et 1
        self.spin_dotation.setSingleStep(0.1)  # Pas de 0.1
        h_dotation = QHBoxLayout()
        h_dotation.addWidget(self.cb_dotation)
        h_dotation.addWidget(QLabel('Coef:'))
        h_dotation.addWidget(self.spin_dotation)
        layout.addRow(h_dotation)
        # Cd et taux - avec QDoubleSpinBox
        self.cd_spin = QDoubleSpinBox()
        self.cd_spin.setValue(1)
        self.cd_spin.setDecimals(2)
        self.cd_spin.setRange(0, 1)  # Coefficient entre 0 et 1
        self.cd_spin.setSingleStep(0.1)  # Pas de 0.1
        layout.addRow('Cd :', self.cd_spin)
        self.taux_spin = QDoubleSpinBox()
        self.taux_spin.setValue(100)  # Valeur par défaut à 50%
        self.taux_spin.setDecimals(2)
        self.taux_spin.setRange(0, 100)
        self.taux_spin.setSuffix('%')
        layout.addRow('Taux de subvention:', self.taux_spin)
        # Boutons
        btns = QHBoxLayout()
        btn_ok = QPushButton('Valider')
        btn_cancel = QPushButton('Annuler')
        btn_ok.clicked.connect(self.validate_and_accept)
        btn_cancel.clicked.connect(self.reject)
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)
        layout.addRow(btns)
        self.setLayout(layout)
        if data:
            self.load_data(data)

    def validate_and_accept(self):
        nom = self.nom_edit.text().strip()
        if not nom:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(self, 'Erreur', "Le nom de la subvention est obligatoire.")
            return
        self.accept()
    def get_data(self):
        return {
            'nom': self.nom_edit.text().strip(),
            'depenses_temps_travail': int(self.cb_temps.isChecked()),
            'coef_temps_travail': float(self.spin_temps.value()),
            'depenses_externes': int(self.cb_externes.isChecked()),
            'coef_externes': float(self.spin_externes.value()),
            'depenses_autres_achats': int(self.cb_autres.isChecked()),
            'coef_autres_achats': float(self.spin_autres.value()),
            'depenses_dotation_amortissements': int(self.cb_dotation.isChecked()),
            'coef_dotation_amortissements': float(self.spin_dotation.value()),
            'cd': float(self.cd_spin.value()),
            'taux': float(self.taux_spin.value())
        }
    def load_data(self, data):
        self.nom_edit.setText(data.get('nom', ''))
        self.cb_temps.setChecked(bool(data.get('depenses_temps_travail', 0)))
        self.spin_temps.setValue(float(data.get('coef_temps_travail', 1)))
        self.cb_externes.setChecked(bool(data.get('depenses_externes', 0)))
        self.spin_externes.setValue(float(data.get('coef_externes', 1)))
        self.cb_autres.setChecked(bool(data.get('depenses_autres_achats', 0)))
        self.spin_autres.setValue(float(data.get('coef_autres_achats', 1)))
        self.cb_dotation.setChecked(bool(data.get('depenses_dotation_amortissements', 0)))
        self.spin_dotation.setValue(float(data.get('coef_dotation_amortissements', 1)))
        self.cd_spin.setValue(float(data.get('cd', 1)))
        self.taux_spin.setValue(float(data.get('taux', 50)))
