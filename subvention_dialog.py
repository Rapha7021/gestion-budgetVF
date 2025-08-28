from PyQt6.QtWidgets import QDialog, QFormLayout, QCheckBox, QDoubleSpinBox, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QLineEdit, QFrame
import sqlite3

class SubventionDialog(QDialog):
    def __init__(self, parent=None, data=None):
        super().__init__(parent)
        self.setWindowTitle('Ajouter/Modifier une subvention')
        layout = QFormLayout()
        
        # Récupérer l'ID du projet depuis le parent
        self.projet_id = parent.projet_id if hasattr(parent, 'projet_id') else None
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
        self.spin_temps.setSingleStep(0.01)  # Pas de 0.01
        # Ajuster les dispositions pour aligner les coefficients
        h_temps = QHBoxLayout()
        h_temps.addWidget(self.cb_temps)
        h_temps.addStretch()
        h_temps.addWidget(QLabel('Coef:'))
        h_temps.addWidget(self.spin_temps)
        layout.addRow(h_temps)
        self.cb_externes = QCheckBox('Dépenses externes')
        self.spin_externes = QDoubleSpinBox()
        self.spin_externes.setValue(1)
        self.spin_externes.setDecimals(2)
        self.spin_externes.setRange(0, 1)  # Coefficient entre 0 et 1
        self.spin_externes.setSingleStep(0.01)  # Pas de 0.01
        h_externes = QHBoxLayout()
        h_externes.addWidget(self.cb_externes)
        h_externes.addStretch()
        h_externes.addWidget(QLabel('Coef:'))
        h_externes.addWidget(self.spin_externes)
        layout.addRow(h_externes)
        self.cb_autres = QCheckBox('Autres dépenses')
        self.spin_autres = QDoubleSpinBox()
        self.spin_autres.setValue(1)
        self.spin_autres.setDecimals(2)
        self.spin_autres.setRange(0, 1)  # Coefficient entre 0 et 1
        self.spin_autres.setSingleStep(0.01)  # Pas de 0.01
        h_autres = QHBoxLayout()
        h_autres.addWidget(self.cb_autres)
        h_autres.addStretch()
        h_autres.addWidget(QLabel('Coef:'))
        h_autres.addWidget(self.spin_autres)
        layout.addRow(h_autres)
        self.cb_dotation = QCheckBox('Dotation amortissements')
        self.spin_dotation = QDoubleSpinBox()
        self.spin_dotation.setValue(1)
        self.spin_dotation.setDecimals(2)
        self.spin_dotation.setRange(0, 1)  # Coefficient entre 0 et 1
        self.spin_dotation.setSingleStep(0.01)  # Pas de 0.01
        h_dotation = QHBoxLayout()
        h_dotation.addWidget(self.cb_dotation)
        h_dotation.addStretch()
        h_dotation.addWidget(QLabel('Coef:'))
        h_dotation.addWidget(self.spin_dotation)
        layout.addRow(h_dotation)
        # Cd et taux - avec QDoubleSpinBox
        self.cd_spin = QDoubleSpinBox()
        self.cd_spin.setValue(1)
        self.cd_spin.setDecimals(2)
        self.cd_spin.setRange(1, 2)  # Coefficient entre 1 et 2
        self.cd_spin.setSingleStep(0.01)  # Pas de 0.01
        layout.addRow('Coef de charge :', self.cd_spin)
        self.taux_spin = QDoubleSpinBox()
        self.taux_spin.setValue(round(100, 2))  # Valeur par défaut arrondie à 2 décimales
        self.taux_spin.setDecimals(2)
        self.taux_spin.setRange(0, 100)
        self.taux_spin.setSuffix('%')
        layout.addRow('Taux de subvention:', self.taux_spin)
        
        # Section pour afficher le montant estimé de la subvention
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addRow(separator)

        montant_layout = QVBoxLayout()

        self.assiette_label = QLabel("0 €")
        self.assiette_label.setStyleSheet("font-weight: bold; font-size: 14px; color: blue;")
        montant_layout.addWidget(QLabel("Assiette éligible:"))
        montant_layout.addWidget(self.assiette_label)

        self.montant_label = QLabel("0 €")
        self.montant_label.setStyleSheet("font-weight: bold; font-size: 14px; color: green;")
        montant_layout.addWidget(QLabel("Montant estimé de la subvention:"))
        montant_layout.addWidget(self.montant_label)

        layout.addRow(montant_layout)
        
        # Nouvelle section pour les champs numériques
        separator_bottom = QFrame()
        separator_bottom.setFrameShape(QFrame.Shape.HLine)
        separator_bottom.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addRow(separator_bottom)

        self.depenses_max_spin = QDoubleSpinBox()
        self.depenses_max_spin.setDecimals(2)
        self.depenses_max_spin.setRange(0, 1_000_000_000)
        self.depenses_max_spin.setSingleStep(100)
        self.depenses_max_spin.setValue(0)
        layout.addRow('Assiette éligibles max (€):', self.depenses_max_spin)

        self.subvention_max_spin = QDoubleSpinBox()
        self.subvention_max_spin.setDecimals(2)
        self.subvention_max_spin.setRange(0, 1_000_000_000)
        self.subvention_max_spin.setSingleStep(100)
        self.subvention_max_spin.setValue(0)
        layout.addRow('Montant subvention max (€):', self.subvention_max_spin)
        
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
        
        # Connecter les changements pour mettre à jour le montant
        self.cb_temps.stateChanged.connect(self.update_montant)
        self.spin_temps.valueChanged.connect(self.update_montant)
        self.cb_externes.stateChanged.connect(self.update_montant)
        self.spin_externes.valueChanged.connect(self.update_montant)
        self.cb_autres.stateChanged.connect(self.update_montant)
        self.spin_autres.valueChanged.connect(self.update_montant)
        self.cb_dotation.stateChanged.connect(self.update_montant)
        self.spin_dotation.valueChanged.connect(self.update_montant)
        self.cd_spin.valueChanged.connect(self.update_montant)
        self.taux_spin.valueChanged.connect(self.update_montant)
        
        # Mettre à jour l'assiette éligible lors des changements
        self.cb_temps.stateChanged.connect(self.update_assiette)
        self.spin_temps.valueChanged.connect(self.update_assiette)
        self.cb_externes.stateChanged.connect(self.update_assiette)
        self.spin_externes.valueChanged.connect(self.update_assiette)
        self.cb_autres.stateChanged.connect(self.update_assiette)
        self.spin_autres.valueChanged.connect(self.update_assiette)
        self.cb_dotation.stateChanged.connect(self.update_assiette)
        self.spin_dotation.valueChanged.connect(self.update_assiette)
        self.cd_spin.valueChanged.connect(self.update_assiette)

        # Charger les données si modification
        if data:
            self.load_data(data)
            
        # Calculer le montant initial
        self.update_montant()
        # Calculer l'assiette initiale après l'initialisation
        self.update_assiette()

        # Ajouter des infobulles pour les coefficients
        self.spin_temps.setToolTip("Ces coefficients permettent de réduire l'assiette éligible pour chaque catégorie de dépenses, selon les règles imposées par le financeur.")
        self.spin_externes.setToolTip("Ces coefficients permettent de réduire l'assiette éligible pour chaque catégorie de dépenses, selon les règles imposées par le financeur.")
        self.spin_autres.setToolTip("Ces coefficients permettent de réduire l'assiette éligible pour chaque catégorie de dépenses, selon les règles imposées par le financeur.")
        self.spin_dotation.setToolTip("Ces coefficients permettent de réduire l'assiette éligible pour chaque catégorie de dépenses, selon les règles imposées par le financeur.")
        # Ajouter des infobulles pour les noms des coefficients
        self.cb_temps.setToolTip("Ces coefficients permettent de réduire l'assiette éligible pour chaque catégorie de dépenses, selon les règles imposées par le financeur.")
        self.cb_externes.setToolTip("Ces coefficients permettent de réduire l'assiette éligible pour chaque catégorie de dépenses, selon les règles imposées par le financeur.")
        self.cb_autres.setToolTip("Ces coefficients permettent de réduire l'assiette éligible pour chaque catégorie de dépenses, selon les règles imposées par le financeur.")
        self.cb_dotation.setToolTip("Ces coefficients permettent de réduire l'assiette éligible pour chaque catégorie de dépenses, selon les règles imposées par le financeur.")
        # Ajouter une infobulle pour le coef de charge
        self.cd_spin.setToolTip("Ce coefficient traduit le volume de charges directes assignées à la catégorie 'temps de travail'.")
        # Ajouter une infobulle pour le texte 'Coef de charge' existant
        for i in range(layout.rowCount()):
            row_label = layout.itemAt(i, QFormLayout.ItemRole.LabelRole)
            if row_label and row_label.widget() and row_label.widget().text() == 'Coef de charge :':
                row_label.widget().setToolTip("Ce coefficient traduit le volume de charges directes assignées à la catégorie 'temps de travail'.")
                break

    def get_project_data(self):
        """Récupère les données du projet pour calculer le montant de la subvention"""
        if not self.projet_id:
            return {
                'temps_travail_total': 0,
                'depenses_externes': 0,
                'autres_achats': 0,
                'amortissements': 0
            }
            
        conn = sqlite3.connect('gestion_budget.db')
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
        # Pour chaque entrée de temps_travail, on doit multiplier les jours par le montant chargé
        # de sa catégorie pour l'année correspondante
        cursor.execute("""
            SELECT tt.annee, tt.categorie, tt.mois, tt.jours 
            FROM temps_travail tt 
            WHERE tt.projet_id = ?
        """, (self.projet_id,))
        
        temps_travail_rows = cursor.fetchall()
        cout_total_temps = 0
        
        for annee, categorie, mois, jours in temps_travail_rows:
            # Convertir la catégorie du temps de travail au format de categorie_cout
            categorie_code = ""
            if "Stagiaire" in categorie:
                categorie_code = "STP"
            elif "Assistante" in categorie or "opérateur" in categorie:
                categorie_code = "AOP"
            elif "Technicien" in categorie:
                categorie_code = "TEP"
            elif "Junior" in categorie:
                categorie_code = "IJP"
            elif "Senior" in categorie:
                categorie_code = "ISP"
            elif "Expert" in categorie:
                categorie_code = "EDP"
            elif "moyen" in categorie:
                categorie_code = "MOY"
            
            # Si on n'a pas trouvé de correspondance, utiliser une valeur par défaut
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
        
        # Mettre à jour les données
        data['temps_travail_total'] = cout_total_temps
        
        # 3. Récupérer toutes les dépenses externes (toutes les dépenses)
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
        
    def update_montant(self):
        """Calcule et met à jour le montant estimé de la subvention"""
        # [coef temps de travail x (temps_travail_total x Cd) +
        #  coef dépenses externes x (dépenses externes) +
        #  coef autres achats x (autres dépenses)+
        #  coef dotation amortissements x (amortissements)] x taux de subvention
        
        # Récupérer les données du projet
        projet_data = self.get_project_data()
        
        montant = 0
        
        # Temps de travail (déjà calculé avec les montants chargés appropriés)
        if self.cb_temps.isChecked():
            temps_travail = projet_data['temps_travail_total'] * self.cd_spin.value()
            montant += self.spin_temps.value() * temps_travail
        
        # Dépenses externes
        if self.cb_externes.isChecked():
            montant += self.spin_externes.value() * projet_data['depenses_externes']
        
        # Autres achats
        if self.cb_autres.isChecked():
            montant += self.spin_autres.value() * projet_data['autres_achats']
        
        # Dotation amortissements
        if self.cb_dotation.isChecked():
            montant += self.spin_dotation.value() * projet_data['amortissements']
        
        # Appliquer le taux de subvention
        montant = montant * (self.taux_spin.value() / 100)
        
        # Créer une chaîne avec le détail du calcul pour l'infobulle
        detail_lines = []
        sub_total = 0
        
        # N'ajouter chaque composante que si elle est cochée
        if self.cb_temps.isChecked():
            temps_travail = projet_data['temps_travail_total'] * self.cd_spin.value()
            montant_temps = temps_travail * self.spin_temps.value()
            detail_lines.append(f"Temps de travail: {projet_data['temps_travail_total']:.2f} € x {self.cd_spin.value():.2f} x {self.spin_temps.value():.2f} = {montant_temps:.2f} €")
            sub_total += montant_temps
            
        if self.cb_externes.isChecked():
            montant_externes = projet_data['depenses_externes'] * self.spin_externes.value()
            detail_lines.append(f"Dépenses externes: {projet_data['depenses_externes']:.2f} € x {self.spin_externes.value():.2f} = {montant_externes:.2f} €")
            sub_total += montant_externes
            
        if self.cb_autres.isChecked():
            montant_autres = projet_data['autres_achats'] * self.spin_autres.value()
            detail_lines.append(f"Autres achats: {projet_data['autres_achats']:.2f} € x {self.spin_autres.value():.2f} = {montant_autres:.2f} €")
            sub_total += montant_autres
            
        if self.cb_dotation.isChecked():
            montant_amort = projet_data['amortissements'] * self.spin_dotation.value()
            detail_lines.append(f"Amortissements: {projet_data['amortissements']:.2f} € x {self.spin_dotation.value():.2f} = {montant_amort:.2f} €")
            sub_total += montant_amort
        
        # Ajouter la ligne de sous-total et le calcul final
        detail_lines.append(f"Sous-total: {sub_total:.2f} € x {self.taux_spin.value():.0f}% = {montant:.2f} €")
        
        # Joindre toutes les lignes en une seule chaîne
        detail_text = "\n".join(detail_lines)
        
        # Mettre à jour le label avec formatage du montant
        self.montant_label.setText(f"{montant:,.2f} €".replace(",", " ").replace(".", ","))
        
        # Ajouter une infobulle pour voir le détail du calcul
        self.montant_label.setToolTip(detail_text)
        
    def update_assiette(self):
        """Met à jour l'assiette éligible sans multiplier par le taux de subvention."""
        assiette = 0
        if self.cb_temps.isChecked():
            assiette += self.spin_temps.value() * self.get_project_data()['temps_travail_total']
        if self.cb_externes.isChecked():
            assiette += self.spin_externes.value() * self.get_project_data()['depenses_externes']
        if self.cb_autres.isChecked():
            assiette += self.spin_autres.value() * self.get_project_data()['autres_achats']
        if self.cb_dotation.isChecked():
            assiette += self.spin_dotation.value() * self.get_project_data()['amortissements']
        self.assiette_label.setText(f"{assiette:,.2f} €")
        
    def validate_and_accept(self):
        # Vérifier que le nom est renseigné
        nom = self.nom_edit.text().strip()
        if not nom:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(self, 'Erreur', "Le nom de la subvention est obligatoire.")
            return
            
        # Vérifier qu'au moins un critère est coché
        if not (self.cb_temps.isChecked() or self.cb_externes.isChecked() or 
                self.cb_autres.isChecked() or self.cb_dotation.isChecked()):
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(self, 'Erreur', "Veuillez sélectionner au moins un critère pour la subvention.")
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
            'taux': float(self.taux_spin.value()),
            'depenses_eligibles_max': float(self.depenses_max_spin.value()),
            'montant_subvention_max': float(self.subvention_max_spin.value())
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
        self.taux_spin.setValue(round(float(data.get('taux', 100)), 2))  # Arrondi explicite
        self.depenses_max_spin.setValue(float(data.get('depenses_eligibles_max', 0)))
        self.subvention_max_spin.setValue(float(data.get('montant_subvention_max', 0)))