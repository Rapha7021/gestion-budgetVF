from PyQt6.QtWidgets import QDialog, QFormLayout, QCheckBox, QDoubleSpinBox, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QLineEdit, QFrame, QDateEdit
from PyQt6.QtCore import QDate
import sqlite3
import datetime
from utils import format_montant

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
        
        # Dates de la subvention
        dates_layout = QHBoxLayout()
        self.date_debut_subv = QDateEdit()
        self.date_debut_subv.setDisplayFormat('MM/yyyy')
        self.date_debut_subv.setDate(QDate.currentDate())
        dates_layout.addWidget(QLabel('Début:'))
        dates_layout.addWidget(self.date_debut_subv)
        dates_layout.addWidget(QLabel('Fin:'))
        self.date_fin_subv = QDateEdit()
        self.date_fin_subv.setDisplayFormat('MM/yyyy')
        self.date_fin_subv.setDate(QDate.currentDate())
        dates_layout.addWidget(self.date_fin_subv)
        layout.addRow('Période de la subvention:', dates_layout)
        
        # Case à cocher pour basculer en mode simplifié
        self.mode_simplifie_cb = QCheckBox('Mode simplifié (montant forfaitaire)')
        layout.addRow(self.mode_simplifie_cb)
        
        # Container pour les éléments du mode détaillé (masquables)
        self.mode_detaille_widget = QFrame()
        mode_detaille_layout = QFormLayout()
        
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
        mode_detaille_layout.addRow(h_temps)
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
        mode_detaille_layout.addRow(h_externes)
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
        mode_detaille_layout.addRow(h_autres)
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
        mode_detaille_layout.addRow(h_dotation)
        # Cd et taux - avec QDoubleSpinBox
        self.cd_spin = QDoubleSpinBox()
        self.cd_spin.setValue(1)
        self.cd_spin.setDecimals(2)
        self.cd_spin.setRange(1, 2)  # Coefficient entre 1 et 2
        self.cd_spin.setSingleStep(0.01)  # Pas de 0.01
        mode_detaille_layout.addRow('Coef de charge :', self.cd_spin)
        self.taux_spin = QDoubleSpinBox()
        self.taux_spin.setValue(round(100, 2))  # Valeur par défaut arrondie à 2 décimales
        self.taux_spin.setDecimals(2)
        self.taux_spin.setRange(0, 100)
        self.taux_spin.setSuffix('%')
        mode_detaille_layout.addRow('Taux de subvention:', self.taux_spin)
        
        self.mode_detaille_widget.setLayout(mode_detaille_layout)
        layout.addRow(self.mode_detaille_widget)
        
        # Container pour le mode simplifié (masqué par défaut)
        self.mode_simplifie_widget = QFrame()
        mode_simplifie_layout = QFormLayout()
        
        self.montant_forfaitaire_spin = QDoubleSpinBox()
        self.montant_forfaitaire_spin.setDecimals(2)
        self.montant_forfaitaire_spin.setRange(0, 1_000_000_000)
        self.montant_forfaitaire_spin.setSingleStep(100)
        self.montant_forfaitaire_spin.setValue(0)
        mode_simplifie_layout.addRow('Montant forfaitaire (€):', self.montant_forfaitaire_spin)
        
        # Affichage du taux calculé automatiquement
        self.taux_calcule_label = QLabel("0.00%")
        self.taux_calcule_label.setStyleSheet("font-weight: bold; font-size: 12px; color: orange;")
        mode_simplifie_layout.addRow('Taux de subvention calculé:', self.taux_calcule_label)
        
        self.mode_simplifie_widget.setLayout(mode_simplifie_layout)
        self.mode_simplifie_widget.setVisible(False)  # Masqué par défaut
        layout.addRow(self.mode_simplifie_widget)
        
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
        # Label dynamique selon le mode
        self.montant_titre_label = QLabel("Montant estimé de la subvention:")
        montant_layout.addWidget(self.montant_titre_label)
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
        self.montant_forfaitaire_spin.valueChanged.connect(self.update_montant)
        
        # Connecter le changement de mode
        self.mode_simplifie_cb.stateChanged.connect(self.toggle_mode)
        
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
        
        # Mettre à jour le taux calculé quand le montant forfaitaire change
        self.montant_forfaitaire_spin.valueChanged.connect(self.update_taux_calcule)

        # Flag pour éviter les confirmations lors du chargement des données
        self.loading_data = False
        
        # Définir les dates par défaut du projet
        self.set_default_dates_from_project()
        
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

    def toggle_mode(self):
        """Bascule entre le mode détaillé et le mode simplifié avec confirmation"""
        from PyQt6.QtWidgets import QMessageBox
        
        # Si on est en train de charger des données, pas de confirmation
        if self.loading_data:
            # Appliquer directement le changement
            mode_simplifie_demande = self.mode_simplifie_cb.isChecked()
            
            if mode_simplifie_demande:
                self.mode_detaille_widget.setVisible(False)
                self.mode_simplifie_widget.setVisible(True)
                self.montant_titre_label.setText("Montant forfaitaire de la subvention:")
            else:
                self.mode_detaille_widget.setVisible(True)
                self.mode_simplifie_widget.setVisible(False)
                self.montant_titre_label.setText("Montant estimé de la subvention:")
            
            # Mettre à jour l'affichage
            self.update_montant()
            self.update_assiette()
            return
        
        # Déterminer le mode actuel et le mode cible
        mode_simplifie_demande = self.mode_simplifie_cb.isChecked()
        
        if mode_simplifie_demande:
            # Passage en mode simplifié
            message = ("Vous allez basculer en mode simplifié.\n\n"
                      "Les paramètres détaillés (coefficients, critères) seront masqués "
                      "mais conservés.\n\n"
                      "Voulez-vous continuer ?")
            titre = "Basculer en mode simplifié"
        else:
            # Passage en mode détaillé
            message = ("Vous allez basculer en mode détaillé.\n\n"
                      "Le montant forfaitaire sera masqué mais conservé.\n\n"
                      "Voulez-vous continuer ?")
            titre = "Basculer en mode détaillé"
        
        # Afficher la boîte de confirmation
        confirmation = QMessageBox.question(
            self, 
            titre, 
            message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No  # Non par défaut pour éviter les accidents
        )
        
        if confirmation == QMessageBox.StandardButton.No:
            # L'utilisateur a annulé, remettre la case dans l'état précédent
            self.mode_simplifie_cb.blockSignals(True)  # Éviter la récursion
            self.mode_simplifie_cb.setChecked(not mode_simplifie_demande)
            self.mode_simplifie_cb.blockSignals(False)
            return
        
        # L'utilisateur a confirmé, procéder au changement
        if mode_simplifie_demande:
            # Mode simplifié activé
            self.mode_detaille_widget.setVisible(False)
            self.mode_simplifie_widget.setVisible(True)
            self.montant_titre_label.setText("Montant forfaitaire de la subvention:")
        else:
            # Mode détaillé activé
            self.mode_detaille_widget.setVisible(True)
            self.mode_simplifie_widget.setVisible(False)
            self.montant_titre_label.setText("Montant estimé de la subvention:")
        
        # Mettre à jour l'affichage
        self.update_montant()
        self.update_assiette()

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
        if self.mode_simplifie_cb.isChecked():
            # Mode simplifié : utiliser le montant forfaitaire
            montant = self.montant_forfaitaire_spin.value()
            self.montant_label.setText(format_montant(montant))
            self.montant_label.setToolTip("Montant forfaitaire saisi manuellement")
            
            # Calculer le taux automatiquement
            self.update_taux_calcule()
            return
        
        # Mode détaillé : calcul avec les coefficients
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
            detail_lines.append(f"Temps de travail: {format_montant(projet_data['temps_travail_total'])} x {self.cd_spin.value():.2f} x {self.spin_temps.value():.2f} = {format_montant(montant_temps)}")
            sub_total += montant_temps
            
        if self.cb_externes.isChecked():
            montant_externes = projet_data['depenses_externes'] * self.spin_externes.value()
            detail_lines.append(f"Dépenses externes: {format_montant(projet_data['depenses_externes'])} x {self.spin_externes.value():.2f} = {format_montant(montant_externes)}")
            sub_total += montant_externes
            
        if self.cb_autres.isChecked():
            montant_autres = projet_data['autres_achats'] * self.spin_autres.value()
            detail_lines.append(f"Autres achats: {format_montant(projet_data['autres_achats'])} x {self.spin_autres.value():.2f} = {format_montant(montant_autres)}")
            sub_total += montant_autres
            
        if self.cb_dotation.isChecked():
            montant_amort = projet_data['amortissements'] * self.spin_dotation.value()
            detail_lines.append(f"Amortissements: {format_montant(projet_data['amortissements'])} x {self.spin_dotation.value():.2f} = {format_montant(montant_amort)}")
            sub_total += montant_amort
        
        # Ajouter la ligne de sous-total et le calcul final
        detail_lines.append(f"Sous-total: {format_montant(sub_total)} x {self.taux_spin.value():.0f}% = {format_montant(montant)}")
        
        # Joindre toutes les lignes en une seule chaîne
        detail_text = "\n".join(detail_lines)
        
        # Mettre à jour le label avec formatage du montant
        self.montant_label.setText(format_montant(montant))
        
        # Ajouter une infobulle pour voir le détail du calcul
        self.montant_label.setToolTip(detail_text)
        
    def update_taux_calcule(self):
        """Calcule et affiche le taux de subvention en mode simplifié"""
        if not self.mode_simplifie_cb.isChecked():
            return
            
        # Récupérer les données du projet
        projet_data = self.get_project_data()
        
        # Calculer l'assiette totale
        assiette_totale = (projet_data['temps_travail_total'] + 
                          projet_data['depenses_externes'] + 
                          projet_data['autres_achats'] + 
                          projet_data['amortissements'])
        
        montant_forfaitaire = self.montant_forfaitaire_spin.value()
        
        # Calculer le taux : (Montant forfaitaire / Assiette éligible) × 100
        if assiette_totale > 0:
            taux_calcule = (montant_forfaitaire / assiette_totale) * 100
            self.taux_calcule_label.setText(f"{taux_calcule:.2f}%")
            
            # Infobulle avec le détail du calcul
            detail_calcul = f"Calcul: {format_montant(montant_forfaitaire)} ÷ {format_montant(assiette_totale)} × 100 = {taux_calcule:.2f}%"
            self.taux_calcule_label.setToolTip(detail_calcul)
        else:
            self.taux_calcule_label.setText("0.00%")
            self.taux_calcule_label.setToolTip("Aucune assiette éligible disponible")
        
    def update_assiette(self):
        """Met à jour l'assiette éligible selon le mode."""
        # Récupérer les données du projet une seule fois pour optimiser
        projet_data = self.get_project_data()
        
        if self.mode_simplifie_cb.isChecked():
            # Mode simplifié : Assiette éligible max = total des dépenses + valorisation salaires chargés
            assiette = (projet_data['temps_travail_total'] + 
                       projet_data['depenses_externes'] + 
                       projet_data['autres_achats'] + 
                       projet_data['amortissements'])
            # Mettre à jour aussi le taux calculé
            self.update_taux_calcule()
        else:
            # Mode détaillé : calcul avec les coefficients comme avant
            assiette = 0
            if self.cb_temps.isChecked():
                # Appliquer le coefficient de charge au temps de travail, comme dans update_montant()
                temps_travail_avec_cd = projet_data['temps_travail_total'] * self.cd_spin.value()
                assiette += self.spin_temps.value() * temps_travail_avec_cd
            if self.cb_externes.isChecked():
                assiette += self.spin_externes.value() * projet_data['depenses_externes']
            if self.cb_autres.isChecked():
                assiette += self.spin_autres.value() * projet_data['autres_achats']
            if self.cb_dotation.isChecked():
                assiette += self.spin_dotation.value() * projet_data['amortissements']
        
        self.assiette_label.setText(format_montant(assiette))
        
    @staticmethod
    def calculate_distributed_subvention(project_id, subvention_data, target_year, target_month=None):
        """
        Calcule la subvention répartie proportionnellement aux dépenses éligibles de la période.
        
        Args:
            project_id: ID du projet
            subvention_data: Dictionnaire contenant toutes les données de la subvention
            target_year: Année cible
            target_month: Mois cible (optionnel, si None = toute l'année)
            
        Returns:
            Montant de la subvention pour la période demandée
        """
        if not project_id or not subvention_data:
            return 0
            
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        
        try:
            # 1. Récupérer les dates de début et fin de la subvention
            date_debut_subv = subvention_data.get('date_debut_subvention')
            date_fin_subv = subvention_data.get('date_fin_subvention')
            
            if not date_debut_subv or not date_fin_subv:
                # Si pas de dates de subvention, utiliser les dates du projet
                cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (project_id,))
                projet_dates = cursor.fetchone()
                if not projet_dates or not projet_dates[0] or not projet_dates[1]:
                    return 0
                date_debut_subv, date_fin_subv = projet_dates[0], projet_dates[1]
            
            # 2. Vérifier si la période cible est dans la période de subvention
            try:
                debut_subv = datetime.datetime.strptime(date_debut_subv, '%m/%Y')
                fin_subv = datetime.datetime.strptime(date_fin_subv, '%m/%Y')
                
                if target_month:
                    # Vérification pour un mois spécifique
                    target_date = datetime.datetime(target_year, target_month, 1)
                    if target_date < debut_subv or target_date > fin_subv:
                        return 0
                else:
                    # Vérification pour une année
                    year_start = datetime.datetime(target_year, 1, 1)
                    year_end = datetime.datetime(target_year, 12, 1)
                    if year_end < debut_subv or year_start > fin_subv:
                        return 0
                        
            except ValueError:
                return 0
            
            # 3. Calculer le montant total de la subvention selon le mode
            if subvention_data.get('mode_simplifie', 0):
                montant_total_subvention = float(subvention_data.get('montant_forfaitaire', 0))
            else:
                # Mode détaillé : calcul avec coefficients
                montant_total_subvention = SubventionDialog._calculate_detailed_subvention_amount(
                    cursor, project_id, subvention_data
                )
            
            if montant_total_subvention <= 0:
                return 0
            
            # 4. Calculer les dépenses éligibles totales de la subvention
            depenses_eligibles_totales = SubventionDialog._calculate_total_eligible_expenses(
                cursor, project_id, subvention_data, debut_subv, fin_subv
            )
            
            if depenses_eligibles_totales <= 0:
                return 0
            
            # 5. Calculer les dépenses éligibles de la période cible
            depenses_eligibles_periode = SubventionDialog._calculate_period_eligible_expenses(
                cursor, project_id, subvention_data, target_year, target_month
            )
            
            # 6. Calculer la proportion et retourner le montant réparti
            proportion = depenses_eligibles_periode / depenses_eligibles_totales
            montant_reparti = montant_total_subvention * proportion
            
            return montant_reparti
            
        except Exception as e:
            print(f"Erreur dans calculate_distributed_subvention: {e}")
            return 0
        finally:
            conn.close()
    
    @staticmethod
    def _calculate_detailed_subvention_amount(cursor, project_id, subvention_data):
        """Calcule le montant total de subvention en mode détaillé"""
        # Récupérer les données projet comme dans update_montant
        from main import get_equipe_categories
        
        # Code simplifié de get_project_data pour calculer les totaux
        # (logique similaire à celle dans get_project_data mais statique)
        
        data = {
            'temps_travail_total': 0,
            'depenses_externes': 0,
            'autres_achats': 0,
            'amortissements': 0
        }
        
        # Calcul du temps de travail total
        cursor.execute("""
            SELECT tt.annee, tt.categorie, tt.mois, tt.jours 
            FROM temps_travail tt 
            WHERE tt.projet_id = ?
        """, (project_id,))
        
        temps_travail_rows = cursor.fetchall()
        cout_total_temps = 0
        
        for annee, categorie, mois, jours in temps_travail_rows:
            # Récupérer directement par le libellé au lieu d'utiliser des codes
            cursor.execute("""
                SELECT montant_charge 
                FROM categorie_cout 
                WHERE libelle = ? AND annee = ?
            """, (categorie, annee))
            
            cout_row = cursor.fetchone()
            if cout_row and cout_row[0]:
                cout_total_temps += float(jours) * float(cout_row[0])
        
        data['temps_travail_total'] = cout_total_temps
        
        # Autres dépenses
        cursor.execute("SELECT SUM(montant) FROM depenses WHERE projet_id = ?", (project_id,))
        depenses_row = cursor.fetchone()
        if depenses_row and depenses_row[0]:
            data['depenses_externes'] = float(depenses_row[0])
        
        cursor.execute("SELECT SUM(montant) FROM autres_depenses WHERE projet_id = ?", (project_id,))
        autres_depenses_row = cursor.fetchone()
        if autres_depenses_row and autres_depenses_row[0]:
            data['autres_achats'] = float(autres_depenses_row[0])
        
        # Amortissements (logique simplifiée)
        cursor.execute("SELECT montant, date_achat, duree FROM investissements WHERE projet_id = ?", (project_id,))
        amortissements_total = 0
        for montant, date_achat, duree in cursor.fetchall():
            try:
                amortissements_total += float(montant)
            except Exception:
                pass
        data['amortissements'] = amortissements_total
        
        # Calcul du montant avec coefficients
        montant = 0
        
        if subvention_data.get('depenses_temps_travail', 0):
            temps_travail = data['temps_travail_total'] * subvention_data.get('cd', 1)
            montant += subvention_data.get('coef_temps_travail', 1) * temps_travail
        
        if subvention_data.get('depenses_externes', 0):
            montant += subvention_data.get('coef_externes', 1) * data['depenses_externes']
        
        if subvention_data.get('depenses_autres_achats', 0):
            montant += subvention_data.get('coef_autres_achats', 1) * data['autres_achats']
        
        if subvention_data.get('depenses_dotation_amortissements', 0):
            montant += subvention_data.get('coef_dotation_amortissements', 1) * data['amortissements']
        
        # Appliquer le taux de subvention
        montant = montant * (subvention_data.get('taux', 100) / 100)
        
        return montant
    
    @staticmethod
    def _calculate_total_eligible_expenses(cursor, project_id, subvention_data, debut_subv, fin_subv):
        """Calcule les dépenses éligibles totales sur toute la période du projet (pas seulement la période de subvention)"""
        # Il faut prendre TOUTES les dépenses éligibles du projet pour faire une répartition correcte
        # Sinon on crée un biais si la période de subvention ne couvre qu'une partie du projet
        
        # Récupérer les dates du projet
        cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (project_id,))
        projet_dates = cursor.fetchone()
        if not projet_dates or not projet_dates[0] or not projet_dates[1]:
            return 0
            
        try:
            debut_projet = datetime.datetime.strptime(projet_dates[0], '%m/%Y')
            fin_projet = datetime.datetime.strptime(projet_dates[1], '%m/%Y')
        except ValueError:
            return 0
        
        # Calculer les dépenses éligibles sur TOUTE la période du projet
        return SubventionDialog._calculate_period_eligible_expenses_range(
            cursor, project_id, subvention_data, debut_projet, fin_projet
        )
    
    @staticmethod 
    def _calculate_period_eligible_expenses(cursor, project_id, subvention_data, target_year, target_month):
        """Calcule les dépenses éligibles pour une période spécifique"""
        depenses_periode = 0
        
        # Condition de filtrage selon la période
        if target_month:
            month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                          "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
            where_condition_tt = "AND tt.annee = ? AND tt.mois = ?"
            where_condition_dep = "AND annee = ? AND mois = ?"
            params_tt = [project_id, target_year, month_names[target_month-1]]
            params_dep = [project_id, target_year, month_names[target_month-1]]
        else:
            where_condition_tt = "AND tt.annee = ?"
            where_condition_dep = "AND annee = ?"
            params_tt = [project_id, target_year]
            params_dep = [project_id, target_year]
        
        # Temps de travail si coché
        if subvention_data.get('depenses_temps_travail', 0):
            cursor.execute(f"""
                SELECT SUM(tt.jours * cc.montant_charge)
                FROM temps_travail tt
                JOIN categorie_cout cc ON cc.libelle = tt.categorie AND cc.annee = tt.annee
                WHERE tt.projet_id = ? {where_condition_tt}
            """, params_tt)
            
            result = cursor.fetchone()
            if result and result[0]:
                # Appliquer les coefficients après le calcul brut
                montant_brut = float(result[0])
                montant_avec_cd = montant_brut * subvention_data.get('cd', 1)
                montant_final = montant_avec_cd * subvention_data.get('coef_temps_travail', 1)
                depenses_periode += montant_final
        
        # Dépenses externes si cochées
        if subvention_data.get('depenses_externes', 0):
            cursor.execute(f"""
                SELECT SUM(montant)
                FROM depenses 
                WHERE projet_id = ? {where_condition_dep}
            """, params_dep)
            
            result = cursor.fetchone()
            if result and result[0]:
                # Appliquer le coefficient après le calcul brut
                montant_brut = float(result[0])
                montant_final = montant_brut * subvention_data.get('coef_externes', 1)
                depenses_periode += montant_final
        
        # Autres achats si cochés
        if subvention_data.get('depenses_autres_achats', 0):
            cursor.execute(f"""
                SELECT SUM(montant)
                FROM autres_depenses 
                WHERE projet_id = ? {where_condition_dep}
            """, params_dep)
            
            result = cursor.fetchone()
            if result and result[0]:
                # Appliquer le coefficient après le calcul brut
                montant_brut = float(result[0])
                montant_final = montant_brut * subvention_data.get('coef_autres_achats', 1)
                depenses_periode += montant_final
            
            result = cursor.fetchone()
            if result and result[0]:
                depenses_periode += float(result[0])
        
        # Amortissements si cochés (calcul simplifié pour la période)
        if subvention_data.get('depenses_dotation_amortissements', 0):
            # Logique simplifiée : répartition équitable sur la durée
            cursor.execute("""
                SELECT SUM(montant) FROM investissements WHERE projet_id = ?
            """, (project_id,))
            
            result = cursor.fetchone()
            if result and result[0]:
                # Approximation : diviser par le nombre total de mois puis multiplier par la période
                amort_brut = float(result[0])
                # Appliquer le coefficient après le calcul brut
                amort_total = amort_brut * subvention_data.get('coef_dotation_amortissements', 1)
                # Pour simplifier, on répartit de façon équitable
                if target_month:
                    # Pour un mois, prendre 1/12 de l'année
                    cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (project_id,))
                    projet_dates = cursor.fetchone()
                    if projet_dates and projet_dates[0] and projet_dates[1]:
                        try:
                            debut = datetime.datetime.strptime(projet_dates[0], '%m/%Y')
                            fin = datetime.datetime.strptime(projet_dates[1], '%m/%Y')
                            nb_mois_total = (fin.year - debut.year) * 12 + (fin.month - debut.month) + 1
                            if nb_mois_total > 0:
                                depenses_periode += amort_total / nb_mois_total
                        except ValueError:
                            pass
                else:
                    # Pour une année, prendre la part annuelle
                    depenses_periode += amort_total / 12  # Approximation simple
        
        return depenses_periode
    
    @staticmethod
    def _calculate_period_eligible_expenses_range(cursor, project_id, subvention_data, debut_date, fin_date):
        """Calcule les dépenses éligibles sur une plage de dates"""
        # Pour simplifier, on fait la somme des mois dans la plage
        total = 0
        current_date = debut_date.replace(day=1)
        
        while current_date <= fin_date:
            montant_mois = SubventionDialog._calculate_period_eligible_expenses(
                cursor, project_id, subvention_data, current_date.year, current_date.month
            )
            total += montant_mois
            
            # Passer au mois suivant
            if current_date.month == 12:
                current_date = current_date.replace(year=current_date.year + 1, month=1)
            else:
                current_date = current_date.replace(month=current_date.month + 1)
        
        return total
        
    def validate_and_accept(self):
        # Vérifier que le nom est renseigné
        nom = self.nom_edit.text().strip()
        if not nom:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(self, 'Erreur', "Le nom de la subvention est obligatoire.")
            return
        
        if self.mode_simplifie_cb.isChecked():
            # Mode simplifié : vérifier que le montant forfaitaire est > 0
            if self.montant_forfaitaire_spin.value() <= 0:
                from PyQt6.QtWidgets import QMessageBox
                QMessageBox.warning(self, 'Erreur', "Veuillez saisir un montant forfaitaire supérieur à 0.")
                return
        else:
            # Mode détaillé : vérifier qu'au moins un critère est coché
            if not (self.cb_temps.isChecked() or self.cb_externes.isChecked() or 
                    self.cb_autres.isChecked() or self.cb_dotation.isChecked()):
                from PyQt6.QtWidgets import QMessageBox
                QMessageBox.warning(self, 'Erreur', "Veuillez sélectionner au moins un critère pour la subvention.")
                return
                
        self.accept()
    def get_data(self):
        return {
            'nom': self.nom_edit.text().strip(),
            'date_debut_subvention': self.date_debut_subv.date().toString('MM/yyyy'),
            'date_fin_subvention': self.date_fin_subv.date().toString('MM/yyyy'),
            'mode_simplifie': int(self.mode_simplifie_cb.isChecked()),
            'montant_forfaitaire': float(self.montant_forfaitaire_spin.value()),
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
        # Activer le flag pour éviter les confirmations lors du chargement
        self.loading_data = True
        
        self.nom_edit.setText(data.get('nom', ''))
        
        # Charger les dates avec gestion des anciennes BDD
        if 'date_debut_subvention' in data and data['date_debut_subvention']:
            try:
                debut_date = datetime.datetime.strptime(data['date_debut_subvention'], '%m/%Y')
                self.date_debut_subv.setDate(QDate(debut_date.year, debut_date.month, 1))
            except (ValueError, TypeError):
                # Si erreur, utiliser les dates du projet par défaut
                self.set_default_dates_from_project()
        else:
            # Anciennes BDD : utiliser les dates du projet
            self.set_default_dates_from_project()
            
        if 'date_fin_subvention' in data and data['date_fin_subvention']:
            try:
                fin_date = datetime.datetime.strptime(data['date_fin_subvention'], '%m/%Y')
                self.date_fin_subv.setDate(QDate(fin_date.year, fin_date.month, 1))
            except (ValueError, TypeError):
                # Si erreur, utiliser les dates du projet par défaut
                self.set_default_dates_from_project()
        else:
            # Anciennes BDD : utiliser les dates du projet
            self.set_default_dates_from_project()
        
        self.mode_simplifie_cb.setChecked(bool(data.get('mode_simplifie', 0)))
        self.montant_forfaitaire_spin.setValue(float(data.get('montant_forfaitaire', 0)))
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
        
        # Mettre à jour l'affichage des widgets après le chargement
        self.toggle_mode()
        
        # Désactiver le flag après le chargement
        self.loading_data = False
        
    def set_default_dates_from_project(self):
        """Définit les dates de subvention par défaut à partir des dates du projet"""
        if not self.projet_id:
            return
            
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (self.projet_id,))
            result = cursor.fetchone()
            
            if result and result[0] and result[1]:
                try:
                    # Dates du projet au format MM/yyyy
                    debut_projet = datetime.datetime.strptime(result[0], '%m/%Y')
                    fin_projet = datetime.datetime.strptime(result[1], '%m/%Y')
                    
                    # Définir les dates de subvention = dates du projet
                    self.date_debut_subv.setDate(QDate(debut_projet.year, debut_projet.month, 1))
                    self.date_fin_subv.setDate(QDate(fin_projet.year, fin_projet.month, 1))
                    
                except ValueError:
                    # Si erreur de format, garder les dates actuelles
                    pass
        except Exception:
            # En cas d'erreur, garder les dates actuelles
            pass
        finally:
            conn.close()