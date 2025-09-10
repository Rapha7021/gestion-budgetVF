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
        
        # Mettre à jour l'assiette et montant quand les dates de subvention changent
        self.date_debut_subv.dateChanged.connect(self.update_montant)
        self.date_debut_subv.dateChanged.connect(self.update_assiette)
        self.date_fin_subv.dateChanged.connect(self.update_montant)
        self.date_fin_subv.dateChanged.connect(self.update_assiette)

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
        """Récupère les données du projet pour calculer le montant de la subvention sur la période de subvention"""
        if not self.projet_id:
            return {
                'temps_travail_total': 0,
                'depenses_externes': 0,
                'autres_achats': 0,
                'amortissements': 0
            }
         
        conn = sqlite3.connect('gestion_budget.db')
        cursor = conn.cursor()
        
        # Récupérer les dates de subvention si définies
        date_debut_subv = self.date_debut_subv.date().toString('MM/yyyy') if hasattr(self, 'date_debut_subv') else None
        date_fin_subv = self.date_fin_subv.date().toString('MM/yyyy') if hasattr(self, 'date_fin_subv') else None
        
        if not date_debut_subv or not date_fin_subv:
            # Si pas de dates de subvention, utiliser les dates du projet
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (self.projet_id,))
            date_row = cursor.fetchone()
            if not date_row or not date_row[0] or not date_row[1]:
                conn.close()
                return {
                    'temps_travail_total': 0,
                    'depenses_externes': 0,
                    'autres_achats': 0,
                    'amortissements': 0
                }
            date_debut_subv, date_fin_subv = date_row[0], date_row[1]
        
        try:
            debut_subv = datetime.datetime.strptime(date_debut_subv, '%m/%Y')
            fin_subv = datetime.datetime.strptime(date_fin_subv, '%m/%Y')
        except ValueError:
            conn.close()
            return {
                'temps_travail_total': 0,
                'depenses_externes': 0,
                'autres_achats': 0,
                'amortissements': 0
            }
        
        # Récupérer les vraies données détaillées du projet pour la période de subvention UNIQUEMENT
        temps_travail_total = 0
        depenses_externes = 0
        autres_achats = 0
        amortissements = 0
        
        # Fonction helper pour vérifier si un mois/année est dans la période de subvention
        def is_in_subvention_period(annee, mois):
            try:
                # Convertir le mois français en numéro
                mois_mapping = {
                    'Janvier': 1, 'Février': 2, 'Mars': 3, 'Avril': 4, 'Mai': 5, 'Juin': 6,
                    'Juillet': 7, 'Août': 8, 'Septembre': 9, 'Octobre': 10, 'Novembre': 11, 'Décembre': 12
                }
                if mois not in mois_mapping:
                    return False
                    
                mois_num = mois_mapping[mois]
                date_entry = datetime.datetime(int(annee), mois_num, 1)
                
                # Vérifier si cette date est dans la période de subvention
                return debut_subv <= date_entry <= fin_subv
            except:
                return False
        
        # Temps de travail avec valorisation pour la période de subvention
        cursor.execute("""
            SELECT tt.annee, tt.mois, tt.jours, cc.montant_charge
            FROM temps_travail tt
            JOIN categorie_cout cc ON cc.libelle = tt.categorie AND cc.annee = tt.annee
            WHERE tt.projet_id = ?
        """, (self.projet_id,))
        temps_results = cursor.fetchall()
        
        for annee, mois, jours, montant_charge in temps_results:
            if is_in_subvention_period(annee, mois):
                temps_travail_total += jours * montant_charge
        
        # Dépenses externes pour la période de subvention
        cursor.execute("""
            SELECT annee, mois, montant
            FROM depenses 
            WHERE projet_id = ?
        """, (self.projet_id,))
        depenses_results = cursor.fetchall()
        
        for annee, mois, montant in depenses_results:
            if is_in_subvention_period(annee, mois):
                depenses_externes += montant
        
        # Autres achats (autres dépenses) pour la période de subvention
        cursor.execute("""
            SELECT annee, mois, montant
            FROM autres_depenses 
            WHERE projet_id = ?
        """, (self.projet_id,))
        autres_results = cursor.fetchall()
        
        for annee, mois, montant in autres_results:
            if is_in_subvention_period(annee, mois):
                autres_achats += montant
        
        # Amortissements (investissements) pour la période de subvention
        # Note: Les investissements n'ont pas de répartition mensuelle, on les répartit sur toute la durée du projet
        cursor.execute("""
            SELECT montant, date_achat, duree
            FROM investissements 
            WHERE projet_id = ?
        """, (self.projet_id,))
        invest_results = cursor.fetchall()
        
        # Pour les investissements, on calcule une répartition mensuelle sur la durée du projet
        # et on ne prend que la part qui correspond à la période de subvention
        if invest_results:
            # Calculer la durée totale du projet en mois
            try:
                duree_projet_mois = (fin_subv.year - debut_subv.year) * 12 + (fin_subv.month - debut_subv.month) + 1
                if duree_projet_mois > 0:
                    for montant, date_achat, duree in invest_results:
                        if montant:
                            # Répartir l'investissement sur la période de subvention
                            # (simplification : on considère que l'amortissement se fait sur la période de subvention)
                            amortissements += float(montant)
            except:
                pass
        
        data = {
            'temps_travail_total': temps_travail_total,
            'depenses_externes': depenses_externes,
            'autres_achats': autres_achats,
            'amortissements': amortissements
        }
        
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
        
        # Calculer l'assiette éligible pour l'infobulle
        assiette_eligible = 0
        if self.cb_temps.isChecked():
            temps_travail = projet_data['temps_travail_total'] * self.cd_spin.value()
            assiette_eligible += self.spin_temps.value() * temps_travail
        if self.cb_externes.isChecked():
            assiette_eligible += self.spin_externes.value() * projet_data['depenses_externes']
        if self.cb_autres.isChecked():
            assiette_eligible += self.spin_autres.value() * projet_data['autres_achats']
        if self.cb_dotation.isChecked():
            assiette_eligible += self.spin_dotation.value() * projet_data['amortissements']
        
        # Appliquer le taux de subvention
        montant = assiette_eligible * (self.taux_spin.value() / 100)
        
        # Mettre à jour le label avec formatage du montant
        self.montant_label.setText(format_montant(montant))
        
        # Infobulle simplifiée : Assiette éligible x Taux = Montant
        detail_text = f"Assiette éligible: {format_montant(assiette_eligible)} × {self.taux_spin.value():.0f}% = {format_montant(montant)}"
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
            
            # Créer l'infobulle pour le mode simplifié
            detail_assiette = []
            detail_assiette.append(f"Temps de travail: {format_montant(projet_data['temps_travail_total'])}")
            detail_assiette.append(f"Dépenses externes: {format_montant(projet_data['depenses_externes'])}")
            detail_assiette.append(f"Autres achats: {format_montant(projet_data['autres_achats'])}")
            detail_assiette.append(f"Amortissements: {format_montant(projet_data['amortissements'])}")
            detail_assiette.append(f"TOTAL: {format_montant(assiette)}")
            
            detail_text = "\n".join(detail_assiette)
            self.assiette_label.setToolTip(f"Détail de l'assiette éligible:\n{detail_text}")
            
            # Mettre à jour aussi le taux calculé
            self.update_taux_calcule()
        else:
            # Mode détaillé : calcul avec les coefficients comme avant
            assiette = 0
            detail_assiette = []
            
            if self.cb_temps.isChecked():
                # Appliquer le coefficient de charge au temps de travail, comme dans update_montant()
                temps_travail_avec_cd = projet_data['temps_travail_total'] * self.cd_spin.value()
                temps_eligible = self.spin_temps.value() * temps_travail_avec_cd
                assiette += temps_eligible
                detail_assiette.append(f"Temps de travail: {format_montant(projet_data['temps_travail_total'])} × {self.cd_spin.value():.2f} × {self.spin_temps.value():.2f} = {format_montant(temps_eligible)}")
            
            if self.cb_externes.isChecked():
                externes_eligible = self.spin_externes.value() * projet_data['depenses_externes']
                assiette += externes_eligible
                detail_assiette.append(f"Dépenses externes: {format_montant(projet_data['depenses_externes'])} × {self.spin_externes.value():.2f} = {format_montant(externes_eligible)}")
            
            if self.cb_autres.isChecked():
                autres_eligible = self.spin_autres.value() * projet_data['autres_achats']
                assiette += autres_eligible
                detail_assiette.append(f"Autres achats: {format_montant(projet_data['autres_achats'])} × {self.spin_autres.value():.2f} = {format_montant(autres_eligible)}")
            
            if self.cb_dotation.isChecked():
                amort_eligible = self.spin_dotation.value() * projet_data['amortissements']
                assiette += amort_eligible
                detail_assiette.append(f"Amortissements: {format_montant(projet_data['amortissements'])} × {self.spin_dotation.value():.2f} = {format_montant(amort_eligible)}")
            
            # Ajouter le total
            if detail_assiette:
                detail_assiette.append(f"TOTAL ASSIETTE: {format_montant(assiette)}")
                detail_text = "\n".join(detail_assiette)
                self.assiette_label.setToolTip(f"Détail du calcul d'assiette éligible:\n{detail_text}")
            else:
                self.assiette_label.setToolTip("Aucune catégorie de dépenses sélectionnée")
        
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
        """Calcule le montant total de subvention en mode détaillé sur la période de subvention uniquement"""
        
        # Récupérer les dates de début et fin de la subvention
        date_debut_subv = subvention_data.get('date_debut_subvention')
        date_fin_subv = subvention_data.get('date_fin_subvention')
        
        if not date_debut_subv or not date_fin_subv:
            # Si pas de dates de subvention, utiliser les dates du projet
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (project_id,))
            projet_dates = cursor.fetchone()
            if not projet_dates or not projet_dates[0] or not projet_dates[1]:
                return 0
            date_debut_subv, date_fin_subv = projet_dates[0], projet_dates[1]
        
        try:
            debut_subv = datetime.datetime.strptime(date_debut_subv, '%m/%Y')
            fin_subv = datetime.datetime.strptime(date_fin_subv, '%m/%Y')
        except ValueError:
            return 0
        
        # Calculer les dépenses éligibles sur la période de subvention seulement
        depenses_eligibles_totales = SubventionDialog._calculate_period_eligible_expenses_range(
            cursor, project_id, subvention_data, debut_subv, fin_subv
        )
        
        # Appliquer le taux de subvention
        montant = depenses_eligibles_totales * (subvention_data.get('taux', 100) / 100)
        
        return montant
    
    @staticmethod
    def _calculate_total_eligible_expenses(cursor, project_id, subvention_data, debut_subv, fin_subv):
        """Calcule les dépenses éligibles totales sur la période de subvention uniquement"""
        
        # Calculer les dépenses éligibles sur la période de subvention
        return SubventionDialog._calculate_period_eligible_expenses_range(
            cursor, project_id, subvention_data, debut_subv, fin_subv
        )
    
    @staticmethod 
    def _calculate_period_eligible_expenses(cursor, project_id, subvention_data, target_year, target_month):
        """Calcule les dépenses éligibles pour une période spécifique, limitée à la période de subvention"""
        depenses_periode = 0
        
        # Récupérer les dates de début et fin de la subvention
        date_debut_subv = subvention_data.get('date_debut_subvention')
        date_fin_subv = subvention_data.get('date_fin_subvention')
        
        if not date_debut_subv or not date_fin_subv:
            # Si pas de dates de subvention, utiliser les dates du projet
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (project_id,))
            projet_dates = cursor.fetchone()
            if not projet_dates or not projet_dates[0] or not projet_dates[1]:
                return 0
            date_debut_subv, date_fin_subv = projet_dates[0], projet_dates[1]
        
        try:
            debut_subv = datetime.datetime.strptime(date_debut_subv, '%m/%Y')
            fin_subv = datetime.datetime.strptime(date_fin_subv, '%m/%Y')
        except ValueError:
            return 0
        
        # Déterminer si on est dans la période de subvention
        if target_month:
            # Cas d'un mois spécifique
            target_date = datetime.datetime(target_year, target_month, 1)
            if target_date < debut_subv or target_date > fin_subv:
                return 0  # Le mois n'est pas dans la période de subvention
        else:
            # Cas d'une année complète : vérifier qu'au moins un mois est couvert
            mois_couverts = []
            for mois in range(1, 13):
                mois_date = datetime.datetime(target_year, mois, 1)
                if debut_subv <= mois_date <= fin_subv:
                    mois_couverts.append(mois)
            
            if not mois_couverts:
                return 0  # Aucun mois de l'année n'est couvert par la subvention
        
        month_names = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                      "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        
        # CALCUL MENSUEL (un seul mois)
        if target_month:
            mois_nom = month_names[target_month - 1]
            
            # 1. Temps de travail pour ce mois
            if subvention_data.get('depenses_temps_travail', 0):
                cursor.execute("""
                    SELECT SUM(tt.jours * cc.montant_charge)
                    FROM temps_travail tt
                    JOIN categorie_cout cc ON cc.libelle = tt.categorie AND cc.annee = tt.annee
                    WHERE tt.projet_id = ? AND tt.annee = ? AND tt.mois = ?
                """, (project_id, target_year, mois_nom))
                
                result = cursor.fetchone()
                if result and result[0]:
                    montant_brut = float(result[0])
                    montant_avec_cd = montant_brut * subvention_data.get('cd', 1)
                    montant_final = montant_avec_cd * subvention_data.get('coef_temps_travail', 1)
                    depenses_periode += montant_final
            
            # 2. Dépenses externes pour ce mois (avec redistribution automatique)
            if subvention_data.get('depenses_externes', 0):
                montant_reparti = SubventionDialog._get_redistributed_monthly_amount(
                    cursor, project_id, target_year, target_month, 'depenses', debut_subv, fin_subv
                )
                if montant_reparti > 0:
                    # Utiliser le montant réparti
                    montant_final = montant_reparti * subvention_data.get('coef_externes', 1)
                    depenses_periode += montant_final
                else:
                    # Calcul normal pour ce mois
                    cursor.execute("""
                        SELECT SUM(montant)
                        FROM depenses 
                        WHERE projet_id = ? AND annee = ? AND mois = ?
                    """, (project_id, target_year, mois_nom))
                    
                    result = cursor.fetchone()
                    if result and result[0]:
                        montant_brut = float(result[0])
                        montant_final = montant_brut * subvention_data.get('coef_externes', 1)
                        depenses_periode += montant_final
            
            # 3. Autres achats pour ce mois (avec redistribution automatique)
            if subvention_data.get('depenses_autres_achats', 0):
                montant_reparti = SubventionDialog._get_redistributed_monthly_amount(
                    cursor, project_id, target_year, target_month, 'autres_depenses', debut_subv, fin_subv
                )
                if montant_reparti > 0:
                    # Utiliser le montant réparti
                    montant_final = montant_reparti * subvention_data.get('coef_autres_achats', 1)
                    depenses_periode += montant_final
                else:
                    # Calcul normal pour ce mois
                    cursor.execute("""
                        SELECT SUM(montant)
                        FROM autres_depenses 
                        WHERE projet_id = ? AND annee = ? AND mois = ?
                    """, (project_id, target_year, mois_nom))
                    
                    result = cursor.fetchone()
                    if result and result[0]:
                        montant_brut = float(result[0])
                        montant_final = montant_brut * subvention_data.get('coef_autres_achats', 1)
                        depenses_periode += montant_final
        
        # CALCUL ANNUEL (tous les mois de l'année dans la période de subvention)
        else:
            # 1. Temps de travail pour tous les mois couverts
            if subvention_data.get('depenses_temps_travail', 0):
                for mois in mois_couverts:
                    mois_nom = month_names[mois - 1]
                    
                    cursor.execute("""
                        SELECT SUM(tt.jours * cc.montant_charge)
                        FROM temps_travail tt
                        JOIN categorie_cout cc ON cc.libelle = tt.categorie AND cc.annee = tt.annee
                        WHERE tt.projet_id = ? AND tt.annee = ? AND tt.mois = ?
                    """, (project_id, target_year, mois_nom))
                    
                    result = cursor.fetchone()
                    if result and result[0]:
                        montant_brut = float(result[0])
                        montant_avec_cd = montant_brut * subvention_data.get('cd', 1)
                        montant_final = montant_avec_cd * subvention_data.get('coef_temps_travail', 1)
                        depenses_periode += montant_final
            
            # 2. Dépenses externes pour toute l'année (avec redistribution automatique)
            if subvention_data.get('depenses_externes', 0):
                montant_redistribue = SubventionDialog._check_and_redistribute_single_expense(
                    cursor, project_id, target_year, 'depenses', debut_subv, fin_subv, 
                    subvention_data.get('coef_externes', 1)
                )
                depenses_periode += montant_redistribue
            
            # 3. Autres achats pour toute l'année (avec redistribution automatique)
            if subvention_data.get('depenses_autres_achats', 0):
                montant_redistribue = SubventionDialog._check_and_redistribute_single_expense(
                    cursor, project_id, target_year, 'autres_depenses', debut_subv, fin_subv, 
                    subvention_data.get('coef_autres_achats', 1)
                )
                depenses_periode += montant_redistribue
        
        # Amortissements si cochés (calcul au prorata de la période de subvention)
        if subvention_data.get('depenses_dotation_amortissements', 0):
            # Récupérer le total des amortissements pour le projet
            cursor.execute("""
                SELECT SUM(montant) FROM investissements WHERE projet_id = ?
            """, (project_id,))
            
            result = cursor.fetchone()
            if result and result[0]:
                amort_brut = float(result[0])
                # Appliquer le coefficient
                amort_total = amort_brut * subvention_data.get('coef_dotation_amortissements', 1)
                
                # Calculer la part d'amortissement pour les mois couverts par la subvention
                # Récupérer les dates du projet pour calculer le nombre total de mois
                cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (project_id,))
                projet_dates = cursor.fetchone()
                if projet_dates and projet_dates[0] and projet_dates[1]:
                    try:
                        debut_projet = datetime.datetime.strptime(projet_dates[0], '%m/%Y')
                        fin_projet = datetime.datetime.strptime(projet_dates[1], '%m/%Y')
                        nb_mois_total_projet = (fin_projet.year - debut_projet.year) * 12 + (fin_projet.month - debut_projet.month) + 1
                        
                        if nb_mois_total_projet > 0:
                            # Part mensuelle des amortissements
                            amort_mensuel = amort_total / nb_mois_total_projet
                            if target_month:
                                # Cas mensuel : ajouter l'amortissement pour 1 mois
                                depenses_periode += amort_mensuel
                            else:
                                # Cas annuel : ajouter l'amortissement pour tous les mois couverts
                                depenses_periode += amort_mensuel * len(mois_couverts)
                    except ValueError:
                        pass
        
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
    
    @staticmethod
    def _check_and_redistribute_single_expense(cursor, project_id, target_year, table_name, debut_subv, fin_subv, coefficient):
        """
        Vérifie s'il n'y a qu'une seule dépense dans l'année pour le projet.
        Si c'est le cas, la répartit sur tous les mois actifs du projet dans cette année.
        
        Args:
            cursor: Curseur de base de données
            project_id: ID du projet
            target_year: Année cible
            table_name: Nom de la table ('depenses' ou 'autres_depenses')
            debut_subv: Date de début de subvention
            fin_subv: Date de fin de subvention
            coefficient: Coefficient à appliquer
            
        Returns:
            Montant réparti sur la période, ou 0 si pas de répartition
        """
        # Récupérer toutes les dépenses de l'année pour ce projet
        cursor.execute(f"""
            SELECT mois, SUM(montant) as total_montant
            FROM {table_name}
            WHERE projet_id = ? AND annee = ?
            GROUP BY mois
        """, (project_id, target_year))
        
        depenses_par_mois = cursor.fetchall()
        
        # Si pas de dépenses ou plus d'une entrée par mois, pas de répartition automatique
        if len(depenses_par_mois) != 1:
            return 0
        
        # Il y a exactement une dépense dans l'année
        mois_unique, montant_total = depenses_par_mois[0]
        
        # Calculer les mois actifs du projet dans cette année (limités par la période de subvention)
        mois_actifs = []
        for mois in range(1, 13):
            mois_date = datetime.datetime(target_year, mois, 1)
            if debut_subv <= mois_date <= fin_subv:
                mois_actifs.append(mois)
        
        if not mois_actifs:
            return 0
        
        # Répartir le montant sur tous les mois actifs
        montant_par_mois = montant_total / len(mois_actifs)
        montant_total_reparti = montant_par_mois * len(mois_actifs)
        
        # Appliquer le coefficient
        return montant_total_reparti * coefficient
    
    @staticmethod
    def _get_redistributed_monthly_amount(cursor, project_id, target_year, target_month, table_name, debut_subv, fin_subv):
        """
        Retourne le montant mensuel réparti si une dépense unique est détectée dans l'année.
        Sinon retourne 0 (calcul normal).
        
        Args:
            cursor: Curseur de base de données
            project_id: ID du projet
            target_year: Année cible
            target_month: Mois cible (1-12)
            table_name: Nom de la table ('depenses' ou 'autres_depenses')
            debut_subv: Date de début de subvention
            fin_subv: Date de fin de subvention
            
        Returns:
            Montant mensuel réparti, ou 0 si pas de répartition
        """
        # Récupérer toutes les dépenses de l'année pour ce projet
        cursor.execute(f"""
            SELECT mois, SUM(montant) as total_montant
            FROM {table_name}
            WHERE projet_id = ? AND annee = ?
            GROUP BY mois
        """, (project_id, target_year))
        
        depenses_par_mois = cursor.fetchall()
        
        # Si pas de dépenses ou plus d'une entrée par mois, pas de répartition automatique
        if len(depenses_par_mois) != 1:
            return 0
        
        # Il y a exactement une dépense dans l'année
        mois_unique, montant_total = depenses_par_mois[0]
        
        # Calculer les mois actifs du projet dans cette année (limités par la période de subvention)
        mois_actifs = []
        for mois in range(1, 13):
            mois_date = datetime.datetime(target_year, mois, 1)
            if debut_subv <= mois_date <= fin_subv:
                mois_actifs.append(mois)
        
        if not mois_actifs or target_month not in mois_actifs:
            return 0
        
        # Vérifier que le mois cible est dans la période active
        target_date = datetime.datetime(target_year, target_month, 1)
        if not (debut_subv <= target_date <= fin_subv):
            return 0
        
        # Répartir le montant sur tous les mois actifs et retourner la part du mois demandé
        montant_par_mois = montant_total / len(mois_actifs)
        return montant_par_mois
    
    def get_projet_dates(self):
        """Récupère les dates de début et de fin du projet"""
        if not self.projet_id:
            return None, None
            
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
                    return debut_projet, fin_projet
                except ValueError:
                    return None, None
            return None, None
        except Exception:
            return None, None
        finally:
            conn.close()
        
    def validate_and_accept(self):
        from PyQt6.QtWidgets import QMessageBox
        
        # Vérifier que le nom est renseigné
        nom = self.nom_edit.text().strip()
        if not nom:
            QMessageBox.warning(self, 'Erreur', "Le nom de la subvention est obligatoire.")
            return

        # Validation des dates de subvention
        date_debut_subv = self.date_debut_subv.date().toPyDate()
        date_fin_subv = self.date_fin_subv.date().toPyDate()
        
        # Vérifier que la date de début est antérieure à la date de fin
        if date_debut_subv >= date_fin_subv:
            QMessageBox.warning(self, 'Erreur', "La date de début de la subvention doit être antérieure à la date de fin.")
            return
        
        # Vérifier que les dates de subvention sont comprises dans les dates du projet
        debut_projet, fin_projet = self.get_projet_dates()
        if debut_projet and fin_projet:
            # Convertir les dates de subvention en datetime pour la comparaison
            debut_subv_dt = datetime.datetime(date_debut_subv.year, date_debut_subv.month, 1)
            fin_subv_dt = datetime.datetime(date_fin_subv.year, date_fin_subv.month, 1)
            
            if debut_subv_dt < debut_projet:
                QMessageBox.warning(self, 'Erreur', 
                    f"La date de début de la subvention ({date_debut_subv.strftime('%m/%Y')}) "
                    f"ne peut pas être antérieure à la date de début du projet ({debut_projet.strftime('%m/%Y')}).")
                return
                
            if fin_subv_dt > fin_projet:
                QMessageBox.warning(self, 'Erreur', 
                    f"La date de fin de la subvention ({date_fin_subv.strftime('%m/%Y')}) "
                    f"ne peut pas être postérieure à la date de fin du projet ({fin_projet.strftime('%m/%Y')}).")
                return

        if self.mode_simplifie_cb.isChecked():
            # Mode simplifié : vérifier que le montant forfaitaire est > 0
            if self.montant_forfaitaire_spin.value() <= 0:
                QMessageBox.warning(self, 'Erreur', "Veuillez saisir un montant forfaitaire supérieur à 0.")
                return
        else:
            # Mode détaillé : vérifier qu'au moins un critère est coché
            if not (self.cb_temps.isChecked() or self.cb_externes.isChecked() or 
                    self.cb_autres.isChecked() or self.cb_dotation.isChecked()):
                QMessageBox.warning(self, 'Erreur', "Veuillez sélectionner au moins un critère pour la subvention.")
                return
                
        self.accept()
    
    def calculate_current_montant_estime(self):
        """Calcule le montant estimé actuel basé sur l'affichage de l'interface"""
        try:
            # Récupérer le montant affiché dans l'interface (en supprimant "€" et les espaces)
            montant_text = self.montant_label.text().replace('€', '').replace(' ', '').replace(',', '.')
            
            # Si c'est vide ou contient des caractères non numériques, retourner 0
            if not montant_text or montant_text == '0':
                return 0.0
            
            # Convertir en float
            return float(montant_text)
            
        except (ValueError, AttributeError):
            # En cas d'erreur, retourner 0
            return 0.0
    
    def calculate_current_assiette_eligible(self):
        """Calcule l'assiette éligible totale basée sur les paramètres actuels de l'interface"""
        try:
            # Si on est en mode simplifié, l'assiette éligible n'est pas pertinente
            if self.mode_simplifie_cb.isChecked():
                return 0.0
            
            # Utiliser la même logique que get_project_data() pour avoir la même période
            projet_data = self.get_project_data()
            
            # Calculer l'assiette éligible selon les critères sélectionnés
            assiette_eligible = 0
            
            if self.cb_temps.isChecked():
                temps_eligible = projet_data['temps_travail_total'] * self.cd_spin.value()
                assiette_eligible += self.spin_temps.value() * temps_eligible
                
            if self.cb_externes.isChecked():
                assiette_eligible += self.spin_externes.value() * projet_data['depenses_externes']
                
            if self.cb_autres.isChecked():
                assiette_eligible += self.spin_autres.value() * projet_data['autres_achats']
                
            if self.cb_dotation.isChecked():
                assiette_eligible += self.spin_dotation.value() * projet_data['amortissements']
            
            return assiette_eligible
                
        except (ValueError, AttributeError, Exception) as e:
            # En cas d'erreur, retourner 0
            return 0.0
    
    def get_data_for_calculation(self):
        """Récupère les données actuelles pour les calculs"""
        return {
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
            'taux': float(self.taux_spin.value())
        }
    
    def get_debut_subvention(self):
        """Récupère la date de début de subvention"""
        return self.date_debut_subv.date().toString('MM/yyyy')
    
    def get_fin_subvention(self):
        """Récupère la date de fin de subvention"""
        return self.date_fin_subv.date().toString('MM/yyyy')
    
    def get_data(self):
        # Calculer le montant estimé et l'assiette éligible actuels
        montant_estime = self.calculate_current_montant_estime()
        assiette_eligible = self.calculate_current_assiette_eligible()
        
        # Obtenir la date actuelle
        from datetime import datetime
        date_maj = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
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
            'montant_subvention_max': float(self.subvention_max_spin.value()),
            'assiette_eligible': assiette_eligible,
            'montant_estime_total': montant_estime,
            'date_derniere_maj': date_maj
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