from PyQt6.QtWidgets import QDialog, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QHBoxLayout, QMessageBox, QLineEdit, QLabel, QComboBox

from database import get_connection

class ProjectManagerDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gérer les chef(fe)s de projet")
        self.setGeometry(100, 100, 600, 400)

        # Layout principal
        self.layout = QVBoxLayout()

        # Table pour afficher les chefs de projet
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Nom", "Prénom", "Direction"])
        self.layout.addWidget(self.table)

        # Boutons Ajouter, Modifier, Supprimer
        self.button_layout = QHBoxLayout()
        self.add_button = QPushButton("Ajouter")
        self.edit_button = QPushButton("Modifier")
        self.delete_button = QPushButton("Supprimer")

        self.button_layout.addWidget(self.add_button)
        self.button_layout.addWidget(self.edit_button)
        self.button_layout.addWidget(self.delete_button)
        self.layout.addLayout(self.button_layout)

        self.setLayout(self.layout)

        # Connexions des boutons
        self.add_button.clicked.connect(self.add_project_manager)
        self.edit_button.clicked.connect(self.edit_project_manager)
        self.delete_button.clicked.connect(self.delete_project_manager)

        # Charger les données initiales
        self.load_data()

    def load_data(self):
        """Charge les données des chefs de projet depuis la base de données."""
        connection = get_connection()
        cursor = connection.cursor()

        query = "SELECT nom, prenom, direction FROM chefs_projet"
        cursor.execute(query)
        rows = cursor.fetchall()

        self.table.setRowCount(0)
        for row in rows:
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            for column, data in enumerate(row):
                self.table.setItem(row_position, column, QTableWidgetItem(str(data)))

        connection.close()

    def add_project_manager(self):
        """Affiche un formulaire pour ajouter un chef de projet."""
        self.show_form()

    def edit_project_manager(self):
        """Affiche un formulaire pour modifier un chef de projet sélectionné."""
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner un chef de projet à modifier.")
            return

        nom = self.table.item(selected_row, 0).text()
        prenom = self.table.item(selected_row, 1).text()
        direction = self.table.item(selected_row, 2).text()

        self.show_form(nom, prenom, direction)

    def delete_project_manager(self):
        """Supprime le chef de projet sélectionné."""
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Erreur", "Veuillez sélectionner un chef de projet à supprimer.")
            return

        nom = self.table.item(selected_row, 0).text()
        prenom = self.table.item(selected_row, 1).text()

        confirm = QMessageBox.question(
            self,
            "Confirmation",
            f"Êtes-vous sûr de vouloir supprimer le chef de projet {nom} {prenom} ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        connection = get_connection()
        cursor = connection.cursor()

        query = "DELETE FROM chefs_projet WHERE nom = ? AND prenom = ?"
        cursor.execute(query, (nom, prenom))
        connection.commit()
        connection.close()

        self.load_data()

    def show_form(self, nom="", prenom="", direction=""):
        """Affiche un formulaire pour ajouter ou modifier un chef de projet."""
        form_dialog = QDialog(self)
        form_dialog.setWindowTitle("Formulaire Chef de Projet")
        form_layout = QVBoxLayout()

        # Champs du formulaire
        nom_label = QLabel("Nom:")
        nom_input = QLineEdit()
        nom_input.setText(nom)

        prenom_label = QLabel("Prénom:")
        prenom_input = QLineEdit()
        prenom_input.setText(prenom)

        direction_label = QLabel("Direction:")
        direction_input = QComboBox()

        # Charger les directions depuis la base de données
        connection = get_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT nom FROM directions")
        directions = cursor.fetchall()
        connection.close()

        for dir_name in directions:
            direction_input.addItem(dir_name[0])

        direction_input.setCurrentText(direction)

        # Ajouter les champs au layout
        form_layout.addWidget(nom_label)
        form_layout.addWidget(nom_input)
        form_layout.addWidget(prenom_label)
        form_layout.addWidget(prenom_input)
        form_layout.addWidget(direction_label)
        form_layout.addWidget(direction_input)

        # Bouton de validation
        save_button = QPushButton("Enregistrer")
        form_layout.addWidget(save_button)

        form_dialog.setLayout(form_layout)

        def save_data():
            new_nom = nom_input.text()
            new_prenom = prenom_input.text()
            new_direction = direction_input.currentText()

            connection = get_connection()
            cursor = connection.cursor()

            if nom and prenom:  # Modification
                query = "UPDATE chefs_projet SET nom = ?, prenom = ?, direction = ? WHERE nom = ? AND prenom = ?"
                cursor.execute(query, (new_nom, new_prenom, new_direction, nom, prenom))
            else:  # Ajout
                query = "INSERT INTO chefs_projet (nom, prenom, direction) VALUES (?, ?, ?)"
                cursor.execute(query, (new_nom, new_prenom, new_direction))

            connection.commit()
            connection.close()

            self.load_data()
            form_dialog.accept()

        save_button.clicked.connect(save_data)
        form_dialog.exec()
