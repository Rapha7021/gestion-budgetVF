from PyQt6.QtWidgets import QDialog, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QHBoxLayout, QMessageBox, QInputDialog, QDateEdit, QTextEdit, QDialogButtonBox, QLabel, QLineEdit, QDoubleSpinBox
from PyQt6.QtCore import QDate

from database import get_connection

class TaskManagerDialog(QDialog):
    def repartir_budget_automatiquement(self):
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, pourcentage_budget FROM taches WHERE projet_id=?', (self.projet_id,))
        rows = cursor.fetchall()
        total = sum([row[1] for row in rows if row[1] > 0])
        zeros = [row[0] for row in rows if row[1] == 0]
        reste = 100 - total
        if zeros and reste > 0:
            if len(zeros) == 1:
                cursor.execute('UPDATE taches SET pourcentage_budget=? WHERE id=?', (reste, zeros[0]))
            else:
                part = round(reste / len(zeros), 2)
                for i, task_id in enumerate(zeros):
                    # Pour la dernière tâche, ajuste pour arriver à 100% pile
                    if i == len(zeros) - 1:
                        part_finale = round(100 - (total + part * (len(zeros) - 1)), 2)
                        cursor.execute('UPDATE taches SET pourcentage_budget=? WHERE id=?', (part_finale, task_id))
                    else:
                        cursor.execute('UPDATE taches SET pourcentage_budget=? WHERE id=?', (part, task_id))
            conn.commit()
        conn.close()
        self.load_tasks()
    def __init__(self, parent, projet_id):
        super().__init__(parent)
        self.setWindowTitle('Gestion des tâches du projet')
        self.projet_id = projet_id
        self.resize(900, 500)
        layout = QVBoxLayout()
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "Nom de la tâche", "Date de début", "Date de fin", "Détails", "% du budget"])
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)  # Tableau non éditable
        layout.addWidget(self.table)
        btn_hbox = QHBoxLayout()
        add_btn = QPushButton('Ajouter une tâche')
        edit_btn = QPushButton('Modifier')
        del_btn = QPushButton('Supprimer')
        btn_hbox.addWidget(add_btn)
        btn_hbox.addWidget(edit_btn)
        btn_hbox.addWidget(del_btn)
        layout.addLayout(btn_hbox)
        self.setLayout(layout)
        self.load_tasks()
        add_btn.clicked.connect(self.add_task)
        edit_btn.clicked.connect(self.edit_task)
        del_btn.clicked.connect(self.delete_task)

    def load_tasks(self):
        self.table.setRowCount(0)
        conn = get_connection()
        cursor = conn.cursor()
        
        # Vérifier la structure actuelle de la table taches
        cursor.execute("PRAGMA table_info(taches)")
        columns = cursor.fetchall()
        column_names = [col[1] for col in columns]
        
        # Migration nécessaire si on a l'ancienne structure
        needs_migration = False
        if 'description' in column_names and 'details' not in column_names:
            needs_migration = True
        elif 'pourcentage_budget' not in column_names:
            needs_migration = True
            
        if needs_migration:
            print("Migration de la base de données en cours...")
            # Créer la nouvelle table avec la structure correcte
            cursor.execute('''CREATE TABLE IF NOT EXISTS taches_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                nom TEXT,
                date_debut TEXT,
                date_fin TEXT,
                details TEXT,
                pourcentage_budget REAL DEFAULT 0
            )''')
            
            # Copier les données existantes
            if 'description' in column_names:
                # Migration depuis l'ancienne structure avec description
                cursor.execute("""
                    INSERT INTO taches_new (id, projet_id, nom, date_debut, date_fin, details, pourcentage_budget)
                    SELECT id, projet_id, nom, date_debut, date_fin, 
                           COALESCE(description, '') as details, 0 as pourcentage_budget
                    FROM taches
                """)
            else:
                # Migration pour ajouter seulement pourcentage_budget
                cursor.execute("""
                    INSERT INTO taches_new (id, projet_id, nom, date_debut, date_fin, details, pourcentage_budget)
                    SELECT id, projet_id, nom, date_debut, date_fin, 
                           COALESCE(details, '') as details, 0 as pourcentage_budget
                    FROM taches
                """)
            
            # Supprimer l'ancienne table et renommer la nouvelle
            cursor.execute("DROP TABLE taches")
            cursor.execute("ALTER TABLE taches_new RENAME TO taches")
            conn.commit()
            print("Migration terminée avec succès.")
        
        cursor.execute('SELECT id, nom, date_debut, date_fin, details, pourcentage_budget FROM taches WHERE projet_id=?', (self.projet_id,))
        self.task_ids = []
        for row_idx, (id_, nom, date_debut, date_fin, details, pourcentage_budget) in enumerate(cursor.fetchall()):
            self.table.insertRow(row_idx)
            self.table.setItem(row_idx, 0, QTableWidgetItem(nom))
            # Display dates directly as they are stored in MM/YYYY format
            self.table.setItem(row_idx, 1, QTableWidgetItem(date_debut if date_debut else ""))
            self.table.setItem(row_idx, 2, QTableWidgetItem(date_fin if date_fin else ""))
            self.table.setItem(row_idx, 3, QTableWidgetItem(details if details else ""))
            self.table.setItem(row_idx, 4, QTableWidgetItem(str(pourcentage_budget) if pourcentage_budget is not None else ""))
            self.task_ids.append(id_)
        conn.close()
    def get_total_pourcentage_budget(self):
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT SUM(pourcentage_budget) FROM taches WHERE projet_id=?', (self.projet_id,))
        result = cursor.fetchone()
        conn.close()
        return result[0] if result[0] is not None else 0

    def add_task(self):
        dialog = AddTaskDialog(self)
        while dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            nom = data["nom"]
            date_debut = parse_date_fr(data["date_debut"])
            date_fin = parse_date_fr(data["date_fin"])
            details = data["details"]
            pourcentage_budget = data["pourcentage_budget"]

            # Check if start date is before or equal to end date
            if date_debut > date_fin:
                QMessageBox.warning(self, "Erreur de date", "La date de début doit être antérieure ou égale à la date de fin.")
                continue

            # Fetch project start and end dates
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT date_debut, date_fin FROM projets WHERE id=?', (self.projet_id,))
            projet_dates = cursor.fetchone()
            conn.close()

            if projet_dates:
                projet_date_debut, projet_date_fin = projet_dates
                if (date_debut < projet_date_debut) or (date_fin > projet_date_fin):
                    QMessageBox.warning(self, "Erreur de date", "Les dates de la tâche doivent être comprises entre {} et {}.".format(projet_date_debut, projet_date_fin))
                    continue

            total = self.get_total_pourcentage_budget()
            if total + pourcentage_budget > 100:
                QMessageBox.warning(self, "Erreur budget", "La somme des pourcentages du budget dépasse 100%. Veuillez ajuster la répartition.")
                continue

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('INSERT INTO taches (projet_id, nom, date_debut, date_fin, details, pourcentage_budget) VALUES (?, ?, ?, ?, ?, ?)',
                           (self.projet_id, nom, date_debut, date_fin, details, pourcentage_budget))
            conn.commit()
            conn.close()
            self.repartir_budget_automatiquement()
            break

    def closeEvent(self, event):
        total = self.get_total_pourcentage_budget()
        if total is None or total == 0:
            # Permet de quitter si aucune tâche n'existe
            event.accept()
        elif total < 100:
            QMessageBox.warning(self, "Erreur budget", "La somme des pourcentages du budget n'atteint pas 100%. Veuillez compléter la répartition avant de quitter.")
            event.ignore()
        else:
            event.accept()

    def delete_task(self):
        idx = self.table.currentRow()
        if idx < 0:
            QMessageBox.warning(self, "Supprimer", "Sélectionnez une tâche à supprimer.")
            return
        # Confirmation avant suppression
        confirm = QMessageBox.question(self, "Confirmation", "Voulez-vous vraiment supprimer cette tâche ?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm != QMessageBox.StandardButton.Yes:
            return
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM taches WHERE projet_id=?', (self.projet_id,))
        rows = cursor.fetchall()
        if idx >= len(rows):
            QMessageBox.warning(self, "Supprimer", "Tâche introuvable.")
            conn.close()
            return
        task_id = rows[idx][0]
        cursor.execute('DELETE FROM taches WHERE id=?', (task_id,))
        conn.commit()
        conn.close()
        self.load_tasks()

    def edit_task(self):
        idx = self.table.currentRow()
        if idx < 0:
            QMessageBox.warning(self, "Modifier", "Sélectionnez une tâche à modifier.")
            return
        nom = self.table.item(idx, 0).text()
        date_debut = self.table.item(idx, 1).text()
        date_fin = self.table.item(idx, 2).text()
        details = self.table.item(idx, 3).text()
        pourcentage_budget = self.table.item(idx, 4).text()
        dialog = AddTaskDialog(self, nom, date_debut, date_fin, details, pourcentage_budget)
        while dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_data()
            nom = data["nom"]
            date_debut = parse_date_fr(data["date_debut"])
            date_fin = parse_date_fr(data["date_fin"])
            details = data["details"]
            pourcentage_budget = data["pourcentage_budget"]

            # Check if start date is before or equal to end date
            if date_debut > date_fin:
                QMessageBox.warning(self, "Erreur de date", "La date de début doit être antérieure ou égale à la date de fin.")
                continue

            # Fetch project start and end dates
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT date_debut, date_fin FROM projets WHERE id=?', (self.projet_id,))
            projet_dates = cursor.fetchone()
            conn.close()

            if projet_dates:
                projet_date_debut, projet_date_fin = projet_dates
                if (date_debut < projet_date_debut) or (date_fin > projet_date_fin):
                    QMessageBox.warning(self, "Erreur de date", "Les dates de la tâche doivent être comprises entre {} et {}.".format(projet_date_debut, projet_date_fin))
                    continue

            task_id = self.task_ids[idx]
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('UPDATE taches SET nom=?, date_debut=?, date_fin=?, details=?, pourcentage_budget=? WHERE id=?',
                           (nom, date_debut, date_fin, details, pourcentage_budget, task_id))
            conn.commit()
            conn.close()
            self.repartir_budget_automatiquement()
            break

    def calculer_pourcentage_budget(self, date_debut, date_fin):
        # date_debut et date_fin sont au format YYYY-MM-DD
        try:
            from datetime import datetime
            if date_debut and date_fin:
                d1 = datetime.strptime(date_debut, "%Y-%m-%d")
                d2 = datetime.strptime(date_fin, "%Y-%m-%d")
                nb_jours = (d2 - d1).days + 1
                return nb_jours
        except Exception:
            pass
        return 0

class AddTaskDialog(QDialog):
    def __init__(self, parent=None, nom='', date_debut='', date_fin='', details='', pourcentage_budget=''):
        super().__init__(parent)
        self.setWindowTitle("Ajouter une tâche" if not nom else "Modifier la tâche")
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Nom de la tâche :"))
        self.nom_edit = QLineEdit()
        self.nom_edit.setText(nom)
        layout.addWidget(self.nom_edit)
        layout.addWidget(QLabel("Date de début :"))
        self.date_debut_edit = QDateEdit()
        self.date_debut_edit.setCalendarPopup(True)
        self.date_debut_edit.setDisplayFormat("MM/yyyy")
        if date_debut:
            from datetime import datetime
            try:
                d = datetime.strptime(date_debut, "%m/%Y")
                self.date_debut_edit.setDate(QDate(d.year, d.month, 1))
            except Exception:
                self.date_debut_edit.setDate(QDate.currentDate())
        else:
            self.date_debut_edit.setDate(QDate.currentDate())
        layout.addWidget(self.date_debut_edit)
        layout.addWidget(QLabel("Date de fin :"))
        self.date_fin_edit = QDateEdit()
        self.date_fin_edit.setCalendarPopup(True)
        self.date_fin_edit.setDisplayFormat("MM/yyyy")
        if date_fin:
            from datetime import datetime
            try:
                d = datetime.strptime(date_fin, "%m/%Y")
                self.date_fin_edit.setDate(QDate(d.year, d.month, 1))
            except Exception:
                self.date_fin_edit.setDate(QDate.currentDate())
        else:
            self.date_fin_edit.setDate(QDate.currentDate())
        layout.addWidget(self.date_fin_edit)
        layout.addWidget(QLabel("Détails :"))
        self.details_edit = QTextEdit()
        self.details_edit.setPlainText(details)
        layout.addWidget(self.details_edit)
        layout.addWidget(QLabel("Pourcentage du budget :"))
        self.pourcentage_edit = QDoubleSpinBox()
        self.pourcentage_edit.setRange(0, 100)
        self.pourcentage_edit.setDecimals(2)
        self.pourcentage_edit.setSuffix(" %")
        # Si le champ n'est pas renseigné, on met 0 par défaut
        try:
            val = float(pourcentage_budget)
        except Exception:
            val = 0
        self.pourcentage_edit.setValue(val)
        layout.addWidget(self.pourcentage_edit)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        layout.addWidget(buttons)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)

    def validate_and_accept(self):
        nom = self.nom_edit.text().strip()
        date_debut = self.date_debut_edit.date().toString("MM/yyyy")
        date_fin = self.date_fin_edit.date().toString("MM/yyyy")
        if not nom:
            QMessageBox.warning(self, "Champ obligatoire", "Le nom de la tâche est obligatoire.")
            return
        if not date_debut:
            QMessageBox.warning(self, "Champ obligatoire", "La date de début est obligatoire.")
            return
        if not date_fin:
            QMessageBox.warning(self, "Champ obligatoire", "La date de fin est obligatoire.")
            return
        self.accept()

    def get_data(self):
        return {
            "nom": self.nom_edit.text(),
            "date_debut": self.date_debut_edit.date().toString("MM/yyyy"),
            "date_fin": self.date_fin_edit.date().toString("MM/yyyy"),
            "details": self.details_edit.toPlainText(),
            "pourcentage_budget": self.pourcentage_edit.value()
        }

def parse_date_fr(date_str):
    # Convertit une date MM/YYYY en MM/YYYY pour le stockage
    try:
        from datetime import datetime
        d = datetime.strptime(date_str, "%m/%Y")
        return d.strftime("%m/%Y")
    except Exception:
        return ""
