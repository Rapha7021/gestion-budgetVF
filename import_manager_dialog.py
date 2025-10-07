from PyQt6.QtWidgets import QDialog, QVBoxLayout, QPushButton, QListWidget, QMessageBox, QHBoxLayout
import pickle

from database import get_connection

class ImportManagerDialog(QDialog):
    def __init__(self, parent, projet_id):
        super().__init__(parent)
        self.setWindowTitle('Gestion des imports Excel')
        self.projet_id = projet_id
        self.resize(800, 100)  # Agrandit la fenêtre (largeur x hauteur)
        layout = QVBoxLayout()
        self.import_list = QListWidget()
        self.import_list.setMinimumHeight(100)  # Agrandit la liste
        layout.addWidget(self.import_list)
        btn_hbox = QHBoxLayout()
        add_btn = QPushButton('Ajouter un import')
        open_btn = QPushButton('Ouvrir')
        del_btn = QPushButton('Supprimer')
        btn_hbox.addWidget(add_btn)
        btn_hbox.addWidget(open_btn)
        btn_hbox.addWidget(del_btn)
        layout.addLayout(btn_hbox)
        self.setLayout(layout)
        self.load_imports()
        add_btn.clicked.connect(self.add_import)
        open_btn.clicked.connect(self.open_import)
        del_btn.clicked.connect(self.delete_import)

    def load_imports(self):
        self.import_list.clear()
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, filename, import_date FROM imports WHERE projet_id=? ORDER BY import_date DESC', (self.projet_id,))
        for id_, filename, import_date in cursor.fetchall():
            self.import_list.addItem(f"{filename} | {import_date}")
        conn.close()

    def add_import(self):
        from excel_import import import_excel, save_import_to_db
        import shutil, os
        from PyQt6.QtWidgets import QFileDialog, QInputDialog
        # Demander le chemin du fichier importé à l'utilisateur
        file_path, _ = QFileDialog.getOpenFileName(self, "Sélectionner le fichier importé", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return
        # Importer le fichier Excel
        df = import_excel(self, file_path=file_path)
        if df is not None:
            # Copier le fichier dans le dossier 'imports'
            imports_dir = os.path.join(os.getcwd(), 'imports')
            if not os.path.exists(imports_dir):
                os.makedirs(imports_dir)
            import datetime
            base_name = os.path.basename(file_path)
            unique_name = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{base_name}"
            dest_path = os.path.join(imports_dir, unique_name)
            shutil.copy2(file_path, dest_path)
            # Demander le nom d'affichage
            filename, ok = QInputDialog.getText(self, "Nom de l'import", "Nom du fichier importé :", text=base_name)
            if ok and filename:
                save_import_to_db(self.projet_id, dest_path, df)
                QMessageBox.information(self, "Import", "Import sauvegardé !")
                self.load_imports()

    def open_import(self):
        idx = self.import_list.currentRow()
        if idx < 0:
            QMessageBox.warning(self, "Ouvrir", "Sélectionnez un import à ouvrir.")
            return
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT filename FROM imports WHERE projet_id=? ORDER BY import_date DESC', (self.projet_id,))
        rows = cursor.fetchall()
        conn.close()
        if idx >= len(rows):
            QMessageBox.warning(self, "Ouvrir", "Import introuvable.")
            return
        filename = rows[idx][0]
        import os
        try:
            os.startfile(filename)
        except Exception as e:
            QMessageBox.critical(self, "Erreur ouverture", f"Impossible d'ouvrir le fichier : {filename}\n{e}")

    def delete_import(self):
        idx = self.import_list.currentRow()
        if idx < 0:
            QMessageBox.warning(self, "Supprimer", "Sélectionnez un import à supprimer.")
            return
        # Confirmation avant suppression
        confirm = QMessageBox.question(self, "Confirmation", "Voulez-vous vraiment supprimer cet import et le fichier associé ?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm != QMessageBox.StandardButton.Yes:
            return
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, filename FROM imports WHERE projet_id=? ORDER BY import_date DESC', (self.projet_id,))
        rows = cursor.fetchall()
        if idx >= len(rows):
            QMessageBox.warning(self, "Supprimer", "Import introuvable.")
            conn.close()
            return
        import_id, file_path = rows[idx]
        # Supprimer le fichier physique
        import os
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
        except Exception:
            pass
        cursor.execute('DELETE FROM imports WHERE id=?', (import_id,))
        conn.commit()
        conn.close()
        self.load_imports()
