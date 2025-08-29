# --- Structure SQL à ajouter dans la création de la base ---
# CREATE TABLE IF NOT EXISTS categorie_cout (
#     id INTEGER PRIMARY KEY AUTOINCREMENT,
#     annee INTEGER,
#     categorie TEXT,
#     montant_charge REAL,
#     cout_production REAL,
#     cout_complet REAL
# );
import pandas as pd
import re
import datetime
from PyQt6.QtWidgets import QFileDialog, QInputDialog, QMessageBox

def import_project_from_excel(form):
    file_path, _ = QFileDialog.getOpenFileName(form, 'Importer fichier Excel', '', 'Excel (*.xlsx *.xls)')
    if not file_path:
        return
    try:
        df = pd.read_excel(file_path, sheet_name="AGRESSO Liste des opération ac")
    except Exception as e:
        QMessageBox.warning(form, "Erreur", f"Impossible de lire le fichier : {e}")
        return
    if df.empty or "Projet" not in df.columns:
        QMessageBox.warning(form, "Erreur", "La feuille ou les colonnes attendues sont absentes.")
        return
    projets = df["Projet"].astype(str).tolist()
    choix, ok = QInputDialog.getItem(form, "Choisir un projet", "Projet à importer :", projets, editable=False)
    if not ok or not choix:
        return
    row = df[df["Projet"].astype(str) == choix].iloc[0]
    mapping = {
        "code_edit": ("Projet", str),
        "date_debut": ("Date Deb", str),
        "date_fin": ("Date Fin", str),
        "details_edit": ("Description", str)
    }
    for attr, (col, typ) in mapping.items():
        widget = getattr(form, attr)
        excel_val = typ(row[col]) if col in row else ""
        if attr in ["date_debut", "date_fin"]:
            val = excel_val
            # Conversion en MM/YYYY puis en datetime.date
            try:
                if hasattr(val, "strftime"):
                    val_str = val.strftime("%m/%Y")
                elif isinstance(val, str) and re.match(r"\d{2}/\d{2}/\d{4}", val):
                    parts = val.split("/")
                    val_str = f"{parts[1]}/{parts[2]}"
                elif isinstance(val, str) and re.match(r"\d{4}-\d{2}-\d{2}", val):
                    parts = val.split("-")
                    val_str = f"{parts[1]}/{parts[0]}"
                elif isinstance(val, str):
                    val_str = val
                else:
                    val_str = ""
                # Conversion MM/YYYY -> datetime.date
                if re.match(r"^(0[1-9]|1[0-2])/\d{4}$", val_str):
                    m, y = val_str.split("/")
                    date_obj = datetime.date(int(y), int(m), 1)
                else:
                    date_obj = widget.date()  # valeur actuelle si conversion impossible
            except Exception:
                date_obj = widget.date()
            current = widget.date()
            # Si le champ est vide (date par défaut), on remplit
            is_empty = current == datetime.date.today() or current == datetime.date(2000, 1, 1)
            if is_empty:
                widget.setDate(date_obj)
            elif current != date_obj:
                rep = QMessageBox.question(form, "Remplacer ?", f"Remplacer la date existante '{current.strftime('%m/%Y')}' par '{date_obj.strftime('%m/%Y')}' pour {col} ?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if rep == QMessageBox.StandardButton.Yes:
                    widget.setDate(date_obj)
        elif attr == "details_edit":
            current = widget.toPlainText().strip()
            if not current:
                widget.setPlainText(excel_val)
            elif current != excel_val:
                rep = QMessageBox.question(form, "Remplacer ?", f"Remplacer le détail existant par celui du fichier ?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if rep == QMessageBox.StandardButton.Yes:
                    widget.setPlainText(excel_val)
        else:
            current = widget.text().strip()
            if not current:
                widget.setText(excel_val)
            elif current != excel_val:
                rep = QMessageBox.question(form, "Remplacer ?", f"Remplacer la valeur existante '{current}' par '{excel_val}' pour {col} ?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if rep == QMessageBox.StandardButton.Yes:
                    widget.setText(excel_val)
