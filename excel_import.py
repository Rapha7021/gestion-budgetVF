# --- Structure SQL à ajouter dans la création de la base ---
# CREATE TABLE IF NOT EXISTS imports (
#     id INTEGER PRIMARY KEY AUTOINCREMENT,
#     projet_id INTEGER,
#     filename TEXT,
#     import_date TEXT,
#     data BLOB,
#     FOREIGN KEY(projet_id) REFERENCES projets(id)
# );

import sqlite3
import os
DB_PATH = 'gestion_budget.db'

def save_import_to_db(projet_id, filename, df):
    """
    Sauvegarde l'import Excel dans la table 'imports' (données en BLOB, format pickle).
    """
    import pickle
    import datetime
    data_blob = pickle.dumps(df)
    import_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''INSERT INTO imports (projet_id, filename, import_date, data) VALUES (?, ?, ?, ?)''',
                   (projet_id, filename, import_date, data_blob))
    conn.commit()
    conn.close()
import pandas as pd
from PyQt6.QtWidgets import QFileDialog, QMessageBox


COLONNES_FIXES = ['Unité', 'Code salarié', 'Nom Salarié']
COLONNES_PAR_ANNEE = [
    'Nb jours', 'Montant chargé', 'frais de structure', 'Ku', 'Charges générales',
    'Montant coût de production', 'Véhicules de service', 'Déplace-ments'
]

def detect_annees(df):
    """
    Détecte dynamiquement toutes les années présentes dans les colonnes du fichier.
    """
    annees = set()
    for col in df.columns:
        col_str = str(col)
        for prefix in COLONNES_PAR_ANNEE:
            if col_str.startswith(prefix + ' '):
                try:
                    annee = col_str.split(' ')[-1]
                    if annee.isdigit():
                        annees.add(annee)
                except Exception:
                    pass
    return sorted(list(annees))

import pandas as pd
from PyQt6.QtWidgets import QFileDialog, QMessageBox

COLONNES_FIXES = ['Unité', 'Code salarié', 'Nom Salarié']
COLONNES_PAR_ANNEE = [
    'Nb jours', 'Montant chargé', 'frais de structure', 'Ku', 'Charges générales',
    'Montant coût de production', 'Véhicules de service', 'Déplace-ments'
]

def import_excel(parent, file_path=None):
    if file_path is None:
        file_path, _ = QFileDialog.getOpenFileName(parent, "Choisir le fichier Excel", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return None
    try:
        # Lire les 10 premières lignes pour trouver la vraie ligne d'en-tête
        preview = pd.read_excel(file_path, sheet_name='Détail sans support', header=None, nrows=10)
        # Chercher la ligne contenant les colonnes fixes
        header_row = None
        for i in range(len(preview)):
            row = preview.iloc[i].astype(str).str.lower()
            if all(any(col.lower() in cell for cell in row) for col in COLONNES_FIXES):
                header_row = i
                break
        if header_row is None:
            QMessageBox.critical(parent, "Erreur import Excel", "Impossible de trouver la ligne d'en-tête correcte.")
            return None
        # Vérifier si l'en-tête est sur deux lignes (multi-index)
        multi_index = False
        if header_row + 1 < len(preview):
            next_row = preview.iloc[header_row + 1].astype(str)
            if any(cell.strip().startswith("ANNÉE") for cell in next_row):
                multi_index = True
        if multi_index:
            df = pd.read_excel(file_path, sheet_name='Détail sans support', header=[header_row, header_row + 1])
            # Fusionner les colonnes multi-index en 'Année - NomColonne'
            new_columns = []
            annees = set()
            for i, (a, b) in enumerate(df.columns):
                annee_str = str(a).strip()
                col_str = str(b).strip()
                if annee_str.lower().startswith('année') and col_str:
                    annee_num = ''.join(filter(str.isdigit, annee_str))
                    new_columns.append(f"{annee_num} - {col_str}")
                    annees.add(annee_num)
                elif col_str in COLONNES_FIXES:
                    new_columns.append(col_str)
                else:
                    new_columns.append(f"{annee_str} - {col_str}")
            df.columns = new_columns
            annees = sorted(list(annees))
        else:
            df = pd.read_excel(file_path, sheet_name='Détail sans support', header=header_row)
            df.columns = [str(c).strip() if c and c != 'nan' else f"col_{i}" for i, c in enumerate(df.columns)]
    # (Affichage des colonnes supprimé)
        # On retourne le DataFrame et la liste des années pour la synthèse
        df._imported_years = annees if multi_index else None
        return df
    except Exception as e:
        QMessageBox.critical(parent, "Erreur import Excel", str(e))
        return None
    except Exception as e:
        QMessageBox.critical(parent, "Erreur import Excel", str(e))
        return None

