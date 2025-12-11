import sqlite3
import datetime
from contextlib import contextmanager
from pathlib import Path
import sys
import os

# Nom du fichier de base de données stocké à la racine du projet
DB_FILENAME = "gestion_budget.db"

# Déterminer le chemin correct selon si on est en .exe ou en développement
if getattr(sys, 'frozen', False):
    # L'application est empaquetée avec PyInstaller
    # sys.executable pointe vers le .exe, on veut le dossier qui le contient
    application_path = Path(sys.executable).parent
else:
    # L'application est lancée normalement (développement)
    application_path = Path(__file__).resolve().parent

DB_FILE = application_path / DB_FILENAME
# Export d'une version chaîne utilisée par l'UI ou les boîtes de dialogue
DB_PATH = str(DB_FILE)


def _configure_connection(conn: sqlite3.Connection) -> sqlite3.Connection:
    """
    Applique les options communes à toutes les connexions SQLite.
    """
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def get_connection() -> sqlite3.Connection:
    """
    Retourne une nouvelle connexion SQLite configurée.
    L'appelant est responsable de la fermeture/commit.
    """
    return _configure_connection(sqlite3.connect(DB_PATH))


@contextmanager
def db_cursor():
    """
    Fournit un curseur prêt à l'emploi dans un contexte gérant commit/rollback.
    """
    conn = get_connection()
    try:
        cursor = conn.cursor()
        yield cursor
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    """
    Crée les tables principales et applique les migrations connues.
    """
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS themes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nom TEXT UNIQUE NOT NULL
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS projets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT NOT NULL,
                nom TEXT NOT NULL,
                details TEXT,
                date_debut TEXT,
                date_fin TEXT,
                livrables TEXT,
                chef TEXT,
                etat TEXT,
                cir INTEGER,
                subvention INTEGER,
                theme_principal TEXT
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS projet_themes (
                projet_id INTEGER,
                theme_id INTEGER,
                FOREIGN KEY(projet_id) REFERENCES projets(id),
                FOREIGN KEY(theme_id) REFERENCES themes(id),
                PRIMARY KEY (projet_id, theme_id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS images (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                nom TEXT,
                data BLOB,
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS investissements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                nom TEXT,
                montant REAL,
                date_achat TEXT,
                duree INTEGER,
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS subventions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                nom TEXT,
                mode_simplifie INTEGER DEFAULT 0,
                montant_forfaitaire REAL DEFAULT 0,
                depenses_temps_travail INTEGER,
                coef_temps_travail REAL,
                depenses_externes INTEGER,
                coef_externes REAL,
                depenses_autres_achats INTEGER,
                coef_autres_achats REAL,
                depenses_dotation_amortissements INTEGER,
                coef_dotation_amortissements REAL,
                cd REAL,
                taux REAL,
                depenses_eligibles_max REAL DEFAULT 0,
                montant_subvention_max REAL DEFAULT 0,
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )

        # Migrations incrémentales sur subventions
        for statement in (
            "ALTER TABLE subventions ADD COLUMN depenses_eligibles_max REAL DEFAULT 0",
            "ALTER TABLE subventions ADD COLUMN montant_subvention_max REAL DEFAULT 0",
            "ALTER TABLE subventions ADD COLUMN mode_simplifie INTEGER DEFAULT 0",
            "ALTER TABLE subventions ADD COLUMN montant_forfaitaire REAL DEFAULT 0",
            "ALTER TABLE subventions ADD COLUMN date_debut_subvention TEXT",
            "ALTER TABLE subventions ADD COLUMN date_fin_subvention TEXT",
            "ALTER TABLE subventions ADD COLUMN assiette_eligible REAL DEFAULT 0",
            "ALTER TABLE subventions ADD COLUMN montant_estime_total REAL DEFAULT 0",
            "ALTER TABLE subventions ADD COLUMN date_derniere_maj TEXT",
        ):
            try:
                cursor.execute(statement)
            except sqlite3.OperationalError:
                pass

        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS equipe (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                type TEXT,
                nombre INTEGER,
                direction TEXT,
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS actualites (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                message TEXT NOT NULL,
                date TEXT NOT NULL,
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS directions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nom TEXT UNIQUE NOT NULL
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS chefs_projet (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nom TEXT NOT NULL,
                prenom TEXT NOT NULL,
                direction TEXT NOT NULL,
                FOREIGN KEY(direction) REFERENCES directions(nom)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS categorie_cout (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                annee INTEGER,
                categorie TEXT,
                libelle TEXT,
                montant_charge REAL,
                cout_production REAL,
                cout_complet REAL
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS temps_travail (
                projet_id INTEGER,
                annee INTEGER,
                direction TEXT,
                categorie TEXT,
                membre_id TEXT,
                mois TEXT,
                jours REAL,
                PRIMARY KEY (projet_id, annee, membre_id, mois),
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS depenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                annee INTEGER,
                mois TEXT,
                libelle TEXT,
                montant REAL,
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS autres_depenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                annee INTEGER,
                mois TEXT,
                libelle TEXT,
                montant REAL,
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS recettes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                annee INTEGER,
                mois TEXT,
                libelle TEXT,
                montant REAL,
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS taches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                nom TEXT NOT NULL,
                description TEXT,
                date_debut TEXT,
                date_fin TEXT,
                statut TEXT,
                responsable TEXT,
                FOREIGN KEY(projet_id) REFERENCES projets(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS amortissements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                projet_id INTEGER,
                investissement_id INTEGER,
                annee INTEGER,
                mois TEXT,
                montant REAL,
                FOREIGN KEY(projet_id) REFERENCES projets(id),
                FOREIGN KEY(investissement_id) REFERENCES investissements(id)
            )'''
        )
        cursor.execute(
            '''CREATE TABLE IF NOT EXISTS cir_coeffs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                annee INTEGER UNIQUE,
                k1 REAL,
                k2 REAL,
                k3 REAL
            )'''
        )
        conn.commit()


def recalculate_all_subventions():
    """
    Recalcule toutes les valeurs dérivées des subventions pour tous les projets.
    Utile après une mise à jour de la logique de calcul.
    """
    conn = get_connection()
    cursor = conn.cursor()
    
    try:
        # Récupérer toutes les subventions
        cursor.execute("""
            SELECT id, projet_id, mode_simplifie, montant_forfaitaire,
                   depenses_temps_travail, coef_temps_travail,
                   depenses_externes, coef_externes,
                   depenses_autres_achats, coef_autres_achats,
                   depenses_dotation_amortissements, coef_dotation_amortissements,
                   cd, taux, date_debut_subvention, date_fin_subvention
            FROM subventions
        """)
        subventions = cursor.fetchall()
        
        recalculated_count = 0
        
        for subv in subventions:
            subv_id = subv[0]
            projet_id = subv[1]
            mode_simplifie = subv[2]
            montant_forfaitaire = subv[3]
            depenses_temps_travail = subv[4]
            coef_temps_travail = subv[5]
            depenses_externes = subv[6]
            coef_externes = subv[7]
            depenses_autres_achats = subv[8]
            coef_autres_achats = subv[9]
            depenses_dotation_amortissements = subv[10]
            coef_dotation_amortissements = subv[11]
            cd = subv[12]
            taux = subv[13]
            date_debut_subv = subv[14]
            date_fin_subv = subv[15]
            
            # Récupérer les dates du projet
            cursor.execute("SELECT date_debut, date_fin FROM projets WHERE id=?", (projet_id,))
            projet_row = cursor.fetchone()
            if not projet_row:
                continue
            
            date_debut_projet = projet_row[0]
            date_fin_projet = projet_row[1]
            
            # Utiliser les dates de subvention ou dates du projet par défaut
            if not date_debut_subv or not date_fin_subv:
                date_debut_subv = date_debut_projet
                date_fin_subv = date_fin_projet
            
            # Calculer les données du projet sur la période de la subvention
            projet_data = _calculate_project_data(cursor, projet_id, date_debut_subv, date_fin_subv, cd)
            
            # Calculer assiette éligible et montant estimé
            if mode_simplifie:
                # Mode simplifié : montant = montant_forfaitaire
                assiette_eligible = 0
                montant_estime = montant_forfaitaire
            else:
                # Mode détaillé : calculer l'assiette éligible
                assiette_eligible = 0
                if depenses_temps_travail:
                    temps_eligible = projet_data['temps_travail_total'] * cd
                    assiette_eligible += coef_temps_travail * temps_eligible
                if depenses_externes:
                    assiette_eligible += coef_externes * projet_data['depenses_externes']
                if depenses_autres_achats:
                    assiette_eligible += coef_autres_achats * projet_data['autres_achats']
                if depenses_dotation_amortissements:
                    assiette_eligible += coef_dotation_amortissements * projet_data['amortissements']
                
                # Montant = assiette éligible × taux
                montant_estime = assiette_eligible * (taux / 100)
            
            # Mettre à jour avec les nouvelles valeurs calculées
            date_maj = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
            
            cursor.execute("""
                UPDATE subventions 
                SET assiette_eligible = ?,
                    montant_estime_total = ?,
                    date_derniere_maj = ?
                WHERE id = ?
            """, (assiette_eligible, montant_estime, date_maj, subv_id))
            
            recalculated_count += 1
        
        conn.commit()
        return recalculated_count
        
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def _calculate_project_data(cursor, projet_id, date_debut, date_fin, cd):
    """
    Calcule les données agrégées d'un projet sur une période donnée.
    Similaire à la logique dans SubventionDialog.get_project_data()
    """
    import datetime
    
    # Parser les dates MM/yyyy
    try:
        debut_mois, debut_annee = map(int, date_debut.split('/'))
        fin_mois, fin_annee = map(int, date_fin.split('/'))
    except:
        # Si erreur de parsing, retourner des valeurs vides
        return {
            'temps_travail_total': 0,
            'depenses_externes': 0,
            'autres_achats': 0,
            'amortissements': 0
        }
    
    debut_date = datetime.date(debut_annee, debut_mois, 1)
    # Fin du mois de fin
    if fin_mois == 12:
        fin_date = datetime.date(fin_annee + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        fin_date = datetime.date(fin_annee, fin_mois + 1, 1) - datetime.timedelta(days=1)
    
    # Calculer le temps de travail total avec coût de catégorie
    cursor.execute("""
        SELECT tt.annee, tt.categorie, SUM(tt.jours)
        FROM temps_travail tt
        WHERE tt.projet_id = ?
        GROUP BY tt.annee, tt.categorie
    """, (projet_id,))
    
    temps_travail_total = 0
    for row in cursor.fetchall():
        annee, categorie, jours = row
        # Récupérer le coût de production pour cette catégorie/année
        cursor.execute("""
            SELECT cout_production FROM categorie_cout
            WHERE annee = ? AND categorie = ?
        """, (annee, categorie))
        cout_row = cursor.fetchone()
        cout_prod = cout_row[0] if cout_row and cout_row[0] else 0
        temps_travail_total += jours * cout_prod
    
    # Dépenses externes
    cursor.execute("""
        SELECT SUM(montant) FROM depenses WHERE projet_id = ?
    """, (projet_id,))
    depenses_externes = cursor.fetchone()[0] or 0
    
    # Autres achats
    cursor.execute("""
        SELECT SUM(montant) FROM autres_depenses WHERE projet_id = ?
    """, (projet_id,))
    autres_achats = cursor.fetchone()[0] or 0
    
    # Amortissements (calcul sur la période de la subvention)
    cursor.execute("""
        SELECT montant, date_achat, duree FROM investissements WHERE projet_id = ?
    """, (projet_id,))
    
    amortissements_total = 0
    for inv_row in cursor.fetchall():
        montant, date_achat, duree = inv_row
        if not date_achat or not duree or not montant:
            continue
        
        try:
            achat_mois, achat_annee = map(int, date_achat.split('/'))
            debut_amort = datetime.date(achat_annee, achat_mois, 1)
            
            # Fin amortissement = début + durée
            fin_mois_amort = achat_mois + int(duree) * 12
            fin_annee_amort = achat_annee + (fin_mois_amort - 1) // 12
            fin_mois_amort = ((fin_mois_amort - 1) % 12) + 1
            fin_amort = datetime.date(fin_annee_amort, fin_mois_amort, 1)
            
            # Intersection avec la période de subvention
            debut_effective = max(debut_amort, debut_date)
            fin_effective = min(fin_amort, fin_date)
            
            if debut_effective > fin_effective:
                continue
            
            # Calculer le nombre de mois d'amortissement
            mois_amort = (fin_effective.year - debut_effective.year) * 12 + fin_effective.month - debut_effective.month + 1
            dotation_mensuelle = float(montant) / (int(duree) * 12)
            amortissements_total += dotation_mensuelle * mois_amort
        except:
            continue
    
    return {
        'temps_travail_total': temps_travail_total,
        'depenses_externes': depenses_externes,
        'autres_achats': autres_achats,
        'amortissements': amortissements_total
    }


__all__ = ["DB_FILE", "DB_PATH", "db_cursor", "get_connection", "init_db", "recalculate_all_subventions"]
