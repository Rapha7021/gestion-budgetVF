import sqlite3
from contextlib import contextmanager
from pathlib import Path

# Nom du fichier de base de données stocké à la racine du projet
DB_FILENAME = "gestion_budget.db"
DB_FILE = Path(__file__).resolve().with_name(DB_FILENAME)
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


__all__ = ["DB_FILE", "DB_PATH", "db_cursor", "get_connection", "init_db"]
