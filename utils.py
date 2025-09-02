"""
Utilitaires pour l'application de gestion de budget
"""

def format_montant(montant):
    """
    Formate un montant en format français avec espaces pour les milliers 
    et virgule pour les décimales.
    
    Exemple: 81036.72 -> "81 036,72 €"
    """
    if montant is None or montant == 0:
        return "0,00 €"
    
    # Formatage avec séparateur de milliers et 2 décimales
    formatted = f"{montant:,.2f}"
    
    # Conversion au format français
    # Remplace les virgules (séparateurs de milliers) par un caractère temporaire
    # Puis remplace les points (séparateurs décimaux) par des virgules
    # Enfin remplace le caractère temporaire par des espaces
    formatted = formatted.replace(",", "TEMP").replace(".", ",").replace("TEMP", " ")
    
    return f"{formatted} €"
