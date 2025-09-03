"""
Utilitaires pour l'application de gestion de budget
"""

def format_montant(montant, align_width=None):
    """
    Formate un montant en format français avec espaces pour les milliers 
    sans décimales.
    
    Args:
        montant: Le montant à formater
        align_width: Largeur pour l'alignement à droite (optionnel)
    
    Exemple: 
        81036.72 -> "81 037 €"
        format_montant(1234, 15) -> "      1 234 €"
    """
    if montant is None or montant == 0:
        result = "0 €"
    else:
        # Arrondir à l'entier le plus proche
        montant_arrondi = round(montant)
        
        # Formatage avec séparateur de milliers sans décimales
        formatted = f"{montant_arrondi:,}"
        
        # Conversion au format français
        # Remplace les virgules (séparateurs de milliers) par des espaces
        formatted = formatted.replace(",", " ")
        
        result = f"{formatted} €"
    
    # Alignement à droite si une largeur est spécifiée
    if align_width is not None:
        result = result.rjust(align_width)
    
    return result


def format_montant_aligne(montant):
    """
    Version alignée de format_montant avec une largeur fixe de 15 caractères.
    Pratique pour l'affichage en colonnes.
    """
    return format_montant(montant, align_width=15)
