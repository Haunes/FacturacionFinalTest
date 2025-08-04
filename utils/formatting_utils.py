def format_currency(value, currency="USD"):
    """
    Formatea un número como moneda.
    
    Args:
        value: Valor numérico
        currency: Código de moneda
        
    Returns:
        Cadena formateada como moneda
    """
    try:
        return f"{currency} {float(value):,.2f}"
    except (ValueError, TypeError):
        return f"{currency} 0.00"
