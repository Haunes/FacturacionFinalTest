def format_currency(value, currency="USD"):
    """Formatea un número como moneda."""
    try:
        return f"{currency} {float(value):,.2f}"
    except (ValueError, TypeError):
        return f"{currency} 0.00"

def find_column(df, possible_names):
    """
    Busca una columna en el DataFrame usando una lista de posibles nombres.
    Retorna el nombre de la primera columna encontrada o None.
    """
    for col_name in possible_names:
        for actual_col in df.columns:
            if col_name.upper() in actual_col.upper():
                return actual_col
    return None

def get_document_count(df):
    """
    Obtiene el número de documentos únicos usando diferentes posibles nombres de columna.
    """
    possible_names = ['NO. CASO', 'NUMERO CASO', 'CASO', 'ID', 'NUMERO', 'DOCUMENTO']
    col_name = find_column(df, possible_names)
    
    if col_name:
        return df[col_name].nunique()
    else:
        # Si no encuentra ninguna columna específica, usa el número de filas
        return len(df)
