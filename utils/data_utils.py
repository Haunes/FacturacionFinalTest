import pandas as pd

def find_column(df: pd.DataFrame, possible_names: list) -> str:
    """
    Busca una columna en el DataFrame usando una lista de posibles nombres.
    
    Args:
        df: DataFrame donde buscar
        possible_names: Lista de posibles nombres de columna
        
    Returns:
        Nombre de la primera columna encontrada o None
    """
    for col_name in possible_names:
        for actual_col in df.columns:
            if col_name.upper() in actual_col.upper():
                return actual_col
    return None

def get_document_count(df: pd.DataFrame) -> int:
    """
    Obtiene el número de documentos únicos.
    
    Args:
        df: DataFrame con los datos
        
    Returns:
        Número de documentos únicos
    """
    possible_names = ['NO. CASO', 'NUMERO CASO', 'CASO', 'ID', 'NUMERO', 'DOCUMENTO']
    col_name = find_column(df, possible_names)
    
    if col_name:
        return df[col_name].nunique()
    else:
        return len(df)

def get_representative_price(data: pd.DataFrame) -> float:
    """
    Obtiene un precio representativo para GWealth (moda de VALOR).
    
    Args:
        data: DataFrame con los datos
        
    Returns:
        Precio representativo
    """
    if 'VALOR' not in data.columns:
        return 0.0
    
    serie = pd.to_numeric(data['VALOR'], errors='coerce').dropna()
    if serie.empty:
        return 0.0
    
    moda = serie.mode()
    if not moda.empty:
        return float(moda.iloc[0])
    
    return float(serie.iloc[0])
