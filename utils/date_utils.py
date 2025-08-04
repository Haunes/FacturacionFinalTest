from datetime import datetime

# Meses en español
MESES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

def fecha_es(dt: datetime) -> str:
    """
    Convierte una fecha a formato español 'dd de <mes> de yyyy'.
    
    Args:
        dt: Objeto datetime
        
    Returns:
        Fecha formateada en español
    """
    return f"{dt.day:02d} de {MESES_ES[dt.month-1]} de {dt.year}"
