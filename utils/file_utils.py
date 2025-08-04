import unicodedata

def safe_filename(name: str) -> str:
    """
    Convierte un nombre a un formato seguro para archivos.
    Quita tildes y caracteres no ASCII; reemplaza espacios por guiones.
    """
    norm = unicodedata.normalize("NFKD", name)
    ascii_only = norm.encode("ascii", "ignore").decode("ascii")
    ascii_only = ascii_only.replace(" ", "-")
    keep = "".join(ch for ch in ascii_only if ch.isalnum() or ch in "-._")
    return keep or "archivo"

def ensure_extension(name: str, ext: str) -> str:
    """
    Asegura que el nombre termine con la extensión indicada.
    
    Args:
        name: Nombre del archivo
        ext: Extensión deseada (con o sin punto)
        
    Returns:
        Nombre con la extensión correcta
    """
    ext = ext if ext.startswith(".") else f".{ext}"
    return name if name.lower().endswith(ext.lower()) else f"{name}{ext}"
