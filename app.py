import streamlit as st
import pandas as pd
from datetime import datetime
from ui.sidebar import render_sidebar
from ui.main_content import render_main_content
from data.data_manager import DataManager
from utils.file_utils import safe_filename, ensure_extension
from report_generator import generate_report, build_report_filename
from excel_generator_ravago import create_ravago_report
from preview_generator_html import generate_preview_html
import unicodedata

# ---------------- Config de página ----------------
st.set_page_config(page_title="Generador de Reportes BIU", layout="wide")
st.title("📄 Generador de Reportes de Facturación")
st.markdown("Cargue sus archivos de Excel para comenzar a generar los reportes.")

def main():
    """Función principal de la aplicación."""
    # Inicializar el gestor de datos
    data_manager = DataManager()
    
    # Renderizar sidebar y obtener configuración
    config = render_sidebar(data_manager)
    
    # Renderizar contenido principal
    render_main_content(data_manager, config)

if __name__ == "__main__":
    main()
