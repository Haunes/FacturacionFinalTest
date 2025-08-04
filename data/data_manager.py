import pandas as pd
import streamlit as st
from typing import List, Optional
from .data_loader import DataLoader
from .data_filter import DataFilter

class DataManager:
    """Gestor centralizado de datos para la aplicación."""
    
    def __init__(self):
        self.data_loader = DataLoader()
        self.data_filter = DataFilter()
        self._initialize_session_state()
    
    def _initialize_session_state(self):
        """Inicializa el estado de sesión si no existe."""
        if 'df_combined' not in st.session_state:
            st.session_state.df_combined = pd.DataFrame()
    
    def load_files(self, uploaded_files: List) -> bool:
        """Carga archivos y actualiza el estado de sesión."""
        if not uploaded_files:
            return False
        
        df_combined = self.data_loader.load_excel_files(uploaded_files)
        st.session_state.df_combined = df_combined
        
        # Limpiar datos de descarga previos
        self._clear_download_data()
        
        return not df_combined.empty
    
    def get_data(self) -> pd.DataFrame:
        """Obtiene los datos combinados."""
        return st.session_state.df_combined
    
    def is_data_loaded(self) -> bool:
        """Verifica si hay datos cargados."""
        return not st.session_state.df_combined.empty
    
    def filter_data(self, empresa: str, anio: str, mes: str) -> pd.DataFrame:
        """Filtra los datos según los criterios especificados."""
        return self.data_filter.filter_data(
            st.session_state.df_combined, empresa, anio, mes
        )
    
    def get_filter_options(self, empresa: Optional[str] = None, anio: Optional[str] = None):
        """Obtiene las opciones disponibles para los filtros."""
        df = st.session_state.df_combined
        
        if df.empty:
            return {
                'empresas': ["Todas"],
                'anios': ["Todos"],
                'meses': ["Todos"]
            }
        
        # Empresas
        empresas = ["Todas"] + sorted(df['EMPRESA'].unique().tolist())
        
        # Filtrar por empresa si está seleccionada
        if empresa and empresa != "Todas":
            df = df[df['EMPRESA'] == empresa]
        
        # Años
        anios = ["Todos"] + sorted(df['AÑO ASIGNACION'].unique().tolist(), reverse=True)
        
        # Filtrar por año si está seleccionado
        if anio and anio != "Todos":
            df = df[df['AÑO ASIGNACION'] == anio]
        
        # Meses
        meses_ordenados = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        meses_disponibles = df['MES ASIGNACION'].unique().tolist()
        meses = ["Todos"] + [mes for mes in meses_ordenados if mes in meses_disponibles]
        
        return {
            'empresas': empresas,
            'anios': anios,
            'meses': meses
        }
    
    def _clear_download_data(self):
        """Limpia los datos de descarga del estado de sesión."""
        keys_to_remove = ["download_bytes", "download_name", "download_mime"]
        for key in keys_to_remove:
            st.session_state.pop(key, None)
    
    def set_download_data(self, bytes_data: bytes, filename: str, mime_type: str):
        """Establece los datos para descarga."""
        st.session_state.download_bytes = bytes_data
        st.session_state.download_name = filename
        st.session_state.download_mime = mime_type
    
    def has_download_data(self) -> bool:
        """Verifica si hay datos listos para descarga."""
        required_keys = ["download_bytes", "download_name", "download_mime"]
        return all(key in st.session_state for key in required_keys)
    
    def get_download_data(self) -> tuple:
        """Obtiene los datos de descarga."""
        if not self.has_download_data():
            return None, None, None
        
        return (
            st.session_state.download_bytes,
            st.session_state.download_name,
            st.session_state.download_mime
        )
