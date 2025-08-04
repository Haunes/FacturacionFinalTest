import pandas as pd
import streamlit as st
from typing import List

class DataLoader:
    """Maneja la carga de archivos Excel."""
    
    def load_excel_files(self, uploaded_files: List) -> pd.DataFrame:
        """
        Carga múltiples archivos Excel y los combina en un DataFrame.
        
        Args:
            uploaded_files: Lista de archivos subidos por Streamlit
            
        Returns:
            DataFrame combinado con todos los datos
        """
        if not uploaded_files:
            return pd.DataFrame()

        all_data = []
        
        for file in uploaded_files:
            try:
                df = pd.read_excel(file, engine='openpyxl')
                all_data.append(df)
            except Exception as e:
                st.error(f"Error al leer el archivo {file.name}: {e}")
                continue

        if not all_data:
            return pd.DataFrame()

        combined_df = pd.concat(all_data, ignore_index=True)
        
        # Convertir columnas de fecha
        self._convert_date_columns(combined_df)
        
        return combined_df
    
    def _convert_date_columns(self, df: pd.DataFrame):
        """Convierte las columnas de fecha al tipo datetime."""
        date_columns = ['FECHA ASIGNACION', 'FECHA ENTREGA']
        
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
