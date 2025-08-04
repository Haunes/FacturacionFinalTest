import pandas as pd
import streamlit as st

def load_excel_files(uploaded_files):
    """
    Carga múltiples archivos Excel subidos a través de Streamlit,
    los combina en un único DataFrame y maneja posibles errores.
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
    
    date_columns = ['FECHA ASIGNACION', 'FECHA ENTREGA']
    for col in date_columns:
        if col in combined_df.columns:
            combined_df[col] = pd.to_datetime(combined_df[col], errors='coerce')

    return combined_df


def filter_data(df, empresa, anio, mes):
    """
    Filtra el DataFrame principal según la empresa, año y mes de asignación.
    """
    if df.empty:
        return pd.DataFrame()

    filtered = df.copy()

    if empresa and empresa != "Todas":
        filtered = filtered[filtered['EMPRESA'] == empresa]
    
    if anio and anio != "Todos":
        filtered = filtered[filtered['AÑO ASIGNACION'] == anio]

    if mes and mes != "Todos":
        filtered = filtered[filtered['MES ASIGNACION'] == mes]
        
    return filtered
