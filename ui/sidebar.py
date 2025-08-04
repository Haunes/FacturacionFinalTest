import streamlit as st
from typing import Dict, Any
from data.data_manager import DataManager

def render_sidebar(data_manager: DataManager) -> Dict[str, Any]:
    """
    Renderiza la barra lateral y retorna la configuración seleccionada.
    
    Args:
        data_manager: Gestor de datos
        
    Returns:
        Diccionario con la configuración seleccionada
    """
    with st.sidebar:
        # Sección 1: Cargar archivos
        st.header("1. Cargar Archivos")
        uploaded_files = st.file_uploader(
            "Seleccione uno o más archivos Excel", 
            type=["xlsx", "xls"], 
            accept_multiple_files=True
        )
        
        # Cargar archivos si se proporcionaron
        if uploaded_files:
            data_manager.load_files(uploaded_files)
        
        # Sección 2: Filtros
        config = _render_filters(data_manager)
        
        # Sección 3: Información del reporte
        config.update(_render_report_info(config.get('empresa')))
        
        return config

def _render_filters(data_manager: DataManager) -> Dict[str, Any]:
    """Renderiza los controles de filtros."""
    st.header("2. Aplicar Filtros")
    
    if not data_manager.is_data_loaded():
        return {
            'empresa': "Todas",
            'anio': "Todos",
            'mes': "Todos"
        }
    
    # Obtener opciones iniciales
    options = data_manager.get_filter_options()
    
    # Selector de empresa
    empresa_sel = st.selectbox("Empresa", options=options['empresas'])
    
    # Actualizar opciones basadas en empresa seleccionada
    is_empresa_selected = empresa_sel != "Todas"
    options = data_manager.get_filter_options(empresa=empresa_sel if is_empresa_selected else None)
    
    # Selector de año
    anio_sel = st.selectbox(
        "Año de Asignación", 
        options=options['anios'], 
        disabled=not is_empresa_selected
    )
    
    # Actualizar opciones basadas en año seleccionado
    is_anio_selected = anio_sel != "Todos"
    if is_anio_selected:
        options = data_manager.get_filter_options(empresa=empresa_sel, anio=anio_sel)
    
    # Selector de mes
    mes_sel = st.selectbox(
        "Mes de Asignación", 
        options=options['meses'], 
        disabled=not is_anio_selected
    )
    
    return {
        'empresa': empresa_sel,
        'anio': anio_sel,
        'mes': mes_sel
    }

def _render_report_info(empresa: str) -> Dict[str, Any]:
    """Renderiza la sección de información del reporte."""
    st.header("3. Información del Reporte")
    
    if empresa == "Ravago Americas LLC":
        st.info("Para Ravago, los campos de funcionarios se llenarán manualmente en el Excel generado.")
        return {
            'func_reporta': "",
            'func_revisor': ""
        }
    else:
        func_reporta = st.text_input("Funcionario que reporta", "")
        func_revisor = st.text_input("Funcionario revisor", "")
        
        return {
            'func_reporta': func_reporta,
            'func_revisor': func_revisor
        }
