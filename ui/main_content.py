import streamlit as st
from typing import Dict, Any
from data.data_manager import DataManager
from reports.report_factory import ReportFactory
from preview.preview_generator import PreviewGenerator
from utils.file_utils import safe_filename, ensure_extension

def render_main_content(data_manager: DataManager, config: Dict[str, Any]):
    """
    Renderiza el contenido principal de la aplicación.
    
    Args:
        data_manager: Gestor de datos
        config: Configuración de filtros y reporte
    """
    if not data_manager.is_data_loaded():
        st.info("Esperando la carga de archivos Excel...")
        return
    
    # Filtrar datos
    df_filtered = data_manager.filter_data(
        config['empresa'], 
        config['anio'], 
        config['mes']
    )
    
    # Mostrar datos filtrados
    _render_filtered_data(df_filtered)
    
    # Mostrar previsualización
    _render_preview(df_filtered, config)
    
    # Mostrar controles de generación y descarga
    _render_report_controls(data_manager, df_filtered, config)

def _render_filtered_data(df_filtered):
    """Renderiza la tabla de datos filtrados."""
    st.header("Vista Previa de Datos Filtrados")
    
    if not df_filtered.empty:
        st.dataframe(df_filtered)
    else:
        st.warning("No se encontraron datos con los filtros seleccionados.")

def _render_preview(df_filtered, config):
    """Renderiza la previsualización del reporte."""
    st.header("Previsualización del Reporte")
    
    empresa = config['empresa']
    anio = config['anio']
    mes = config['mes']
    
    if empresa != "Todas" and anio != "Todos" and mes != "Todos" and not df_filtered.empty:
        with st.spinner("Generando previsualización..."):
            funcionarios = {
                'reporta': config['func_reporta'], 
                'revisor': config['func_revisor']
            }
            
            preview_generator = PreviewGenerator()
            preview_html = preview_generator.generate_preview_html(
                df_filtered, empresa, anio, mes, funcionarios
            )
            
            st.components.v1.html(preview_html, height=650, scrolling=True)
    else:
        st.warning("Por favor, seleccione una Empresa, Año y Mes específicos para generar un reporte.")

def _render_report_controls(data_manager: DataManager, df_filtered, config):
    """Renderiza los controles de generación y descarga de reportes."""
    empresa = config['empresa']
    anio = config['anio']
    mes = config['mes']
    
    # Solo mostrar controles si hay datos filtrados y selecciones válidas
    if df_filtered.empty or empresa == "Todas" or anio == "Todos" or mes == "Todos":
        return
    
    # Campo de nombre de archivo
    suggested_name = _get_suggested_filename(empresa)
    file_name_input = st.text_input(
        "Nombre del archivo (puedes modificarlo antes de descargar)",
        value=suggested_name,
        key=f"nombre_archivo_{empresa}_{anio}_{mes}"
    )
    
    # Botón de generación
    if st.button("✅ Generar Reporte"):
        _generate_report(data_manager, df_filtered, config, file_name_input, suggested_name)
    
    # Botón de descarga
    _render_download_button(data_manager)

def _get_suggested_filename(empresa: str) -> str:
    """Obtiene el nombre de archivo sugerido según la empresa."""
    report_factory = ReportFactory()
    
    if empresa == "Ravago Americas LLC":
        return report_factory.build_report_filename(empresa).replace(".docx", ".xlsx")
    else:
        return report_factory.build_report_filename(empresa)

def _generate_report(data_manager: DataManager, df_filtered, config, file_name_input: str, suggested_name: str):
    """Genera el reporte según la configuración."""
    empresa = config['empresa']
    anio = config['anio']
    mes = config['mes']
    funcionarios = {
        'reporta': config['func_reporta'], 
        'revisor': config['func_revisor']
    }
    
    with st.spinner("Creando documento..."):
        try:
            report_factory = ReportFactory()
            buffer, mime_type = report_factory.create_report(
                df_filtered, empresa, anio, mes, funcionarios
            )
            
            # Determinar nombre final del archivo
            ext = ".xlsx" if empresa == "Ravago Americas LLC" else ".docx"
            raw_name = (file_name_input or suggested_name).strip()
            final_name = ensure_extension(safe_filename(raw_name), ext)
            
            # Guardar datos para descarga
            data_manager.set_download_data(buffer.getvalue(), final_name, mime_type)
            
            st.success(f"¡Reporte generado! Nombre: **{final_name}**")
            
        except Exception as e:
            st.error(f"Ocurrió un error al generar el reporte: {e}")

def _render_download_button(data_manager: DataManager):
    """Renderiza el botón de descarga si hay datos disponibles."""
    if not data_manager.has_download_data():
        return
    
    bytes_data, filename, mime_type = data_manager.get_download_data()
    
    st.download_button(
        label=f"📥 Descargar {filename}",
        data=bytes_data,
        file_name=filename,
        mime=mime_type,
        key=f"download_{filename}"
    )
