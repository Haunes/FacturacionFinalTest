import streamlit as st
import pandas as pd
from datetime import datetime

# Intentar usar los nuevos módulos, si no usar los originales
try:
    from ui.sidebar import render_sidebar
    from ui.main_content import render_main_content
    from data.data_manager import DataManager
    from utils.file_utils import safe_filename, ensure_extension
except ImportError:
    # Fallback a funcionalidad original si los nuevos módulos no están disponibles
    from data_handler import load_excel_files, filter_data
    from utils import format_currency, get_document_count, find_column

# Imports que siempre deben funcionar
from report_generator import generate_report, build_report_filename
from excel_generator_ravago import create_ravago_report
from preview_generator_html import generate_preview_html
import unicodedata

# ---------------- Helpers para nombres ----------------
def safe_filename(name: str) -> str:
    """Quita tildes y caracteres no ASCII; reemplaza espacios por guiones; deja solo [A-Za-z0-9-._]."""
    norm = unicodedata.normalize("NFKD", name)
    ascii_only = norm.encode("ascii", "ignore").decode("ascii")
    ascii_only = ascii_only.replace(" ", "-")
    keep = "".join(ch for ch in ascii_only if ch.isalnum() or ch in "-._")
    return keep or "archivo"

def ensure_extension(name: str, ext: str) -> str:
    """Asegura que el nombre termine con la extensión indicada (con punto)."""
    ext = ext if ext.startswith(".") else f".{ext}"
    return name if name.lower().endswith(ext.lower()) else f"{name}{ext}"

# ---------------- Config de página ----------------
st.set_page_config(page_title="Generador de Reportes BIU", layout="wide")
st.title("📄 Generador de Reportes de Facturación")
st.markdown("Cargue sus archivos de Excel para comenzar a generar los reportes.")

def main():
    """Función principal de la aplicación."""
    try:
        # Usar nueva arquitectura si está disponible
        data_manager = DataManager()
        config = render_sidebar(data_manager)
        render_main_content(data_manager, config)
    except NameError:
        # Fallback a la implementación original
        main_original()

def main_original():
    """Implementación original como fallback."""
    # ---------------- Carga de archivos ----------------
    with st.sidebar:
        st.header("1. Cargar Archivos")
        uploaded_files = st.file_uploader(
            "Seleccione uno o más archivos Excel", type=["xlsx", "xls"], accept_multiple_files=True
        )

    if 'df_combined' not in st.session_state:
        st.session_state.df_combined = pd.DataFrame()

    # Limpia datos de descarga si el usuario carga nuevos archivos
    if uploaded_files:
        st.session_state.df_combined = load_excel_files(uploaded_files)
        st.session_state.pop("download_bytes", None)
        st.session_state.pop("download_name", None)
        st.session_state.pop("download_mime", None)

    # ---------------- Interfaz principal ----------------
    if not st.session_state.df_combined.empty:
        df = st.session_state.df_combined
        
        with st.sidebar:
            st.header("2. Aplicar Filtros")
            empresas_options = ["Todas"] + sorted(df['EMPRESA'].unique().tolist())
            empresa_sel = st.selectbox("Empresa", options=empresas_options)

            is_empresa_selected = empresa_sel != "Todas"
            df_empresa = df[df['EMPRESA'] == empresa_sel] if is_empresa_selected else df
            
            anio_options = ["Todos"] + sorted(df_empresa['AÑO ASIGNACION'].unique().tolist(), reverse=True)
            anio_sel = st.selectbox("Año de Asignación", options=anio_options, disabled=not is_empresa_selected)

            is_anio_selected = anio_sel != "Todos"
            df_anio = df_empresa[df_empresa['AÑO ASIGNACION'] == anio_sel] if is_anio_selected else df_empresa
            
            meses_ordenados = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            meses_disponibles = df_anio['MES ASIGNACION'].unique().tolist()
            mes_options = ["Todos"] + [mes for mes in meses_ordenados if mes in meses_disponibles]
            mes_sel = st.selectbox("Mes de Asignación", options=mes_options, disabled=not is_anio_selected)
            
            st.header("3. Información del Reporte")
            if empresa_sel == "Ravago Americas LLC":
                st.info("Para Ravago, los campos de funcionarios se llenarán manualmente en el Excel generado.")
                func_reporta = ""
                func_revisor = ""
            else:
                func_reporta = st.text_input("Funcionario que reporta", "")
                func_revisor = st.text_input("Funcionario revisor", "")
            
        df_filtered = filter_data(df, empresa_sel, anio_sel, mes_sel)

        st.header("Vista Previa de Datos Filtrados")
        if not df_filtered.empty:
            st.dataframe(df_filtered)

            st.header("Previsualización del Reporte")
            if empresa_sel != "Todas" and anio_sel != "Todos" and mes_sel != "Todos":
                with st.spinner("Generando previsualización..."):
                    funcionarios = {'reporta': func_reporta, 'revisor': func_revisor}
                    preview_html = generate_preview_html(df_filtered, empresa_sel, anio_sel, mes_sel, funcionarios)
                    st.components.v1.html(preview_html, height=650, scrolling=True)
            else:
                st.warning("Por favor, seleccione una Empresa, Año y Mes específicos para generar un reporte.")

            # ---------- Nombre de archivo editable ----------
            if empresa_sel != "Todas" and anio_sel != "Todos" and mes_sel != "Todos":
                if empresa_sel == "Ravago Americas LLC":
                    suggested_name = build_report_filename(empresa_sel).replace(".docx", ".xlsx")
                    ext = ".xlsx"
                else:
                    suggested_name = build_report_filename(empresa_sel)
                    ext = ".docx"

                file_name_input = st.text_input(
                    "Nombre del archivo (puedes modificarlo antes de descargar)",
                    value=suggested_name,
                    key=f"nombre_archivo_{empresa_sel}_{anio_sel}_{mes_sel}"
                )

            # -------------- Generar archivo --------------
            if st.button("✅ Generar Reporte"):
                if empresa_sel != "Todas" and anio_sel != "Todos" and mes_sel != "Todos":
                    with st.spinner("Creando documento..."):
                        funcionarios = {'reporta': func_reporta, 'revisor': func_revisor}
                        try:
                            if empresa_sel == "Ravago Americas LLC":
                                # Excel
                                buffer = create_ravago_report(df_filtered, anio_sel, mes_sel, funcionarios)
                                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            else:
                                # Word
                                buffer = generate_report(df_filtered, empresa_sel, anio_sel, mes_sel, funcionarios)
                                mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                            # Toma lo que el usuario escribió; si está vacío usa sugerido
                            raw_name = (file_name_input or suggested_name).strip()
                            final_name = ensure_extension(safe_filename(raw_name), ext)

                            # Persistencia para el download_button
                            st.session_state.download_bytes = buffer.getvalue()
                            st.session_state.download_name = final_name
                            st.session_state.download_mime = mime

                            st.success(f"¡Reporte generado! Nombre: **{final_name}**")
                        except Exception as e:
                            st.error(f"Ocurrió un error al generar el reporte: {e}")
                else:
                    st.error("Debe seleccionar una Empresa, Año y Mes para generar el reporte.")

            # -------------- Botón de descarga --------------
            if all(k in st.session_state for k in ("download_bytes", "download_name", "download_mime")):
                st.download_button(
                    label=f"📥 Descargar {st.session_state.download_name}",
                    data=st.session_state.download_bytes,
                    file_name=st.session_state.download_name,
                    mime=st.session_state.download_mime,
                    key=f"download_{st.session_state.download_name}"
                )

        else:
            st.warning("No se encontraron datos con los filtros seleccionados.")
    else:
        st.info("Esperando la carga de archivos Excel...")

if __name__ == "__main__":
    main()
