import pandas as pd

# Remover esta línea:
# from utils import format_currency, get_document_count

# Agregar estas funciones directamente:
def format_currency(value, currency="USD"):
    """Formatea un número como moneda."""
    try:
        return f"{currency} {float(value):,.2f}"
    except (ValueError, TypeError):
        return f"{currency} 0.00"

def find_column(df, possible_names):
    """Busca una columna en el DataFrame usando una lista de posibles nombres."""
    for col_name in possible_names:
        for actual_col in df.columns:
            if col_name.upper() in actual_col.upper():
                return actual_col
    return None

def get_document_count(df):
    """Obtiene el número de documentos únicos."""
    possible_names = ['NO. CASO', 'NUMERO CASO', 'CASO', 'ID', 'NUMERO', 'DOCUMENTO']
    col_name = find_column(df, possible_names)
    
    if col_name:
        return df[col_name].nunique()
    else:
        return len(df)

# ------------------------
# Utilidades
# ------------------------
def get_representative_price(data: pd.DataFrame) -> float:
    """Precio único para GWealth (moda de VALOR; si no, primer no nulo)."""
    if 'VALOR' not in data.columns:
        return 0.0
    serie = pd.to_numeric(data['VALOR'], errors='coerce').dropna()
    if serie.empty:
        return 0.0
    moda = serie.mode()
    if not moda.empty:
        return float(moda.iloc[0])
    return float(serie.iloc[0])

# ------------------------
# HTML
# ------------------------
def generate_preview_html(data, empresa, anio, mes, funcionarios):
    """Genera una previsualización HTML del reporte."""

    css_styles = """
    <style>
        .preview-container { font-family: 'Calibri', sans-serif; font-size: 11pt; padding: 20px; background-color: #f7f7f7; border-radius: 8px; }
        .header-info { margin-bottom: 20px; }
        .logo { float: right; width: 135px; height: 60px; background-color: #e0e0e0; text-align: center; line-height: 60px; font-weight: bold; }

        /* --- Estilo por defecto (Ravago, parecido a Excel) --- */
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid black; padding: 8px; text-align: left; }
        th { background-color: #002060; color: white; text-align: center; }       /* Ravago */
        .total-row td { background-color: #808080; color: white; font-weight: bold; } /* Ravago */

        /* --- Estilos que imitan Word, SOLO dentro de contenedores .word --- */
        .word th { background-color: #003366; color: #FFFFFF; text-align: center; }   /* PRIMARY Word */
        .word .body-row td { background-color: #F0F0F0; color: #000000; }             /* cuerpo Word */
        .word .total-row td { background-color: #E20074; color: #FFFFFF; font-weight: bold; } /* ACCENT Word */

        .right-align { text-align: right; }
        .left-align { text-align: left; }
        .center-align { text-align: center; }
        .footer-note { font-size: 9pt; font-style: italic; margin-top: 15px; }
    </style>
    """

    html_body = ""
    if empresa == "Ravago Americas LLC":
        # --- Vista tipo Excel (se mantiene como estaba) ---
        num_docs = get_document_count(data)
        total_valor = data['VALOR'].sum()

        html_body = f"""
        <h4>Hoja: Resumen</h4>
        <table>
            <tr><th>Año</th><th>Mes</th><th>Documentos Revisados (Ver Anexo 1)</th></tr>
            <tr><td class='center-align'>{anio}</td><td class='center-align'>{mes}</td><td class='center-align'>{num_docs}</td></tr>
            <tr class='total-row'><td colspan='2' class='center-align'>Total Por Facturar</td><td class='center-align'>{num_docs}</td></tr>
        </table>
        <table>
            <tr><th>Concepto</th><th>Total (antes de I.V.A)</th></tr>
            <tr><td>Revisión de {num_docs} documentos durante el mes de {mes} de {anio}</td><td class='right-align'>{format_currency(total_valor)}</td></tr>
            <tr class='total-row'><td class='right-align'>SUBTOTAL</td><td class='right-align'>{format_currency(total_valor)}</td></tr>
        </table>
        <div class='footer-note'>TRM Aplicable: Según la propuesta, es aquella de emisión de la factura.</div>
        <div class='footer-note'>biu usually issues monthly invoices...</div>
        <hr>
        <h4>Hoja: Honorarios</h4>
        <table>
            <tr><th>FECHA</th><th>NOMBRE CONTRAPARTE</th><th>TIPO DE DOCUMENTO</th><th>TOTAL</th></tr>
        """
        for idx, row in data.iterrows():
            html_body += f"<tr><td class='center-align'>{idx+1}</td><td>{row.get('NOMBRE', '')}</td><td>{row.get('TIPO DE DOCUMENTO', '')}</td><td class='right-align'>{format_currency(row.get('VALOR', 0))}</td></tr>"
        html_body += f"<tr class='total-row'><td colspan='3' class='right-align'>SUBTOTAL</td><td class='right-align'>{format_currency(total_valor)}</td></tr></table>"

    else:
        # --- Vista estilo Word (Altimetrik y GWealth) ---
        main_table_html = generate_main_table_html(data, empresa)
        summary_tables_html = generate_summary_tables_html(data, empresa, anio, mes)

        html_body = f"""
        <div class="header-info">
            <div class="logo">BIU<br>Logo</div>
        </div>

        <h4>FACTURACIÓN {mes.upper()} {anio}</h4>
        <h4>{empresa.upper()}</h4>

        <div class="info-section">
            <div class="info-line">Fecha de corte del reporte: </div>
            <div class="info-line">Funcionario que reporta: &nbsp;&nbsp; {funcionarios['reporta']}</div>
            <div class="info-line">Funcionario revisor: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {funcionarios['revisor']}</div>
        </div>

        <div class="word">
            {main_table_html}
            {summary_tables_html}
        </div>

        <div class="footer-note">Nota: Este es un documento de previsualización.</div>
        <hr>
        <div class="footer">
            Número: 601 - 7455289 | Dirección: Carrera 7 No. 74B-56, Oficina 301 | Correo: info@biu.com.co
        </div>
        """

    return f"<div class='preview-container'>{css_styles}{html_body}</div>"

def generate_main_table_html(data, empresa: str):
    """Genera la tabla principal de datos (la fila Total respeta la regla de GWealth)."""
    total_valor_sum = data['VALOR'].sum() if 'VALOR' in data.columns else 0.0
    total_gw = get_representative_price(data) if empresa == "Gwealth" else total_valor_sum
    etiqueta = "Total (precio único)" if empresa == "Gwealth" else "Total"

    html = """
    <table>
        <thead>
            <tr>
                <th>MES ASIGNACION</th>
                <th>AÑO ASIGNACION</th>
                <th>NOMBRE</th>
                <th>MONEDA</th>
                <th>VALOR</th>
            </tr>
        </thead>
        <tbody>
    """
    # Filas de datos con sombreado estilo Word
    for _, row in data.iterrows():
        html += f"""
            <tr class="body-row">
                <td>{row.get('MES ASIGNACION', '')}</td>
                <td>{row.get('AÑO ASIGNACION', '')}</td>
                <td>{row.get('NOMBRE', '')}</td>
                <td>{row.get('MONEDA', '')}</td>
                <td>{row.get('VALOR', 0):,.2f}</td>
            </tr>
        """

    # Fila de total: combinamos las 4 primeras columnas y dejamos VALOR para el importe
    html += f"""
            <tr class="total-row">
                <td colspan="4" class="center-align">{etiqueta}</td>
                <td>{total_gw:,.2f}</td>
            </tr>
        </tbody>
    </table>
    """
    return html

def generate_summary_tables_html(data, empresa, anio, mes):
    """Genera las tablas de resumen específicas por empresa."""
    total_valor_sum = data['VALOR'].sum() if 'VALOR' in data.columns else 0.0

    if empresa == "Altimetrik":
        return f"""
        <table class="summary-table">
            <thead>
                <tr>
                    <th>Mes</th>
                    <th>Concepto</th>
                    <th>Total</th>
                </tr>
            </thead>
            <tbody>
                <tr class="body-row">
                    <td>{mes}</td>
                    <td>Consultas en listas recibidas en {mes} de {anio}</td>
                    <td>USD {total_valor_sum:,.2f}</td>
                </tr>
            </tbody>
        </table>
        """

    elif empresa == "Gwealth":
        precio_unico = get_representative_price(data)
        iva = precio_unico * 0.19
        total_con_iva = precio_unico + iva
        return f"""
        <table class="summary-table">
            <thead>
                <tr>
                    <th>Mes</th>
                    <th>Concepto</th>
                    <th>Total</th>
                </tr>
            </thead>
            <tbody>
                <tr class="body-row">
                    <td>{mes}</td>
                    <td>Consultas en listas recibidas en {mes} de {anio}</td>
                    <td>USD {precio_unico:,.2f}</td>
                </tr>
                <tr class="total-row">
                    <td colspan="2" class="center-align">TOTAL</td>
                    <td>USD {precio_unico:,.2f}</td>
                </tr>
                <tr class="total-row">
                    <td colspan="2" class="center-align">TOTAL CON IVA</td>
                    <td>USD {total_con_iva:,.2f}</td>
                </tr>
            </tbody>
        </table>
        """
    else:
        return ""
