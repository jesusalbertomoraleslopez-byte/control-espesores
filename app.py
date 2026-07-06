import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import scipy.stats as stats
from datetime import datetime
import io
import importlib
import database
importlib.reload(database)
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 1. CONFIGURACIÓN ÚNICA DE LA PÁGINA (Debe ser la primera instrucción de Streamlit)
st.set_page_config(page_title="Suite de Riesgo y Control de Espesores", layout="wide")

# 2. CONFIGURACIÓN GRÁFICA VECTORIAL Y REPORTES NATIVOS
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

import matplotlib
matplotlib.use('Agg') # Renderizado seguro en backend para servidores sin entorno gráfico
import matplotlib.pyplot as plt

# 3. MATRIZ TÉCNICA PLANTA (Límites rígidos de diseño nominal)
ESTANDAR = {
    "Galvanizado": {10: 0.138, 12: 0.108, 14: 0.079, 16: 0.064},
    "Decapado": {10: 0.135, 12: 0.105, 14: 0.075, 16: 0.060}
}
TOLERANCIA_INTERNA = 0.008

# Despliegue de banner corporativo principal
st.image(os.path.join(BASE_DIR, "BANNER CONTROL DE ESPESORES APP.png"), use_container_width=True)
def colorear_matriz_resumen(v):
    """Aplica formato semafórico condicional estricto: Verde para aceptados y Rojo para riesgosos/rechazados."""
    if not isinstance(v, str): return ''
    if "BAJO" in v or "ACEPTADO" == v: 
        return 'background-color: #C6EFCE; color: #006100; font-weight: bold;'
    if "MODERADO" in v: 
        return 'background-color: #FFF2CC; color: #7F6000; font-weight: bold;'
    return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold;'

def generar_excel_plantilla():
    """Construye el archivo binario de Excel en la memoria RAM con el formato corporativo oficial."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    datos = {
        "Numero_Rollo": ["ROLLO-A", "ROLLO-B", "ROLLO-C"],
        "Material": ["Galvanizado", "Galvanizado", "Decapado"],
        "Calibre": [12, 16, 14],
        "Espesor_Medido": [0.1028, 1.47, 1.85],
        "Unidad": ["Pulgadas", "Milimetros", "Milimetros"],
        "Tolerancia_Proveedor": [0.006, 0.005, 0.006]
    }
    
    df_tpl = pd.DataFrame(datos)
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_tpl.to_excel(writer, index=False, sheet_name="Datos_Simulacion")
        workbook = writer.book
        worksheet = workbook["Datos_Simulacion"]
        
        # Estilos oficiales Sigrama
        header_fill = PatternFill(start_color="111111", end_color="111111", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        data_font = Font(name="Calibri", size=11, color="111111")
        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        
        thin_side = Side(border_style="thin", color="D2D3D5")
        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        
        # Formatear cabeceras
        for col_num, col_name in enumerate(df_tpl.columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = cell_border
            
        # Formatear celdas de datos
        for row_num in range(2, len(df_tpl) + 2):
            for col_num in range(1, len(df_tpl.columns) + 1):
                cell = worksheet.cell(row=row_num, column=col_num)
                cell.font = data_font
                cell.border = cell_border
                # Alineación
                if col_num in [1, 2]: # Rollo, Material
                    cell.alignment = left_align
                else: # Calibre, Espesor, Unidad, Tolerancia
                    cell.alignment = center_align
                    
        # Autoajustar ancho de columnas para evitar cortes de texto (###)
        for col in worksheet.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            # Margen de seguridad para el ancho
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 16)
            
    return output.getvalue()
def crear_pdf_formal(df_final, tol_p, cert_file_data=None, email_img_data=None, df_raw=None, meta_info=None):
    """Genera la estructura del documento técnico formal incorporando los nuevos formatos de color y títulos, y miniaturas."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    styles = getSampleStyleSheet()
    
    t_st = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor('#EC2024'), spaceAfter=4, alignment=2)
    m_st = ParagraphStyle('DocMeta', fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#111111'), spaceAfter=2)
    h2_st = ParagraphStyle('SectionH2', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor('#111111'), spaceBefore=10, spaceAfter=4, backColor=colors.HexColor('#F1F5F9'), borderPadding=6)
    h3_st = ParagraphStyle('SectionH3', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#EC2024'), spaceBefore=6, spaceAfter=4, alignment=1)
    h_style = ParagraphStyle('HStyle', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1)
    c_style = ParagraphStyle('CStyle', fontName='Helvetica', fontSize=9, alignment=1)
    
    logo_path = os.path.join(database.BASE_DIR, "logo_sigrama.png")
    if os.path.exists(logo_path):
        logo = RLImage(logo_path, width=160, height=45, kind='proportional')
        header_data = [[logo, Paragraph("REPORTE TÉCNICO DE INGENIERÍA<br/><font size=10 color='#111111'>EVALUACIÓN DE SUMINISTRO</font>", t_st)]]
        t_head = Table(header_data, colWidths=[180, 340])
        t_head.setStyle(TableStyle([('ALIGN', (0,0), (0,0), 'LEFT'), ('ALIGN', (1,0), (1,0), 'RIGHT'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LINEBELOW', (0,0), (-1,-1), 1.5, colors.HexColor('#EC2024'))]))
        story.append(t_head)
        story.append(Spacer(1, 10))
    else:
        story.append(Paragraph("SIGRAMA PLANTA METALES", t_st))
        
    if meta_info:
        meta_data = [
            [Paragraph("<b>Folio Oficial:</b>", m_st), Paragraph(meta_info.get("Folio", "Borrador"), m_st), Paragraph("<b>Fecha de Análisis:</b>", m_st), Paragraph(meta_info.get("Fecha", datetime.now().strftime('%d/%m/%Y %H:%M')), m_st)],
            [Paragraph("<b>Proveedor:</b>", m_st), Paragraph(meta_info.get("Proveedor", "N/D"), m_st), Paragraph("<b>Contacto (Email/Tel):</b>", m_st), Paragraph(meta_info.get("Contacto", "N/D"), m_st)],
            [Paragraph("<b>Certificado (Lote/ID):</b>", m_st), Paragraph(meta_info.get("Certificado", "N/D"), m_st), Paragraph("<b>Tolerancia Base:</b>", m_st), Paragraph(f"±{tol_p:.3f}\"", m_st)]
        ]
        t_meta = Table(meta_data, colWidths=[110, 180, 110, 120])
        t_meta.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#F8F9FA')),
            ('BACKGROUND', (2,0), (2,-1), colors.HexColor('#F8F9FA')),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3')),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6)
        ]))
        story.append(t_meta)
    else:
        story.append(Paragraph("<b>Documento:</b> Reporte Técnico de Ingeniería de Calidad y Evaluación de Suministro", m_st))
        story.append(Paragraph(f"<b>Fecha de Análisis:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", m_st))
        story.append(Paragraph(f"<b>Parámetro Comercial:</b> Desviación Ofertada por Proveedor en ±{tol_p:.3f}\"", m_st))
        
    story.append(Spacer(1, 15))
    
    # 1. Tabla de calibración general
    story.append(Paragraph("1. Calibración del Muestreo por Unidad (Rollo por Rollo)", h2_st))
    t_rollos_d = [[Paragraph("Rollo", h_style), Paragraph("Material", h_style), Paragraph("Calibre", h_style), Paragraph("Espesor (in)", h_style), Paragraph("Riesgo %", h_style), Paragraph("Riesgo", h_style)]]
    for idx, f in df_final.reset_index(drop=True).iterrows():
        t_rollos_d.append([
            Paragraph(str(f['Rollo']), c_style), Paragraph(f['Material'], c_style), Paragraph(str(f['Calibre']), c_style),
            Paragraph(f"{f['Espesor Real (in)']:.4f}\"", c_style), Paragraph(f"{f['% de Riesgo']:.2f}%", c_style), Paragraph(f['Riesgo'], c_style)
        ])
    t_1 = Table(t_rollos_d, colWidths=[90, 100, 70, 90, 80, 100])
    t_1.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#111111')), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3'))]))
    story.append(t_1)
    
    # 2. Análisis jerárquico estructurado
    story.append(Paragraph("2. Análisis Estructurado y Clasificación por Espesor Nominal", h2_st))
    df_grouped = df_final.groupby(['Material', 'Calibre', 'Nominal Estándar'])
    
    for (mat, calibre, nominal), group in df_grouped:
        # Formato de especificación exacta solicitada con Tolerancia Aceptable
        texto_especificacion = f"Especificación: {mat} - {calibre} - Espesor Teórico: {nominal:.3f}\" | Tolerancia Aceptable: ±{TOLERANCIA_INTERNA:.3f}\""
        story.append(Paragraph(texto_especificacion, h3_st))
        
        # Estructura limpia sin la columna 'Espesor Original' en los desgloses
        t_group_d = [[Paragraph("Número Rollo", h_style), Paragraph("Espesor Medido (in)", h_style), Paragraph("Desviación Real", h_style), Paragraph("Probabilidad de Fallo", h_style), Paragraph("Dictamen Final", h_style)]]
        est_estilo_grupo = [('BACKGROUND', (0,0), (-1,0), colors.HexColor('#111111')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#A0A0A0'))]
        
        for idx_fila, (_, fila) in enumerate(group.iterrows()):
            fila_pdf = idx_fila + 1
            if fila['Dictamen Final'] == "ACEPTADO":
                bg_color = colors.HexColor('#C6EFCE')
                text_color_hex = '#006100'
            else:
                bg_color = colors.HexColor('#FFC7CE')
                text_color_hex = '#9C0006'
                
            p_dictamen = Paragraph(f"<font color='{text_color_hex}'><b>{fila['Dictamen Final']}</b></font>", c_style)
            t_group_d.append([
                Paragraph(str(fila['Rollo']), c_style), Paragraph(f"{fila['Espesor Real (in)']:.4f}\"", c_style),
                Paragraph(f"{fila['Desviación Real (in)']:+4f}\"", c_style), Paragraph(f"{fila['% de Riesgo']:.2f}%", c_style), p_dictamen
            ])
            est_estilo_grupo.append(('BACKGROUND', (4, fila_pdf), (4, fila_pdf), bg_color))
            
        t_g = Table(t_group_d, colWidths=[110, 110, 110, 110, 110])
        t_g.setStyle(TableStyle(est_estilo_grupo))
        story.append(t_g)
        story.append(Spacer(1, 8))
        
    # 3. Distribución estadística de Gauss por sección aislada
    story.append(Spacer(1, 10))
    story.append(Paragraph("3. Análisis de Distribución Probabilística por Especificación Técnica", h2_st))
    
    sigma_p = tol_p / 3.0
    for (mat, calibre, nominal), group in df_grouped:
        try:
            plt.figure(figsize=(6.5, 2.8))
            x_desv = np.linspace(-0.015, 0.015, 400)
            
            for _, fila in group.iterrows():
                m_desv = fila['Espesor Real (in)'] - nominal
                tol_r = fila.get('Tolerancia_Rollo', tol_p)
                sigma_rollo = tol_r / 3.0
                y_g = stats.norm.pdf(x_desv, loc=m_desv, scale=sigma_rollo)
                plt.plot(x_desv, y_g, label=f"{fila['Rollo']} (Tol: ±{tol_r:.3f}\", {fila['% de Riesgo']:.1f}%)", linewidth=1.5)
                
            plt.axvspan(-TOLERANCIA_INTERNA, TOLERANCIA_INTERNA, color='green', alpha=0.04)
            plt.axvline(0, color='darkgreen', linestyle='--')
            plt.axvline(TOLERANCIA_INTERNA, color='red', linestyle=':')
            plt.axvline(-TOLERANCIA_INTERNA, color='red', linestyle=':')
            plt.title(f"{mat} - {calibre} - Nominal: {nominal:.3f}\"", fontsize=10, color='#EC2024', weight='bold')
            plt.legend(loc="upper right", fontsize=7)
            plt.grid(True, linestyle=':', alpha=0.5)
            plt.tight_layout()
            
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=180)
            img_buffer.seek(0)
            plt.close()
            
            img_pdf = RLImage(img_buffer, width=420, height=180)
            img_pdf.hAlign = 'CENTER'
            story.append(img_pdf)
            story.append(Spacer(1, 10))
        except Exception as e:
            story.append(Paragraph(f"<i>No se pudo renderizar gráfico para {mat} {calibre}: {str(e)}</i>", c_style))

    story.append(Spacer(1, 15))
    
    if cert_file_data or email_img_data:
        story.append(PageBreak())
        story.append(Paragraph("4. Documentos de Respaldo", h2_st))
        
        row_titles = []
        row_images = []
        
        # 4.1 Certificado
        if cert_file_data:
            row_titles.append(Paragraph("4.1. Certificado de Calidad", h3_st))
            try:
                import fitz
                doc_pdf = fitz.open(stream=cert_file_data, filetype="pdf")
                page = doc_pdf.load_page(0)
                pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
                img_buffer = io.BytesIO(pix.tobytes("png"))
                cert_img = RLImage(img_buffer, width=250, height=250, kind='proportional')
                row_images.append(cert_img)
            except Exception as e:
                row_images.append(Paragraph(f"<i>Error: {e}</i>", c_style))
            
        # 4.2 Correo
        if email_img_data:
            row_titles.append(Paragraph("4.2. Correo de Compras", h3_st))
            try:
                img_buffer = io.BytesIO(email_img_data)
                email_img = RLImage(img_buffer, width=250, height=250, kind='proportional')
                row_images.append(email_img)
            except Exception as e:
                row_images.append(Paragraph(f"<i>Error: {e}</i>", c_style))
                
        if row_titles:
            if len(row_titles) == 2:
                t_docs = Table([row_titles, row_images], colWidths=[260, 260])
            else:
                t_docs = Table([row_titles, row_images], colWidths=[520])
            t_docs.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
            story.append(KeepTogether(t_docs))
            
        story.append(Spacer(1, 15))
        
    if df_raw is not None and not df_raw.empty:
        story.append(PageBreak())
        story.append(Paragraph("Anexo: Base de Datos Original (Importación de Excel)", h2_st))
        
        # Build table data from df_raw
        raw_headers = [Paragraph(str(c).replace("_", " "), h_style) for c in df_raw.columns]
        raw_data = [raw_headers]
        for _, row in df_raw.iterrows():
            row_data = [Paragraph(str(val), c_style) for val in row]
            raw_data.append(row_data)
            
        num_cols = len(df_raw.columns)
        col_w = 520 / num_cols if num_cols > 0 else 100
        
        t_raw = Table(raw_data, colWidths=[col_w]*num_cols, repeatRows=1)
        t_raw.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#111111')), 
            ('ALIGN', (0,0), (-1,-1), 'CENTER'), 
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3'))
        ]))
        story.append(t_raw)
        story.append(Spacer(1, 15))

    f_st = ParagraphStyle('FText', fontName='Helvetica', fontSize=10, alignment=1, spaceAfter=2)
    story.append(Paragraph("___________________________________________________", f_st))
    story.append(Paragraph("<b>Ing. Jesús Morales</b>", f_st))
    story.append(Paragraph("Dir. Planta Metales | SIGRAMA PLANTA METALES", f_st))
    
    doc.build(story)
    buffer.seek(0)
    return buffer
# Inyección de Estilos CSS Corporativos Oficiales
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;700&family=Questrial&display=swap');

    /* Fuentes globales y fondo claro */
    html, body, [class*="css"], .stApp {
        font-family: 'Questrial', sans-serif !important;
        background-color: #FFFFFF !important;
    }

    h1, h2, h3, h4, h5, h6, .main-title {
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        color: #111111 !important;
    }

    /* Barra lateral corporativa en Negro profundo #111111 */
    [data-testid="stSidebar"] {
        background-color: #111111 !important;
        border-right: 1px solid #1E293B !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p, 
    [data-testid="stSidebar"] span, 
    [data-testid="stSidebar"] label {
        color: #FFFFFF !important;
        font-family: 'Questrial', sans-serif !important;
    }
    
    /* Botones de navegación en barra lateral */
    [data-testid="stSidebar"] div[role="radiogroup"] label {
        color: #FFFFFF !important;
        font-size: 14px !important;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        color: #EC2024 !important;
    }

    /* Estilos para inputs de contraseña o textos en barra lateral */
    [data-testid="stSidebar"] input {
        background-color: #1E293B !important;
        color: #FFFFFF !important;
        border-color: #334155 !important;
    }
    [data-testid="stSidebar"] input:focus {
        border-color: #EC2024 !important;
    }

    /* Estilo de Botones Oficiales - Rojo Corporativo #EC2024 */
    div.stButton > button,
    div.stDownloadButton > button,
    div.stFormSubmitButton > button,
    button[data-testid="baseButton-secondary"]:not([role="tab"]):not([data-baseweb="tab"]),
    button[data-testid="baseButton-primary"]:not([role="tab"]):not([data-baseweb="tab"]),
    button[kind="secondary"]:not([role="tab"]):not([data-baseweb="tab"]),
    button[kind="primary"]:not([role="tab"]):not([data-baseweb="tab"]) {
        background-color: #EC2024 !important;
        color: #FFFFFF !important;
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        border-radius: 4px !important;
        border: 1px solid #EC2024 !important;
        padding: 8px 20px !important;
        transition: all 0.3s ease !important;
        text-transform: uppercase;
        font-size: 13px !important;
    }
    div.stButton > button:hover,
    div.stDownloadButton > button:hover,
    div.stFormSubmitButton > button:hover,
    button[data-testid="baseButton-secondary"]:not([role="tab"]):not([data-baseweb="tab"]):hover,
    button[data-testid="baseButton-primary"]:not([role="tab"]):not([data-baseweb="tab"]):hover,
    button[kind="secondary"]:not([role="tab"]):not([data-baseweb="tab"]):hover,
    button[kind="primary"]:not([role="tab"]):not([data-baseweb="tab"]):hover {
        background-color: #FFFFFF !important;
        color: #EC2024 !important;
        border: 1px solid #EC2024 !important;
        box-shadow: 0 4px 12px rgba(236, 32, 36, 0.15) !important;
    }

    /* Tarjetas de Métricas */
    [data-testid="metric-container"] {
        background-color: #FFFFFF !important;
        border: 1px solid #D2D3D5 !important;
        border-left: 5px solid #EC2024 !important;
        border-radius: 4px !important;
        padding: 12px 18px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05) !important;
    }
    [data-testid="metric-container"] label {
        font-family: 'Montserrat', sans-serif !important;
        color: #111111 !important;
        font-weight: 500 !important;
    }
    [data-testid="metric-container"] div[data-testid="stMetricValue"] {
        font-family: 'Montserrat', sans-serif !important;
        color: #EC2024 !important;
        font-weight: 700 !important;
    }
    
    /* Configuración del Editor de Datos y Tablas */
    .stTable header, th {
        background-color: #111111 !important;
        color: #FFFFFF !important;
        font-family: 'Montserrat', sans-serif !important;
    }
    
    /* Inputs y Selectores */
    div[data-baseweb="input"], div[data-baseweb="select"], textarea {
        border-color: #D2D3D5 !important;
        border-radius: 4px !important;
    }
    div[data-baseweb="input"]:focus-within, div[data-baseweb="select"]:focus-within {
        border-color: #EC2024 !important;
    }

    /* Reset del File Uploader para encajar en el estilo secundario */
    [data-testid="stFileUploader"] button,
    .stFileUploader button {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: 1px solid #D2D3D5 !important;
        border-radius: 4px !important;
        padding: 6px 12px !important;
        font-weight: 500 !important;
        box-shadow: none !important;
        transform: none !important;
    }
    
    [data-testid="stFileUploader"] button *,
    .stFileUploader button * {
        background-color: transparent !important;
        color: #111111 !important;
    }
    
    [data-testid="stFileUploader"] button:hover,
    .stFileUploader button:hover {
        background-color: #F8F9FA !important;
        border-color: #EC2024 !important;
        color: #EC2024 !important;
    }
    
    [data-testid="stFileUploader"] button:hover *,
    .stFileUploader button:hover * {
        color: #EC2024 !important;
    }
</style>
""", unsafe_allow_html=True)

import os

# Renderizado de Logo en Barra Lateral
logo_neg_path = os.path.join(BASE_DIR, "logo_sigrama_negative.png")
logo_pos_path = os.path.join(BASE_DIR, "logo_sigrama.png")
if os.path.exists(logo_neg_path):
    st.sidebar.image(logo_neg_path, use_container_width=True)
elif os.path.exists(logo_pos_path):
    st.sidebar.image(logo_pos_path, use_container_width=True)
else:
    st.sidebar.subheader("INDUSTRIA SIGRAMA")

# Sección de Perfil de Usuario
st.sidebar.markdown("""
<div style="background-color: #1E293B; border: 1px solid #334155; padding: 12px; border-radius: 6px; margin-bottom: 15px; margin-top: 10px;">
    <p style="margin: 0; color: #FFFFFF; font-family: 'Questrial', sans-serif; font-size: 13px;">
        👤 Usuario: <b>Ingeniero de Calidad</b>
    </p>
    <p style="margin: 5px 0 0 0; color: #EC2024; font-family: 'Montserrat', sans-serif; font-size: 12px; font-weight: bold;">
        🔑 Rol: Administrador de Planta
    </p>
</div>
""", unsafe_allow_html=True)

st.sidebar.write("---")
opcion_menu = st.sidebar.radio("Módulos del Sistema:", ["📊 Suite de Análisis", "🔍 Historial de Reportes", "🏢 Catálogo de Proveedores"])

st.sidebar.write("---")
st.sidebar.subheader("🛠️ Parámetros de Control")
tol_proveedor = st.sidebar.slider(
    'Tolerancia por Defecto (Fallback) (±):',
    min_value=0.001, max_value=0.008, value=0.006, step=0.001, format="%.3f"
)

st.sidebar.write("---")
st.sidebar.subheader("📋 Acciones y Plantillas")
excel_data = generar_excel_plantilla()
st.sidebar.download_button(
    label="📝 Descargar Plantilla Excel",
    data=excel_data,
    file_name="plantilla_rollos.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.sidebar.write("---")
st.sidebar.subheader("⚠️ Zona de Peligro")
if st.sidebar.button("🗑️ Limpiar Base de Datos", type="primary", use_container_width=True):
    database.limpiar_base_datos()
    st.sidebar.success("Base de datos y registros limpiados con éxito.")
    st.rerun()

st.sidebar.markdown("""
    <div style="text-align: center; margin-top: 50px; padding-top: 20px; border-top: 1px solid #334155;">
        <span style="font-family: 'Questrial', sans-serif; font-style: italic; font-size: 13px; color: #FFFFFF; border-bottom: 2px solid #EC2024; padding-bottom: 4px; display: inline-block;">
            Ingeniería que da resultados!!
        </span>
    </div>
""", unsafe_allow_html=True)

# Slogan corporativo en Main Panel
st.markdown('<p style="text-align: center; font-size: 16px; font-weight: bold; color: #EC2024; font-family: \'Montserrat\', sans-serif; margin-top: 15px; text-transform: uppercase; letter-spacing: 1px;">SOLUCIONES QUE TRANSFORMAN TU EMPRESA</p>', unsafe_allow_html=True)
st.markdown('<hr style="border: 1px solid #EC2024; margin: 15px 0;">', unsafe_allow_html=True)

# Enrutamiento según navegación lateral
if opcion_menu == "📊 Suite de Análisis":
    st.title("⚙️ Suite de Riesgo y Control de Suministros")
    st.markdown(f"**Estándar Fijo Planta (Norma Interna de Diseño):** `±{TOLERANCIA_INTERNA}\"`")

    archivo_cargado = st.file_uploader("📥 Cargar datos industriales para simulación (Excel)", type=["xlsx"])
    if archivo_cargado is not None:
        try:
            df = pd.read_excel(archivo_cargado)
            st.session_state["df_raw_excel"] = df.copy()
            st.session_state["raw_excel_bytes"] = archivo_cargado.getvalue()
            res = []
            
            # Buscar si existe columna de Tolerancia en el Excel
            col_tol = None
            for col in df.columns:
                if col.strip().lower() in ["tolerancia_proveedor", "tolerancia", "tolerancia proveedor", "tolerancia ofertada"]:
                    col_tol = col
                    break
            
            for _, f in df.iterrows():
                mat = str(f["Material"]).strip()
                cal = int(f["Calibre"])
                if mat not in ESTANDAR or cal not in ESTANDAR[mat]:
                    continue
                
                med = float(f["Espesor_Medido"])
                uni = str(f["Unidad"]).strip().lower()
                esp_in = round(med * 0.0393701, 4) if ("mm" in uni or "mili" in uni) else med
                
                # Tolerancia del rollo
                tol_r = float(f[col_tol]) if (col_tol and pd.notna(f[col_tol])) else tol_proveedor
                
                res.append({
                    "Rollo": str(f["Numero_Rollo"]),
                    "Material": mat,
                    "Calibre": f"CAL {cal}",
                    "Espesor Original": f"{med} {'mm' if 'mm' in uni else 'in'}",
                    "Espesor Real (in)": esp_in,
                    "Nominal Estándar": ESTANDAR[mat][cal],
                    "Desviación Real (in)": round(esp_in - ESTANDAR[mat][cal], 4),
                    "Tolerancia_Rollo": tol_r
                })
                
            df_datos_cargados = pd.DataFrame(res)
            
            if not df_datos_cargados.empty:
                # ======================================================================
                # EVALUACIÓN DE RIESGO ESTADÍSTICO POR ROLLO (GAUSS)
                # ======================================================================
                est_l = []
                riesgo_l = []
                dictamen_final_l = []
                
                for _, fila in df_datos_cargados.iterrows():
                    med_individual = fila['Espesor Real (in)']
                    nom_individual = fila['Nominal Estándar']
                    tol_r = fila['Tolerancia_Rollo']
                    sigma_individual = tol_r / 3.0
                    
                    p_inf_ind = stats.norm.cdf(nom_individual - TOLERANCIA_INTERNA, loc=med_individual, scale=sigma_individual)
                    p_sup_ind = 1.0 - stats.norm.cdf(nom_individual + TOLERANCIA_INTERNA, loc=med_individual, scale=sigma_individual)
                    riesgo_rollo_pct = (p_inf_ind + p_sup_ind) * 100
                    
                    riesgo_l.append(riesgo_rollo_pct)
                    
                    # Clasificación micrométrica basada en probabilidad
                    if riesgo_rollo_pct < 1.0:
                        est_l.append("RIESGO BAJO")
                        dictamen_final_l.append("ACEPTADO")
                    elif riesgo_rollo_pct <= 5.0:
                        est_l.append("RIESGO MODERADO")
                        dictamen_final_l.append("ACEPTADO")
                    else:
                        est_l.append("ALTO RIESGO")
                        dictamen_final_l.append("NO ACEPTADO")
                
                # Incorporación de nuevas variables calculadas
                df_datos_cargados['% de Riesgo'] = riesgo_l
                df_datos_cargados['Riesgo'] = est_l
                df_datos_cargados['Dictamen Final'] = dictamen_final_l
                
                # Despliegue de métricas clave (Estilo Tarjetas Corporativas)
                total_rollos = len(df_datos_cargados)
                aceptados = len(df_datos_cargados[df_datos_cargados["Dictamen Final"] == "ACEPTADO"])
                rechazados = len(df_datos_cargados[df_datos_cargados["Dictamen Final"] == "NO ACEPTADO"])
                prom_riesgo = df_datos_cargados["% de Riesgo"].mean()
                
                st.write("---")
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                col_m1.metric("Total Rollos Analizados", f"{total_rollos}")
                col_m2.metric("Rollos Aceptados", f"{aceptados}")
                col_m3.metric("Rollos Rechazados", f"{rechazados}")
                col_m4.metric("Riesgo Promedio", f"{prom_riesgo:.2f}%")
                st.write("---")
                
                # RENDERIZADO TABLA 1 (MODIFICACIÓN 1 EXPOSITIVA)
                st.subheader("📊 Calibración del Muestreo por Unidad (Rollo por Rollo)")
                formatos = {'Espesor Real (in)': '{:.4f}"', 'Nominal Estándar': '{:.3f}"', 'Desviación Real (in)': '{:+.4f}"', '% de Riesgo': '{:.2f}%', 'Tolerancia_Rollo': '{:.3f}"'}
                styler_individual = df_datos_cargados.style.format(formatos).map(colorear_matriz_resumen, subset=['Riesgo'])
                st.dataframe(styler_individual, use_container_width=True)
                
                # ======================================================================
                # CLASIFICACIÓN JERÁRQUICA Y SE REMOVE "ESPESOR ORIGINAL"
                # ======================================================================
                st.subheader("📋 Análisis Clasificado Estructurado por Espesor Nominal Teórico")
                df_grouped = df_datos_cargados.groupby(['Material', 'Calibre', 'Nominal Estándar'])
                
                for (material, calibre, nominal), grupo in df_grouped:
                    # Encabezado corregido con especificaciones y tolerancia aceptable
                    st.markdown(f"#### 🌐 {material} — `{calibre}` — Espesor Teórico: `{nominal:.3f}\"` | Tolerancia Aceptable: `±{TOLERANCIA_INTERNA:.3f}\"`")
                    
                    # SE EXCLUYE COMPLETAMENTE 'Espesor Original' para la vista solicitada
                    columnas_vista = ['Rollo', 'Espesor Real (in)', 'Desviación Real (in)', 'Tolerancia_Rollo', '% de Riesgo', 'Dictamen Final']
                    df_vista_grupo = grupo[columnas_vista]
                    
                    # Renderizado con colores condicionales (Verde = ACEPTADO, Rojo = NO ACEPTADO)
                    styler_grupo = df_vista_grupo.style.format(formatos).map(colorear_matriz_resumen, subset=['Dictamen Final'])
                    st.dataframe(styler_grupo, use_container_width=True)
                # ======================================================================
                # GENERACIÓN DE UNA GRÁFICA AISLADA POR CADA ESPESOR
                # ======================================================================
                st.subheader("📈 Distribuciones Probabilísticas por Especificación Técnica")
                
                for (material, calibre, nominal), grupo in df_grouped:
                    st.markdown(f"##### Análisis Gaussiano Aislado: `{material} — {calibre} ({nominal:.3f}\")`")
                    
                    fig = go.Figure()
                    x_desv = np.linspace(-0.015, 0.015, 400)
                    
                    # Inyección exclusiva de curvas correspondientes a este espesor particular
                    for _, fila in grupo.iterrows():
                        m_desv = fila['Espesor Real (in)'] - nominal
                        tol_r = fila.get('Tolerancia_Rollo', tol_proveedor)
                        sigma_individual = tol_r / 3.0
                        y_g = stats.norm.pdf(x_desv, loc=m_desv, scale=sigma_individual)
                        fig.add_trace(go.Scatter(
                            x=x_desv, y=y_g, mode='lines', 
                            name=f"{fila['Rollo']} (Tol: ±{tol_r:.3f}\", Riesgo: {fila['% de Riesgo']:.1f}%)",
                            line=dict(width=2.5)
                        ))
                    
                    # Delimitadores de franja interna de diseño de planta
                    fig.add_vrect(x0=-TOLERANCIA_INTERNA, x1=TOLERANCIA_INTERNA, fillcolor="green", opacity=0.05, line_width=0)
                    fig.add_vline(x=0, line_dash="dash", line_color="darkgreen")
                    fig.add_vline(x=TOLERANCIA_INTERNA, line_color="red", line_width=1.5, line_dash="dot")
                    fig.add_vline(x=-TOLERANCIA_INTERNA, line_color="red", line_width=1.5, line_dash="dot")
                    
                    fig.update_layout(
                        xaxis_title="Desviación Micrométrica Real (in)", 
                        yaxis_title="Densidad Probabilística de Gauss", 
                        height=300, 
                        margin=dict(l=40, r=40, t=15, b=40),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )
                    st.plotly_chart(fig, use_container_width=True)
                
                # --- COMPONENTE DE ACCIÓN FINAL: BOTÓN DE DESCARGA PDF ---
                st.subheader("📄 Entregables de Ingeniería de Calidad")
                
                draft_meta = {
                    "Folio": "BORRADOR (No Guardado)",
                    "Fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "Proveedor": "Pendiente (Se asigna al guardar)",
                    "Contacto": "Pendiente",
                    "Certificado": "Pendiente de Carga"
                }
                
                pdf_bytes = crear_pdf_formal(df_datos_cargados, tol_proveedor, df_raw=st.session_state.get("df_raw_excel"), meta_info=draft_meta)
                st.download_button(
                    label="📄 DESCARGAR REPORTE PDF DE CONTROL CORPORATIVO",
                    data=pdf_bytes.getvalue(),
                    file_name=f"Reporte_Riesgo_Espesores_{datetime.now().strftime('%H%M')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
                # ======================================================================
                # FORMULARIO DE GUARDADO Y REGISTRO EN EL HISTORIAL
                # ======================================================================
                st.write("---")
                st.subheader("💾 Guardar Reporte en el Historial del Sistema")
                st.markdown("Guarde los resultados del análisis de la propuesta del proveedor junto con los documentos de respaldo para auditoría y rastreo futuro.")
                
                col_h1, col_h2 = st.columns(2)
                with col_h1:
                    st.write(f"DEBUG DB PATH: {getattr(database, '__file__', 'No file')}")
                    listado_proveedores = database.listar_proveedores_nombres()
                    if listado_proveedores:
                        prov_input = st.selectbox("🏢 Proveedor Ofertante:", listado_proveedores, key="prov_input_h")
                    else:
                        st.warning("⚠️ No hay proveedores registrados. Registre uno primero en el Catálogo de Proveedores.")
                        prov_input = ""
                with col_h2:
                    cert_info_input = st.text_input("📑 Información del Certificado (ID/Número):", value="", placeholder="Ej. Certificado N° TX-98810", key="cert_info_input_h")
                    
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    cert_file_input = st.file_uploader("📂 Archivo de Certificado de Calidad (PDF)", type=["pdf"], key="cert_file_h")
                with col_f2:
                    st.write("📧 Captura del Correo de Compras (Imagen)")
                    
                    email_img_data = None
                    email_img_name = None
                    
                    col_img_file, col_img_paste = st.columns([0.5, 0.5])
                    with col_img_file:
                        email_img_input = st.file_uploader("Subir imagen:", type=["png", "jpg", "jpeg"], label_visibility="collapsed", key="email_img_h")
                        if email_img_input is not None:
                            email_img_data = email_img_input.read()
                            email_img_name = email_img_input.name
                    with col_img_paste:
                        from streamlit_paste_button import paste_image_button as pbutton
                        paste_result = pbutton(
                            label="📋 PEGAR DE PORTAPAPELES",
                            text_color="#FFFFFF",
                            background_color="#EC2024",
                            hover_background_color="#111111",
                            errors="ignore"
                        )
                        if paste_result.image_data is not None:
                            img_byte_arr = io.BytesIO()
                            paste_result.image_data.save(img_byte_arr, format='PNG')
                            email_img_data = img_byte_arr.getvalue()
                            email_img_name = "Captura_Portapapeles.png"
                    
                    if email_img_data is not None:
                        st.image(email_img_data, caption=f"Imagen cargada: {email_img_name}", width=250)
                            
                btn_save = st.button("Confirmar y Guardar en Base de Datos", key="btn_save_h")
                
                if btn_save:
                    if not prov_input.strip():
                        st.error("❌ El nombre del proveedor es obligatorio.")
                    elif not cert_file_input:
                        st.error("❌ El archivo del Certificado de Calidad en formato PDF es obligatorio.")
                    elif email_img_data is None:
                        st.error("❌ La captura de pantalla del correo de Compras es obligatoria (cárguela por archivo o péguela desde el portapapeles).")
                    else:
                        with st.spinner("Guardando registro y archivos en el expediente..."):
                            import database
                            
                            # 1. Generar Folio
                            nuevo_folio = database.generar_siguiente_folio()
                            folder_exp = os.path.join(database.EXPEDIENTES_DIR, nuevo_folio)
                            os.makedirs(folder_exp, exist_ok=True)
                            
                            # 2. Guardar archivos cargados
                            cert_ext = os.path.splitext(cert_file_input.name)[1]
                            email_ext = ".png" if email_img_name == "Captura_Portapapeles.png" else os.path.splitext(email_img_name)[1]
                            
                            ruta_cert_dest = os.path.join(folder_exp, f"{nuevo_folio} - CERTIFICADO{cert_ext}")
                            ruta_correo_dest = os.path.join(folder_exp, f"{nuevo_folio} - IMAGEN{email_ext}")
                            ruta_reporte_dest = os.path.join(folder_exp, f"{nuevo_folio} - REPORTE.pdf")
                            ruta_excel_dest = os.path.join(folder_exp, f"{nuevo_folio} - DATOS.xlsx")
                            
                            with open(ruta_cert_dest, "wb") as f_out:
                                f_out.write(cert_file_input.read())
                            with open(ruta_correo_dest, "wb") as f_out:
                                f_out.write(email_img_data)
                            if "raw_excel_bytes" in st.session_state:
                                with open(ruta_excel_dest, "wb") as f_out:
                                    f_out.write(st.session_state["raw_excel_bytes"])
                                
                            # Construir Metadata para el PDF
                            contact_info = "N/D"
                            for p in database.listar_proveedores():
                                if p["nombre"] == prov_input:
                                    # Formatear la cadena de contacto omitiendo partes vacías
                                    parts = [p['contacto']]
                                    if p['correo']: parts.append(p['correo'])
                                    if p['telefono']: parts.append(p['telefono'])
                                    contact_info = " / ".join([str(x) for x in parts if str(x).strip()])
                                    break
                                    
                            meta_info = {
                                "Folio": nuevo_folio,
                                "Fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
                                "Proveedor": prov_input,
                                "Contacto": contact_info,
                                "Certificado": cert_info_input
                            }
                                
                            # 3. Generar y guardar PDF de reporte técnico
                            report_pdf_bytes = crear_pdf_formal(df_datos_cargados, tol_proveedor, cert_file_input.getvalue(), email_img_data, df_raw=st.session_state.get("df_raw_excel"), meta_info=meta_info)
                            with open(ruta_reporte_dest, "wb") as f_out:
                                f_out.write(report_pdf_bytes.getvalue())
                                
                            # 4. Registrar en base de datos
                            # Rutas relativas para portabilidad de almacenamiento
                            rel_cert = os.path.relpath(ruta_cert_dest, database.BASE_DIR)
                            rel_correo = os.path.relpath(ruta_correo_dest, database.BASE_DIR)
                            rel_reporte = os.path.relpath(ruta_reporte_dest, database.BASE_DIR)
                            
                            fecha_hoy = datetime.now().strftime("%d/%m/%Y")
                            
                            database.guardar_reporte(
                                folio=nuevo_folio,
                                fecha=fecha_hoy,
                                proveedor=prov_input.strip(),
                                certificado_info=cert_info_input.strip(),
                                ruta_certificado=rel_cert,
                                ruta_correo=rel_correo,
                                ruta_reporte=rel_reporte,
                                desviacion_ofertada_def=tol_proveedor,
                                total_rollos=total_rollos,
                                aceptados=aceptados,
                                rechazados=rechazados,
                                riesgo_promedio=prom_riesgo
                            )
                            st.success(f"✅ Reporte guardado exitosamente bajo el Folio: **{nuevo_folio}**")
                                
        except Exception as e:
            st.error(f"❌ Error crítico en el procesamiento del lote técnico: {str(e)}")
    else:
        st.info("💡 Tablero listo. Por favor, cargue un archivo de Excel utilizando la plantilla estándar en la barra lateral para iniciar las simulaciones estadísticas.")

elif opcion_menu == "🔍 Historial de Reportes":
    import database
    st.title("🔍 Historial de Consultas de Propuestas")
    st.markdown("Busque y consulte expedientes históricos de propuestas de proveedores cargados en el sistema.")
    
    # 1. Filtros
    with st.expander("🔍 Filtros de Búsqueda", expanded=True):
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            rango_fecha = st.date_input("Rango de Fechas:", value=(datetime.now() - pd.Timedelta(days=90), datetime.now()))
        with col_f2:
            proveedores_list = ["Todos"] + database.obtener_proveedores()
            prov_filtro = st.selectbox("Filtrar por Proveedor:", proveedores_list)
            
    # Parsear rango de fechas
    fecha_ini = None
    fecha_fi = None
    if isinstance(rango_fecha, tuple) and len(rango_fecha) == 2:
        fecha_ini = rango_fecha[0].strftime("%Y-%m-%d")
        fecha_fi = rango_fecha[1].strftime("%Y-%m-%d")
    elif isinstance(rango_fecha, list) and len(rango_fecha) == 2:
        fecha_ini = rango_fecha[0].strftime("%Y-%m-%d")
        fecha_fi = rango_fecha[1].strftime("%Y-%m-%d")
    elif isinstance(rango_fecha, datetime):
        fecha_ini = rango_fecha.strftime("%Y-%m-%d")
        fecha_fi = rango_fecha.strftime("%Y-%m-%d")
        
    # Obtener datos
    records = database.obtener_reportes(fecha_inicio=fecha_ini, fecha_fin=fecha_fi, proveedor=prov_filtro)
    
    if not records:
        st.info("No se encontraron registros en el historial con los filtros aplicados.")
    else:
        df_hist = pd.DataFrame(records)
        # Mostrar tabla resumida
        st.write("### Resumen de Expedientes Encontrados")
        df_hist_view = df_hist[[
            "folio", "fecha", "proveedor", "certificado_info", 
            "total_rollos", "aceptados", "rechazados", "riesgo_promedio"
        ]].rename(columns={
            "folio": "Folio",
            "fecha": "Fecha",
            "proveedor": "Proveedor",
            "certificado_info": "Certificado",
            "total_rollos": "Rollos Totales",
            "aceptados": "Aceptados",
            "rechazados": "Rechazados",
            "riesgo_promedio": "Riesgo Promedio"
        })
        st.dataframe(df_hist_view, use_container_width=True, hide_index=True)
        
        st.write("---")
        st.write("### 📁 Detalle de Expediente")
        folio_sel = st.selectbox("Seleccione un Folio para visualizar y descargar:", df_hist["folio"].tolist())
        
        if folio_sel:
            rec_sel = df_hist[df_hist["folio"] == folio_sel].iloc[0]
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.markdown(f"#### Folio: **{rec_sel['folio']}**")
                st.write(f"- **Fecha:** {rec_sel['fecha']}")
                st.write(f"- **Proveedor:** {rec_sel['proveedor']}")
                st.write(f"- **Certificado:** {rec_sel['certificado_info']}")
                st.write(f"- **Total Rollos:** {rec_sel['total_rollos']}")
                st.write(f"- **Aceptados:** {rec_sel['aceptados']} | **Rechazados:** {rec_sel['rechazados']}")
                st.write(f"- **Riesgo Promedio:** {rec_sel['riesgo_promedio']:.2f}%")
                
                # Descargas
                st.write("##### Descargar Documentos:")
                
                # Reporte PDF
                rep_path = os.path.join(database.BASE_DIR, rec_sel["ruta_reporte"])
                if os.path.exists(rep_path):
                    rep_name = os.path.basename(rep_path)
                    with open(rep_path, "rb") as f_pdf:
                        st.download_button(
                            label="📄 Descargar Reporte Técnico (PDF)",
                            data=f_pdf.read(),
                            file_name=rep_name,
                            mime="application/pdf",
                            use_container_width=True,
                            key=f"btn_dl_rep_{rec_sel['folio']}"
                        )
                else:
                    st.warning("⚠️ Archivo de reporte no encontrado.")
                    
                # Certificado PDF
                cert_path = os.path.join(database.BASE_DIR, rec_sel["ruta_certificado"])
                if os.path.exists(cert_path):
                    cert_name = os.path.basename(cert_path)
                    with open(cert_path, "rb") as f_pdf:
                        st.download_button(
                            label="📂 Descargar Certificado de Calidad (PDF)",
                            data=f_pdf.read(),
                            file_name=cert_name,
                            mime="application/pdf",
                            use_container_width=True,
                            key=f"btn_dl_cert_{rec_sel['folio']}"
                        )
                else:
                    st.warning("⚠️ Archivo de certificado no encontrado.")
                    
                # Datos Excel Originales
                excel_path = os.path.join(database.BASE_DIR, "expedientes", rec_sel["folio"], f"{rec_sel['folio']} - DATOS.xlsx")
                if os.path.exists(excel_path):
                    with open(excel_path, "rb") as f_xls:
                        st.download_button(
                            label="📊 Descargar Datos Base (Excel)",
                            data=f_xls.read(),
                            file_name=os.path.basename(excel_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"btn_dl_xls_{rec_sel['folio']}"
                        )
                    
            with col_d2:
                # Mostrar imagen del correo de compras
                correo_path = os.path.join(database.BASE_DIR, rec_sel["ruta_correo"])
                if os.path.exists(correo_path):
                    st.write("##### Captura del Correo de Compras:")
                    st.image(correo_path, use_container_width=True)
                else:
                    st.warning("⚠️ Captura del correo de compras no encontrada.")
                    
            st.write("---")
            st.write("##### 🗑️ Zona de Peligro: Eliminar Expediente")
            conf_eliminar_exp = st.checkbox(f"Confirmo que deseo eliminar definitivamente el expediente **{rec_sel['folio']}** de la base de datos.", key=f"chk_eliminar_{rec_sel['folio']}")
            if conf_eliminar_exp:
                if st.button("ELIMINAR EXPEDIENTE", type="primary", key=f"btn_eliminar_exp_{rec_sel['folio']}"):
                    database.eliminar_reporte(rec_sel['folio'])
                    st.rerun()

elif opcion_menu == "🏢 Catálogo de Proveedores":
    importlib.reload(database)
    st.title("🏢 Catálogo de Proveedores de Materia Prima")
    st.markdown("Registre, consulte y administre los proveedores oficiales de la planta.")
    
    # 1. Crear Nuevo Proveedor
    with st.expander("➕ Registrar Nuevo Proveedor", expanded=False):
        with st.form("form_registro_proveedor", clear_on_submit=True):
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                nombre_prov = st.text_input("Nombre del Proveedor (Obligatorio):", placeholder="Ej. Ternium, Nucor, etc.")
                contacto_prov = st.text_input("Nombre del Contacto:", placeholder="Ej. Ing. Jesús Morales")
            with col_p2:
                tel_prov = st.text_input("Teléfono de Contacto:", placeholder="Ej. 81-1234-5678")
                correo_prov = st.text_input("Correo Electrónico:", placeholder="Ej. contacto@proveedor.com")
                
            btn_registrar = st.form_submit_button("Guardar Proveedor")
            
            if btn_registrar:
                if not nombre_prov.strip():
                    st.error("❌ El nombre del proveedor es obligatorio.")
                else:
                    success = database.crear_proveedor(
                        nombre=nombre_prov.strip(),
                        contacto=contacto_prov.strip(),
                        telefono=tel_prov.strip(),
                        correo=correo_prov.strip()
                    )
                    if success:
                        st.success(f"✅ Proveedor **{nombre_prov.strip()}** registrado exitosamente.")
                        st.rerun()
                    else:
                        st.error("❌ Error: Ya existe un proveedor registrado con ese nombre.")
                        
    # 2. Listado de Proveedores Registrados
    st.write("### 📋 Proveedores Oficiales Registrados")
    listado = database.listar_proveedores()
    
    if not listado:
        st.info("No hay proveedores registrados en el catálogo.")
    else:
        df_provs = pd.DataFrame(listado)
        df_provs_view = df_provs[["id", "nombre", "contacto", "telefono", "correo", "fecha_registro"]].rename(columns={
            "id": "ID",
            "nombre": "Nombre del Proveedor",
            "contacto": "Contacto",
            "telefono": "Teléfono",
            "correo": "Correo Electrónico",
            "fecha_registro": "Fecha de Registro"
        })
        st.dataframe(df_provs_view, use_container_width=True, hide_index=True)
        
        # 3. Eliminar Proveedor
        st.write("---")
        st.write("### 🗑️ Eliminar Proveedor del Catálogo")
        prov_a_eliminar = st.selectbox(
            "Seleccione el proveedor que desea eliminar:",
            options=df_provs["nombre"].tolist(),
            key="selectbox_eliminar_prov"
        )
        
        if prov_a_eliminar:
            row_prov = df_provs[df_provs["nombre"] == prov_a_eliminar].iloc[0]
            
            # Confirmación
            confirm_eliminar = st.checkbox(f"Confirmo que deseo eliminar definitivamente a **{prov_a_eliminar}** del sistema.", key="check_eliminar_prov")
            if confirm_eliminar:
                btn_eliminar = st.button("Eliminar Proveedor", type="primary", key="btn_eliminar_prov")
                if btn_eliminar:
                    # Castear explicitamente a int nativo de Python para evitar fallos silenciosos de SQLite con numpy.int64
                    id_prov_int = int(row_prov["id"])
                    database.eliminar_proveedor(id_prov_int)
                    st.rerun()

