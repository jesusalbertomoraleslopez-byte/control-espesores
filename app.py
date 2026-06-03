import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import scipy.stats as stats
from datetime import datetime
import io

# 1. CONFIGURACIÓN ÚNICA DE LA PÁGINA (Previene errores de duplicación de Streamlit)
st.set_page_config(page_title="Suite de Riesgo y Control de Espesores", layout="wide")

# 2. CONFIGURACIÓN GRÁFICA VECTORIAL (Para uso local o despliegue seguro en la nube)
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

import matplotlib
matplotlib.use('Agg') # Renderizado matemático en backend sin requerir servidor gráfico X11
import matplotlib.pyplot as plt

# 3. MATRIZ TÉCNICA PLANTA (Límites rígidos de diseño nominal)
ESTANDAR = {
    "Galvanizado": {10: 0.138, 12: 0.108, 14: 0.079, 16: 0.064},
    "Decapado": {12: 0.105, 14: 0.075, 16: 0.060}
}
TOLERANCIA_INTERNA = 0.008
def colorear_matriz_resumen(v):
    """Aplica formato semafórico condicional basado en las cadenas de texto de riesgo y aceptación."""
    if not isinstance(v, str): return ''
    if "BAJO" in v or "ACEPTADO" == v: return 'background-color: #C6EFCE; color: #006100; font-weight: bold;'
    if "MODERADO" in v: return 'background-color: #FFF2CC; color: #7F6000; font-weight: bold;'
    return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold;'

def generar_excel_plantilla():
    """Construye el archivo binario de Excel en la memoria RAM para descarga dinámica inmediata."""
    datos = {
        "Numero_Rollo": ["ROLLO-A", "ROLLO-B", "ROLLO-C"],
        "Material": ["Galvanizado", "Galvanizado", "Decapado"],
        "Calibre": [12, 16, 14],
        "Espesor_Medido": [0.1028, 1.47, 1.85],
        "Unidad": ["Pulgadas", "Milimetros", "Milimetros"]
    }
    output = io.BytesIO()
    df_tpl = pd.DataFrame(datos)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_tpl.to_excel(writer, index=False)
    return output.getvalue()
def crear_pdf_formal(df_final, tol_p):
    """Genera la estructura del documento técnico formal usando un buffer en memoria RAM."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    styles = getSampleStyleSheet()
    
    t_st = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#2b579a'), spaceAfter=12)
    m_st = ParagraphStyle('DocMeta', fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#444'), spaceAfter=4)
    h2_st = ParagraphStyle('SectionH2', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#2b579a'), spaceBefore=12, spaceAfter=6)
    h3_st = ParagraphStyle('SectionH3', fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor('#333333'), spaceBefore=8, spaceAfter=4)
    h_style = ParagraphStyle('HStyle', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1)
    c_style = ParagraphStyle('CStyle', fontName='Helvetica', fontSize=9, alignment=1)
    
    story.append(Paragraph("SIGRAMA PLANTA METALES", t_st))
    story.append(Paragraph("<b>Documento:</b> Reporte Técnico de Ingeniería de Calidad y Evaluación de Suministro", m_st))
    story.append(Paragraph(f"<b>Fecha de Análisis:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", m_st))
    story.append(Paragraph(f"<b>Parámetro Comercial:</b> Desviación Ofertada por Proveedor en ±{tol_p:.3f}\"", m_st))
    story.append(Spacer(1, 10))
    
    # 1. Muestreo por Unidad General
    story.append(Paragraph("1. Calibración del Muestreo por Unidad (Rollo por Rollo)", h2_st))
    t_rollos_d = [[Paragraph("Rollo", h_style), Paragraph("Material", h_style), Paragraph("Calibre", h_style), Paragraph("Espesor (in)", h_style), Paragraph("Riesgo %", h_style), Paragraph("Riesgo", h_style)]]
    for idx, f in df_final.reset_index(drop=True).iterrows():
        t_rollos_d.append([
            Paragraph(str(f['Rollo']), c_style), Paragraph(f['Material'], c_style), Paragraph(str(f['Calibre']), c_style),
            Paragraph(f"{f['Espesor Real (in)']:.4f}\"", c_style), Paragraph(f"{f['% de Riesgo']:.2f}%", c_style), Paragraph(f['Riesgo'], c_style)
        ])
    t_1 = Table(t_rollos_d, colWidths=[80, 100, 60, 90, 80, 100])
    t_1.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2b579a')), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3'))]))
    story.append(t_1)
    
    # 2. Análisis clasificado por Espesor Nominal (Modificación 2 en Reporte)
    story.append(Paragraph("2. Análisis Estructurado y Clasificación por Espesor Nominal", h2_st))
    
    df_grouped = df_final.groupby(['Material', 'Nominal Estándar'])
    for (mat, nominal), group in df_grouped:
        story.append(Paragraph(f"Especificación: {mat} - Espesor Nominal Teórico: {nominal:.3f}\"", h3_st))
        t_group_d = [[Paragraph("Número Rollo", h_style), Paragraph("Espesor Medido (in)", h_style), Paragraph("Desviación Real", h_style), Paragraph("Probabilidad de Fallo", h_style), Paragraph("Dictamen Final", h_style)]]
        
        for _, fila in group.iterrows():
            t_group_d.append([
                Paragraph(str(fila['Rollo']), c_style), Paragraph(f"{fila['Espesor Real (in)']:.4f}\"", c_style),
                Paragraph(f"{fila['Desviación Real (in)']:+4f}\"", c_style), Paragraph(f"{fila['% de Riesgo']:.2f}%", c_style), Paragraph(fila['Dictamen Final'], c_style)
            ])
        t_g = Table(t_group_d, colWidths=[100, 110, 100, 100, 100])
        t_g.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#708090')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#A0A0A0'))]))
        story.append(t_g)
        story.append(Spacer(1, 8))
        
    # 3. Sección Gráficas Individuales (Modificación 3 en Reporte)
    story.append(PageBreak())
    story.append(Paragraph("3. Análisis de Distribución Probabilística por Especificación Técnica", h2_st))
    
    sigma_p = tol_p / 3.0
    for (mat, nominal), group in df_grouped:
        try:
            plt.figure(figsize=(6.5, 3))
            x_desv = np.linspace(-0.015, 0.015, 400)
            
            for _, fila in group.iterrows():
                m_desv = fila['Espesor Real (in)'] - nominal
                y_g = stats.norm.pdf(x_desv, loc=m_desv, scale=sigma_p)
                plt.plot(x_desv, y_g, label=f"{fila['Rollo']} ({fila['% de Riesgo']:.1f}%)", linewidth=1.5)
                
            plt.axvspan(-TOLERANCIA_INTERNA, TOLERANCIA_INTERNA, color='green', alpha=0.04)
            plt.axvline(0, color='darkgreen', linestyle='--')
            plt.axvline(TOLERANCIA_INTERNA, color='red', linestyle=':')
            plt.axvline(-TOLERANCIA_INTERNA, color='red', linestyle=':')
            plt.title(f"Curvas para {mat} - Nominal: {nominal:.3f}\"", fontsize=10, color='#2b579a', weight='bold')
            plt.legend(loc="upper right", fontsize=7)
            plt.grid(True, linestyle=':', alpha=0.5)
            plt.tight_layout()
            
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=180)
            img_buffer.seek(0)
            plt.close()
            
            img_pdf = RLImage(img_buffer, width=420, height=195)
            img_pdf.hAlign = 'CENTER'
            story.append(Paragraph(f"<b>Distribución Micro-Muestral:</b> {mat} {nominal:.3f}\"", styles['Normal']))
            story.append(img_pdf)
            story.append(Spacer(1, 10))
        except Exception as e:
            story.append(Paragraph(f"<i>No se pudo renderizar gráfico para {mat} {nominal:.3f}\": {str(e)}</i>", c_style))

    story.append(Spacer(1, 15))
    f_st = ParagraphStyle('FText', fontName='Helvetica', fontSize=10, alignment=1, spaceAfter=2)
    story.append(Paragraph("___________________________________________________", f_st))
    story.append(Paragraph("<b>Ing. Jesús Morales</b>", f_st))
    story.append(Paragraph("Dir. Planta Metales | SIGRAMA PLANTA METALES", f_st))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# Despliegue de banner corporativo principal
st.image("BANNER CONTROL DE ESPESORES APP.png", use_container_width=True)
st.title("⚙️ Suite Interactiva de Riesgo y Control de Suministros")
st.markdown(f"**Estándar Fijo Planta (Norma Interna de Diseño):** `±{TOLERANCIA_INTERNA}\"`")

col_control1, col_control2 = st.columns(2)

with col_control1:
    tol_proveedor = st.slider(
        'Desviación Proveedor Ofertada (±):',
        min_value=0.001, max_value=0.008, value=0.006, step=0.001, format="%.3f"
    )

with col_control2:
    st.markdown("**Flujo de Trabajo Corporativo**")
    excel_data = generar_excel_plantilla()
    st.download_button(
        label="📝 Descargar Plantilla Excel Estándar",
        data=excel_data,
        file_name="plantilla_rollos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

archivo_cargado = st.file_uploader("📥 Cargar datos industriales para simulación (Excel)", type=["xlsx"])
if archivo_cargado is not None:
    try:
        df = pd.read_excel(archivo_cargado)
        res = []
        
        for _, f in df.iterrows():
            mat = str(f["Material"]).strip()
            cal = int(f["Calibre"])
            if mat not in ESTANDAR or cal not in ESTANDAR[mat]:
                continue
            
            med = float(f["Espesor_Medido"])
            uni = str(f["Unidad"]).strip().lower()
            esp_in = round(med * 0.0393701, 4) if ("mm" in uni or "mili" in uni) else med
            
            res.append({
                "Rollo": str(f["Numero_Rollo"]),
                "Material": mat,
                "Calibre": f"CAL {cal}",
                "Espesor Original": f"{med} {'mm' if 'mm' in uni else 'in'}",
                "Espesor Real (in)": esp_in,
                "Nominal Estándar": ESTANDAR[mat][cal],
                "Desviación Real (in)": round(esp_in - ESTANDAR[mat][cal], 4)
            })
            
        df_datos_cargados = pd.DataFrame(res)
        
        if not df_datos_cargados.empty:
            # ======================================================================
            # MODIFICACIÓN 1: EVALUACIÓN DE RIESGO ESTADÍSTICO INDIVIDUAL (GAUSS)
            # ======================================================================
            est_l = []
            riesgo_l = []
            dictamen_final_l = []
            sigma_individual = tol_proveedor / 3.0
            
            for _, fila in df_datos_cargados.iterrows():
                med_individual = fila['Espesor Real (in)']
                nom_individual = fila['Nominal Estándar']
                
                # Integrales normales en colas críticas
                p_inf_ind = stats.norm.cdf(nom_individual - TOLERANCIA_INTERNA, loc=med_individual, scale=sigma_individual)
                p_sup_ind = 1.0 - stats.norm.cdf(nom_individual + TOLERANCIA_INTERNA, loc=med_individual, scale=sigma_individual)
                riesgo_rollo_pct = (p_inf_ind + p_sup_ind) * 100
                
                riesgo_l.append(riesgo_rollo_pct)
                
                # Clasificación matemática pura en columna "Riesgo" (Sustituye Estatus Planta)
                if riesgo_rollo_pct < 1.0:
                    est_l.append("RIESGO BAJO")
                    dictamen_final_l.append("ACEPTADO")
                elif riesgo_rollo_pct <= 5.0:
                    st_desc = "RIESGO MODERADO"
                    est_l.append(st_desc)
                    # Un riesgo moderado es aceptable condicionalmente bajo criterios standard
                    dictamen_final_l.append("ACEPTADO")
                else:
                    est_l.append("ALTO RIESGO")
                    dictamen_final_l.append("NO ACEPTADO")
            
            df_datos_cargados['% de Riesgo'] = riesgo_l
            df_datos_cargados['Riesgo'] = est_l # Columna renombrada
            df_datos_cargados['Dictamen Final'] = dictamen_final_l
            
            # --- DESPLIEGUE TABLA 1 (CON REQUERIMIENTO DE MODIFICACIÓN 1) ---
            st.subheader("📊 Calibración del Muestreo por Unidad (Rollo por Rollo)")
            formatos = {'Espesor Real (in)': '{:.4f}"', 'Nominal Estándar': '{:.3f}"', 'Desviación Real (in)': '{:+.4f}"', '% de Riesgo': '{:.2f}%'}
            styler_individual = df_datos_cargados.style.format(formatos).map(colorear_matriz_resumen, subset=['Riesgo'])
            st.dataframe(styler_individual, use_container_width=True)
            
            # ======================================================================
            # MODIFICACIÓN 2: ELIMINAR TABLA GLOBAL Y CLASIFICAR POR ESPESOR NOMINAL
            # ======================================================================
            st.subheader("📋 Análisis Clasificado Estructurado por Espesor Nominal Teórico")
            
            df_grouped = df_datos_cargados.groupby(['Material', 'Nominal Estándar'])
            
            for (material, nominal), grupo in df_grouped:
                # Encabezado por tipo de espesor de planta detectado en el muestreo
                st.markdown(f"#### 🌐 Material: `{material}` — Espesor Nominal de Diseño: `{nominal:.3f}\"`")
                
                # Despliegue de los rollos particulares correspondientes a esta categoría
                columnas_vista = ['Rollo', 'Calibre', 'Espesor Original', 'Espesor Real (in)', 'Desviación Real (in)', '% de Riesgo', 'Dictamen Final']
                df_vista_grupo = grupo[columnas_vista]
                
                styler_grupo = df_vista_grupo.style.format(formatos).map(colorear_matriz_resumen, subset=['Dictamen Final'])
                st.dataframe(styler_grupo, use_container_width=True)
            # ======================================================================
            # MODIFICACIÓN 3: UNA GRÁFICA INTERACTIVA POR CADA ESPESOR DETECTADO
            # ======================================================================
            st.subheader("📈 Distribuciones Probabilísticas por Especificación Técnica")
            
            # Iteración sobre los mismos grupos para aislar las campanas de Gauss de sus rollos
            for (material, nominal), grupo in df_grouped:
                st.markdown(f"##### Análisis Gaussiano Isolado: `{material} ({nominal:.3f}\")`")
                
                fig = go.Figure()
                x_desv = np.linspace(-0.015, 0.015, 400)
                
                # Trazado de líneas/curvas de los rollos que obedecen estrictamente a este espesor
                for _, fila in grupo.iterrows():
                    m_desv = fila['Espesor Real (in)'] - nominal
                    y_g = stats.norm.pdf(x_desv, loc=m_desv, scale=sigma_individual)
                    fig.add_trace(go.Scatter(
                        x=x_desv, y=y_g, mode='lines', 
                        name=f"{fila['Rollo']} (Riesgo: {fila['% de Riesgo']:.1f}%)",
                        line=dict(width=2.5)
                    ))
                
                # Región segura e indicadores de límites técnicos de la planta
                fig.add_vrect(x0=-TOLERANCIA_INTERNA, x1=TOLERANCIA_INTERNA, fillcolor="green", opacity=0.05, line_width=0)
                fig.add_vline(x=0, line_dash="dash", line_color="darkgreen")
                fig.add_vline(x=TOLERANCIA_INTERNA, line_color="red", line_width=1.5, line_dash="dot")
                fig.add_vline(x=-TOLERANCIA_INTERNA, line_color="red", line_width=1.5, line_dash="dot")
                
                fig.update_layout(
                    xaxis_title="Desviación Micrométrica Real (in)", 
                    yaxis_title="Densidad Probabilística de Gauss", 
                    height=320, 
                    margin=dict(l=40, r=40, t=15, b=40),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig, use_container_width=True)
            
            # --- EXPORTACIÓN DE REPORTE PDF INDUSTRIAL ACTUALIZADO ---
            st.subheader("📄 Entregables de Ingeniería de Calidad")
            pdf_buffer = crear_pdf_formal(df_datos_cargados, tol_proveedor)
            st.download_button(
                label="📥 Descargar Reporte PDF de Control Corporativo",
                data=pdf_buffer,
                file_name=f"Reporte_Riesgo_Espesores_{datetime.now().strftime('%H%M')}.pdf",
                mime="application/pdf"
            )
            
    except Exception as e:
        st.error(f"❌ Error crítico en el procesamiento del lote técnico: {str(e)}")
else:
    st.info("💡 Tablero listo. Por favor, cargue un archivo de Excel utilizando la plantilla estándar para iniciar las simulaciones estadísticas.")
