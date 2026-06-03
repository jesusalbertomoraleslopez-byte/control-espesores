import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import scipy.stats as stats
from datetime import datetime
import io

# 1. CONFIGURACIÓN ÚNICA DE LA PÁGINA (Debe ser la primera instrucción de Streamlit)
st.set_page_config(page_title="Suite de Riesgo y Control de Espesores", layout="wide")

# 2. CONFIGURACIÓN GRÁFICA VECTORIAL Y REPORTES NATIVOS
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

import matplotlib
matplotlib.use('Agg') # Renderizado seguro en backend para servidores sin entorno gráfico
import matplotlib.pyplot as plt

# 3. MATRIZ TÉCNICA PLANTA (Límites rígidos de diseño nominal)
ESTANDAR = {
    "Galvanizado": {10: 0.138, 12: 0.108, 14: 0.079, 16: 0.064},
    "Decapado": {12: 0.105, 14: 0.075, 16: 0.060}
}
TOLERANCIA_INTERNA = 0.008

# Despliegue de banner corporativo principal
st.image("BANNER CONTROL DE ESPESORES APP.png", use_container_width=True)
def colorear_matriz_resumen(v):
    """Aplica formato semafórico condicional estricto: Verde para aceptados y Rojo para riesgosos/rechazados."""
    if not isinstance(v, str): return ''
    if "BAJO" in v or "ACEPTADO" == v: 
        return 'background-color: #C6EFCE; color: #006100; font-weight: bold;'
    if "MODERADO" in v: 
        return 'background-color: #FFF2CC; color: #7F6000; font-weight: bold;'
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
    """Genera la estructura del documento técnico formal incorporando los nuevos formatos de color y títulos."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    styles = getSampleStyleSheet()
    
    t_st = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#2b579a'), spaceAfter=12)
    m_st = ParagraphStyle('DocMeta', fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#444'), spaceAfter=4)
    h2_st = ParagraphStyle('SectionH2', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#2b579a'), spaceBefore=12, spaceAfter=6)
    h3_st = ParagraphStyle('SectionH3', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#222222'), spaceBefore=8, spaceAfter=4)
    h_style = ParagraphStyle('HStyle', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1)
    c_style = ParagraphStyle('CStyle', fontName='Helvetica', fontSize=9, alignment=1)
    
    story.append(Paragraph("SIGRAMA PLANTA METALES", t_st))
    story.append(Paragraph("<b>Documento:</b> Reporte Técnico de Ingeniería de Calidad y Evaluación de Suministro", m_st))
    story.append(Paragraph(f"<b>Fecha de Análisis:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", m_st))
    story.append(Paragraph(f"<b>Parámetro Comercial:</b> Desviación Ofertada por Proveedor en ±{tol_p:.3f}\"", m_st))
    story.append(Spacer(1, 10))
    
    # 1. Tabla de calibración general
    story.append(Paragraph("1. Calibración del Muestreo por Unidad (Rollo por Rollo)", h2_st))
    t_rollos_d = [[Paragraph("Rollo", h_style), Paragraph("Material", h_style), Paragraph("Calibre", h_style), Paragraph("Espesor (in)", h_style), Paragraph("Riesgo %", h_style), Paragraph("Riesgo", h_style)]]
    for idx, f in df_final.reset_index(drop=True).iterrows():
        t_rollos_d.append([
            Paragraph(str(f['Rollo']), c_style), Paragraph(f['Material'], c_style), Paragraph(str(f['Calibre']), c_style),
            Paragraph(f"{f['Espesor Real (in)']:.4f}\"", c_style), Paragraph(f"{f['% de Riesgo']:.2f}%", c_style), Paragraph(f['Riesgo'], c_style)
        ])
    t_1 = Table(t_rollos_d, colWidths=[90, 100, 70, 90, 80, 100])
    t_1.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2b579a')), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3'))]))
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
        est_estilo_grupo = [('BACKGROUND', (0,0), (-1,0), colors.HexColor('#708090')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#A0A0A0'))]
        
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
    story.append(PageBreak())
    story.append(Paragraph("3. Análisis de Distribución Probabilística por Especificación Técnica", h2_st))
    
    sigma_p = tol_p / 3.0
    for (mat, calibre, nominal), group in df_grouped:
        try:
            plt.figure(figsize=(6.5, 2.8))
            x_desv = np.linspace(-0.015, 0.015, 400)
            
            for _, fila in group.iterrows():
                m_desv = fila['Espesor Real (in)'] - nominal
                y_g = stats.norm.pdf(x_desv, loc=m_desv, scale=sigma_p)
                plt.plot(x_desv, y_g, label=f"{fila['Rollo']} ({fila['% de Riesgo']:.1f}%)", linewidth=1.5)
                
            plt.axvspan(-TOLERANCIA_INTERNA, TOLERANCIA_INTERNA, color='green', alpha=0.04)
            plt.axvline(0, color='darkgreen', linestyle='--')
            plt.axvline(TOLERANCIA_INTERNA, color='red', linestyle=':')
            plt.axvline(-TOLERANCIA_INTERNA, color='red', linestyle=':')
            plt.title(f"{mat} - {calibre} - Nominal: {nominal:.3f}\"", fontsize=10, color='#2b579a', weight='bold')
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
    f_st = ParagraphStyle('FText', fontName='Helvetica', fontSize=10, alignment=1, spaceAfter=2)
    story.append(Paragraph("___________________________________________________", f_st))
    story.append(Paragraph("<b>Ing. Jesús Morales</b>", f_st))
    story.append(Paragraph("Dir. Planta Metales | SIGRAMA PLANTA METALES", f_st))
    
    doc.build(story)
    buffer.seek(0)
    return buffer
st.title("⚙️ Suite Interactive de Riesgo y Control de Suministros")
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
            # MODIFICACIÓN 1: EVALUACIÓN DE RIESGO ESTADÍSTICO POR ROLLO (GAUSS)
            # ======================================================================
            est_l = []
            riesgo_l = []
            dictamen_final_l = []
            sigma_individual = tol_proveedor / 3.0
            
            for _, fila in df_datos_cargados.iterrows():
                med_individual = fila['Espesor Real (in)']
                nom_individual = fila['Nominal Estándar']
                
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
            df_datos_cargados['Riesgo'] = est_l # Estatus Planta cambia de título a Riesgo
            df_datos_cargados['Dictamen Final'] = dictamen_final_l
            
            # RENDERIZADO TABLA 1 (MODIFICACIÓN 1 EXPOSITIVA)
            st.subheader("📊 Calibración del Muestreo por Unidad (Rollo por Rollo)")
            formatos = {'Espesor Real (in)': '{:.4f}"', 'Nominal Estándar': '{:.3f}"', 'Desviación Real (in)': '{:+.4f}"', '% de Riesgo': '{:.2f}%'}
            styler_individual = df_datos_cargados.style.format(formatos).map(colorear_matriz_resumen, subset=['Riesgo'])
            st.dataframe(styler_individual, use_container_width=True)
            
            # ======================================================================
            # MODIFICACIÓN 2: CLASIFICACIÓN JERÁRQUICA Y SE REMOVE "ESPESOR ORIGINAL"
            # ======================================================================
            st.subheader("📋 Análisis Clasificado Estructurado por Espesor Nominal Teórico")
            df_grouped = df_datos_cargados.groupby(['Material', 'Calibre', 'Nominal Estándar'])
            
            for (material, calibre, nominal), grupo in df_grouped:
                # Encabezado corregido con especificaciones y tolerancia aceptable
                st.markdown(f"#### 🌐 {material} — `{calibre}` — Espesor Teórico: `{nominal:.3f}\"` | Tolerancia Aceptable: `±{TOLERANCIA_INTERNA:.3f}\"`")
                
                # SE EXCLUYE COMPLETAMENTE 'Espesor Original' para la vista solicitada
                columnas_vista = ['Rollo', 'Espesor Real (in)', 'Desviación Real (in)', '% de Riesgo', 'Dictamen Final']
                df_vista_grupo = grupo[columnas_vista]
                
                # Renderizado con colores condicionales (Verde = ACEPTADO, Rojo = NO ACEPTADO)
                styler_grupo = df_vista_grupo.style.format(formatos).map(colorear_matriz_resumen, subset=['Dictamen Final'])
                st.dataframe(styler_grupo, use_container_width=True)
            # ======================================================================
            # MODIFICACIÓN 3: GENERACIÓN DE UNA GRÁFICA AISLADA POR CADA ESPESOR
            # ======================================================================
            st.subheader("📈 Distribuciones Probabilísticas por Especificación Técnica")
            
            for (material, calibre, nominal), grupo in df_grouped:
                st.markdown(f"##### Análisis Gaussiano Aislado: `{material} — {calibre} ({nominal:.3f}\")`")
                
                fig = go.Figure()
                x_desv = np.linspace(-0.015, 0.015, 400)
                
                # Inyección exclusiva de curvas correspondientes a este espesor particular
                for _, fila in grupo.iterrows():
                    m_desv = fila['Espesor Real (in)'] - nominal
                    y_g = stats.norm.pdf(x_desv, loc=m_desv, scale=sigma_individual)
                    fig.add_trace(go.Scatter(
                        x=x_desv, y=y_g, mode='lines', 
                        name=f"{fila['Rollo']} (Riesgo: {fila['% de Riesgo']:.1f}%)",
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
