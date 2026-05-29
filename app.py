import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import scipy.stats as stats
from datetime import datetime
import io

# Componentes oficiales de ReportLab para la maquetación formal del informe
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Configuración obligatoria para el renderizado de gráficos vectoriales sin servidor en la nube
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# Configuración del entorno de visualización en la pestaña web
st.set_page_config(page_title="Suite de Riesgo y Control", layout="wide")

# Matriz técnica basada estrictamente en los espesores nominales de la planta
ESTANDAR = {
    "Galvanizado": {10: 0.138, 12: 0.108, 14: 0.079, 16: 0.064},
    "Decapado": {12: 0.105, 14: 0.075, 16: 0.060}
}
TOLERANCIA_INTERNA = 0.008
def calcular_riesgo_grupos(df_entrada, tol_proveedor):
    """Calcula el área bajo la curva normal para hallar el porcentaje exacto de riesgo de rechazo."""
    sigma_p = tol_proveedor / 3.0
    resumen = []
    # Agrupamos sobre una copia profunda para evitar efectos colaterales en el dataframe principal
    for (m, c), sub in df_entrada.groupby(['Material', 'Calibre']):
        cal_i = int(c.replace("CAL ", ""))
        nom = ESTANDAR[m][cal_i]
        med = float(sub['Espesor Real (in)'].mean())
        
        p_inf = stats.norm.cdf(nom - TOLERANCIA_INTERNA, loc=med, scale=sigma_p)
        p_sup = 1.0 - stats.norm.cdf(nom + TOLERANCIA_INTERNA, loc=med, scale=sigma_p)
        r_pct = (p_inf + p_sup) * 100
        
        dictam = "RIESGO BAJO" if r_pct < 1.0 else ("MODERADO" if r_pct <= 5.0 else "ALTO RIESGO")
        resumen.append({"Clave": f"{m} {c}", "Media": med, "Riesgo": r_pct, "Dictamen": dictam, "Nominal": nom})
    return resumen

def colorear_matriz_resumen(v):
    """Aplica formato semafórico basado en las cadenas de texto de riesgo."""
    if not isinstance(v, str): return ''
    if "BAJO" in v: return 'background-color: #C6EFCE; color: #006100; font-weight: bold;'
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
def crear_pdf_formal(df_final, resumen, tol_p):
    """Genera la estructura del documento técnico formal usando un buffer en memoria."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    styles = getSampleStyleSheet()
    
    t_st = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#2b579a'), spaceAfter=12)
    m_st = ParagraphStyle('DocMeta', fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#444'), spaceAfter=4)
    h2_st = ParagraphStyle('SectionH2', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#2b579a'), spaceBefore=12, spaceAfter=6)
    h_style = ParagraphStyle('HStyle', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1)
    c_style = ParagraphStyle('CStyle', fontName='Helvetica', fontSize=9, alignment=1)
    c_bold = ParagraphStyle('CBold', fontName='Helvetica-Bold', fontSize=9, alignment=1)
    s_bajo = ParagraphStyle('S1', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#006100'), alignment=1)
    s_mod = ParagraphStyle('S2', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#7F6000'), alignment=1)
    s_alto = ParagraphStyle('S3', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#9C0006'), alignment=1)
    
    story.append(Paragraph("SIGRAMA PLANTA METALES", t_st))
    story.append(Paragraph("<b>Documento:</b> Reporte Técnico de Ingeniería de Calidad y Evaluación de Suministro", m_st))
    story.append(Paragraph(f"<b>Fecha de Análisis:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", m_st))
    story.append(Paragraph(f"<b>Parámetro Comercial:</b> Desviación Ofertada por Proveedor en ±{tol_p:.3f}\"", m_st))
    story.append(Spacer(1, 10))
    
    # Tabla Sección 1: Criterios
    story.append(Paragraph("1. Especificaciones Técnicas y Criterios de Aceptación Interna", h2_st))
    t_req_d = [[Paragraph("Material / Calibre", h_style), Paragraph("Espesor Requerido<br/>(Nominal)", h_style), Paragraph("Límite Mínimo<br/>(-0.008\")", h_style), Paragraph("Límite Máximo<br/>(+0.008\")", h_style)]]
    for r in resumen:
        t_req_d.append([Paragraph(r['Clave'], c_bold), Paragraph(f"{r['Nominal']:.3f}\"", c_style), Paragraph(f"{r['Nominal']-TOLERANCIA_INTERNA:.3f}\"", c_style), Paragraph(f"{r['Nominal']+TOLERANCIA_INTERNA:.3f}\"", c_style)])
    t_1 = Table(t_req_d, colWidths=[130, 110, 110, 110])
    t_1.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#708090')), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#A0A0A0')), ('PADDING', (0,0), (-1,-1), 5)]))
    story.append(t_1)
    
    # Tabla Sección 2: Diagnóstico
    story.append(Paragraph("2. Diagnóstico Probabilístico y Matriz de Riesgo Colectiva", h2_st))
    t_riesgo_d = [[Paragraph("Material / Calibre", h_style), Paragraph("Media Lote Real", h_style), Paragraph("Prob. Fuera de Norma", h_style), Paragraph("Dictamen Técnico", h_style)]]
    est_riesgo = []
    for idx, r in enumerate(resumen):
        p_d, bg = (Paragraph(r['Dictamen'], s_bajo), colors.HexColor('#C6EFCE')) if "BAJO" in r['Dictamen'] else ((Paragraph(r['Dictamen'], s_mod), colors.HexColor('#FFF2CC')) if "MODERADO" in r['Dictamen'] else (Paragraph(r['Dictamen'], s_alto), colors.HexColor('#FFC7CE')))
        t_riesgo_d.append([Paragraph(r['Clave'], c_bold), Paragraph(f"{r['Media']:.4f}\"", c_style), Paragraph(f"{r['Riesgo']:.2f}%", c_style), p_d])
        est_riesgo.append(('BACKGROUND', (3, idx+1), (3, idx+1), bg))
    t_2 = Table(t_riesgo_d, colWidths=[130, 110, 110, 110])
    t_2.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2b579a')), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#A0A0A0')), ('PADDING', (0,0), (-1,-1), 5)] + est_riesgo))
    story.append(t_2)
    
    # Tabla Sección 3: Muestreo Físico por Unidades
    story.append(Paragraph("3. Desglose Analítico por Unidad (Muestreo Físico)", h2_st))
    t_rollos_d = [[Paragraph("Rollo", h_style), Paragraph("Material", h_style), Paragraph("Calibre", h_style), Paragraph("Espesor (in)", h_style), Paragraph("Desviación vs Est.", h_style), Paragraph("Estatus Planta", h_style)]]
    est_rollos = []
    for idx, f in df_final.reset_index(drop=True).iterrows():
        p_e, bg = (Paragraph(f['Estatus Planta'], s_bajo), colors.HexColor('#C6EFCE')) if "BAJO" in f['Estatus Planta'] else ((Paragraph(f['Estatus Planta'], s_mod), colors.HexColor('#FFF2CC')) if "MODERADO" in f['Estatus Planta'] else (Paragraph(f['Estatus Planta'], s_alto), colors.HexColor('#FFC7CE')))
        t_rollos_d.append([Paragraph(str(f['Rollo']), c_style), Paragraph(f['Material'], c_style), Paragraph(f['Calibre'], c_style), Paragraph(f"{f['Espesor Real (in)']:.4f}\"", c_style), Paragraph(f"{f['Desviación Real (in)']:+4f}\"", c_style), p_e])
        est_rollos.append(('BACKGROUND', (5, idx+1), (5, idx+1), bg))
    t_3 = Table(t_rollos_d, colWidths=[90, 80, 60, 75, 85, 90])
    t_3.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2b579a')), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D3D3D3')), ('PADDING', (0,0), (-1,-1), 4)] + est_rollos))
    story.append(t_3)
    
    story.append(PageBreak())
    story.append(Paragraph("4. Distribution Estadística de Calidad (Campanas de Gauss)", h2_st))
    
    try:
        plt.figure(figsize=(7, 3.5))
        x_desv = np.linspace(-0.015, 0.015, 400)
        sigma_p = tol_p / 3.0
        
        for idx, r in enumerate(resumen):
            m_desv = r['Media'] - r['Nominal']
            y_g = stats.norm.pdf(x_desv, loc=m_desv, scale=sigma_p)
            plt.plot(x_desv, y_g, label=f"{r['Clave']} ({r['Riesgo']:.1f}%)", linewidth=2)
            
        plt.axvspan(-TOLERANCIA_INTERNA, TOLERANCIA_INTERNA, color='green', alpha=0.06)
        plt.axvline(0, color='darkgreen', linestyle='--')
        plt.axvline(TOLERANCIA_INTERNA, color='red', linestyle=':')
        plt.axvline(-TOLERANCIA_INTERNA, color='red', linestyle=':')
        plt.title("Curvas de Distribución de Gauss", fontsize=11, color='#2b579a', weight='bold')
        plt.legend(loc="upper right", fontsize=8)
        plt.grid(True, linestyle=':', alpha=0.5)
        plt.tight_layout()
        
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', dpi=180)
        img_buffer.seek(0)
        plt.close()
        
        img_pdf = RLImage(img_buffer, width=450, height=225)
        img_pdf.hAlign = 'CENTER'
        story.append(img_pdf)
    except Exception as e:
        story.append(Paragraph(f"<i>No se pudo renderizar la gráfica técnica en esta hoja: {str(e)}</i>", c_style))
        
    story.append(Spacer(1, 20))
    f_st = ParagraphStyle('FText', fontName='Helvetica', fontSize=10, alignment=1, spaceAfter=2)
    story.append(Paragraph("___________________________________________________", f_st))
    story.append(Paragraph("<b>Ing. Jesús Morales</b>", f_st))
    story.append(Paragraph("Dir. Planta Metales | SIGRAMA PLANTA METALES", f_st))
    
    doc.build(story)
    buffer.seek(0)
    return buffer
st.title("⚙️ Suite Interactiva de Riesgo y Control de Suministros")
st.markdown(f"**Estándar Fijo Planta (Norma Interna):** `±{TOLERANCIA_INTERNA}\"`")

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
            
        # Generamos un DataFrame base y limpio
        df_datos_cargados = pd.DataFrame(res)
        
        if not df_datos_cargados.empty:
            # --- EVALUACIÓN MATEMÁTICA PURA LÍNEA POR LÍNEA ---
            est_l = []
            for _, fila in df_datos_cargados.iterrows():
                desv_abs = abs(fila['Desviación Real (in)'])
                
                # Clasificación micrométrica inflexible sin promediar
                if desv_abs <= 0.0040:
                    est_l.append("RIESGO BAJO")
                elif desv_abs <= TOLERANCIA_INTERNA: # Rango exacto de 0.0041" a 0.0080"
                    est_l.append("RIESGO MODERADO")
                else: # Excede 0.0080"
                    est_l.append("ALTO RIESGO")
            
            # Asignamos la lista evaluada a la columna final
            df_datos_cargados['Estatus Planta'] = est_l
            
            # --- RENDERIZADO DE LA TABLA 1: MUESTREO INDIVIDUAL ---
            st.subheader("📊 Calibración del Muestreo por Unidad (Rollo por Rollo)")
            formatos = {
                'Espesor Real (in)': '{:.4f}"', 
                'Nominal Estándar': '{:.3f}"', 
                'Desviación Real (in)': '{:+.4f}"'
            }
            
            # Renderizado directo e independiente en la pantalla
            styler_individual = df_datos_cargados.style.format(formatos).map(
                colorear_matriz_resumen, 
                subset=['Estatus Planta']
            )
            st.dataframe(styler_individual, use_container_width=True)
            
            # --- RENDERIZADO DE LA TABLA 2: MATRIZ DE RIESGO DE COMPRA ---
            # Pasamos una copia profunda limpia para que la agregación estadística no contamine la vista individual
            df_para_resumen = df_datos_cargados.copy()
            resumen = calcular_riesgo_grupos(df_para_resumen, tol_proveedor)
            
            st.subheader("📋 Matriz Estratégica de Riesgo Técnico del Lote (Análisis de Gauss)")
            df_resumen = pd.DataFrame(resumen)
            styler_resumen = df_resumen.style.map(
                colorear_matriz_resumen, 
                subset=['Dictamen']
            )
            st.dataframe(styler_resumen, use_container_width=True)
            
            # --- 3. GRÁFICA WEB INTERACTIVA CON PLOTLY ---
            st.subheader("📈 Distribución Probabilística (Campanas de Gauss)")
            fig = go.Figure()
            x_desv = np.linspace(-0.015, 0.015, 400)
            colores = ['#1f77b4', '#ff7f0e', '#2ca02c', '#9467bd']
            sig_t = tol_proveedor / 3.0
            
            for idx, r in enumerate(resumen):
                m_desv = r['Media'] - r['Nominal']
                y_g = stats.norm.pdf(x_desv, loc=m_desv, scale=sig_t)
                fig.add_trace(go.Scatter(x=x_desv, y=y_g, mode='lines', name=f"{r['Clave']} ({r['Riesgo']:.1f}%)", line=dict(color=colores[idx % len(colores)], width=3)))
            
            fig.add_vrect(x0=-TOLERANCIA_INTERNA, x1=TOLERANCIA_INTERNA, fillcolor="green", opacity=0.06, line_width=0)
            fig.add_vline(x=0, line_dash="dash", line_color="darkgreen")
            fig.add_vline(x=TOLERANCIA_INTERNA, line_color="red", line_width=1.5)
            fig.add_vline(x=-TOLERANCIA_INTERNA, line_color="red", line_width=1.5)
            fig.update_layout(xaxis_title="Desviación (in)", yaxis_title="Densidad", height=380, margin=dict(l=40, r=40, t=10, b=40))
            st.plotly_chart(fig, use_container_width=True)
            
            # --- 4. EXPORTACIÓN A PDF ---
            st.subheader("📄 Entregables de Ingeniería de Calidad")
            pdf_buffer = crear_pdf_formal(df_datos_cargados, resumen, tol_proveedor)
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
