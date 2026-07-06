import sqlite3
import os
import shutil
from datetime import datetime

# Definir la ruta de la base de datos en el mismo directorio del proyecto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "espesores_historial.db")
EXPEDIENTES_DIR = os.path.join(BASE_DIR, "expedientes")

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Inicializa la base de datos y crea la tabla de historial y de proveedores si no existen."""
    os.makedirs(EXPEDIENTES_DIR, exist_ok=True)
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS historial_reportes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        folio TEXT UNIQUE NOT NULL,
        fecha TEXT NOT NULL,
        proveedor TEXT NOT NULL,
        certificado_info TEXT,
        ruta_certificado TEXT,
        ruta_correo TEXT,
        ruta_reporte TEXT,
        desviacion_ofertada_def REAL,
        total_rollos INTEGER,
        aceptados INTEGER,
        rechazados INTEGER,
        riesgo_promedio REAL,
        fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS proveedores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT UNIQUE NOT NULL,
        contacto TEXT,
        telefono TEXT,
        correo TEXT,
        fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    
    # Insertar proveedores por defecto si la tabla está vacía
    cursor.execute("SELECT COUNT(*) as cnt FROM proveedores")
    if cursor.fetchone()["cnt"] == 0:
        cursor.execute("INSERT INTO proveedores (nombre, contacto, correo) VALUES ('Ternium', 'Jesús Morales', 'jmorales@ternium.com')")
        cursor.execute("INSERT INTO proveedores (nombre, contacto, correo) VALUES ('Nucor', 'Brenda Martínez', 'brenda.martinez@nucor.com')")
        cursor.execute("INSERT INTO proveedores (nombre, contacto, correo) VALUES ('AHMSA', 'Carlos Sánchez', 'csanchez@ahmsa.com')")
        
    conn.commit()
    conn.close()

def guardar_reporte(folio, fecha, proveedor, certificado_info, ruta_certificado, ruta_correo, ruta_reporte, desviacion_ofertada_def, total_rollos, aceptados, rechazados, riesgo_promedio):
    """Guarda un registro de reporte en la base de datos."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    INSERT INTO historial_reportes (
        folio, fecha, proveedor, certificado_info, ruta_certificado, ruta_correo, ruta_reporte, desviacion_ofertada_def, total_rollos, aceptados, rechazados, riesgo_promedio
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (folio, fecha, proveedor, certificado_info, ruta_certificado, ruta_correo, ruta_reporte, desviacion_ofertada_def, total_rollos, aceptados, rechazados, riesgo_promedio))
    conn.commit()
    conn.close()
    push_to_github()

def obtener_reportes(fecha_inicio=None, fecha_fin=None, proveedor=None):
    """Obtiene los reportes filtrados por rango de fechas y/o proveedor."""
    conn = get_connection()
    cursor = conn.cursor()
    
    query = "SELECT * FROM historial_reportes WHERE 1=1"
    params = []
    
    if fecha_inicio:
        query += " AND date(substr(fecha, 7, 4) || '-' || substr(fecha, 4, 2) || '-' || substr(fecha, 1, 2)) >= date(?)"
        params.append(fecha_inicio)
        
    if fecha_fin:
        query += " AND date(substr(fecha, 7, 4) || '-' || substr(fecha, 4, 2) || '-' || substr(fecha, 1, 2)) <= date(?)"
        params.append(fecha_fin)
        
    if proveedor and proveedor != "Todos":
        query += " AND proveedor = ?"
        params.append(proveedor)
        
    query += " ORDER BY folio DESC"
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def obtener_proveedores():
    """Obtiene la lista de proveedores únicos registrados para los filtros."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT proveedor FROM historial_reportes ORDER BY proveedor")
    rows = cursor.fetchall()
    conn.close()
    return [r["proveedor"] for r in rows]

def generar_siguiente_folio():
    """Genera el siguiente folio secuencial basado en el año actual (ej. REP-ESP-2026-0001)."""
    conn = get_connection()
    cursor = conn.cursor()
    year = datetime.now().year
    
    # Buscar el último folio del año actual
    cursor.execute("SELECT folio FROM historial_reportes WHERE folio LIKE ? ORDER BY id DESC LIMIT 1", (f"REP-ESP-{year}-%",))
    row = cursor.fetchone()
    conn.close()
    
    if row:
        ultimo_folio = row["folio"]
        try:
            secuencia = int(ultimo_folio.split("-")[-1])
            siguiente = secuencia + 1
        except ValueError:
            siguiente = 1
    else:
        siguiente = 1
        
    return f"REP-ESP-{year}-{siguiente:04d}"

def crear_proveedor(nombre, contacto="", telefono="", correo=""):
    """Crea un nuevo proveedor en la base de datos."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
        INSERT INTO proveedores (nombre, contacto, telefono, correo)
        VALUES (?, ?, ?, ?)
        """, (nombre, contacto, telefono, correo))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = False
    conn.close()
    if success:
        push_to_github()
    return success

def listar_proveedores():
    """Retorna una lista de diccionarios con todos los proveedores registrados."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM proveedores ORDER BY nombre")
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def listar_proveedores_nombres():
    """Retorna una lista de nombres de todos los proveedores registrados."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT nombre FROM proveedores ORDER BY nombre")
    rows = cursor.fetchall()
    conn.close()
    return [r["nombre"] for r in rows]

def eliminar_proveedor(id_prov):
    """Elimina un proveedor por su ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM proveedores WHERE id = ?", (id_prov,))
    conn.commit()
    conn.close()
    push_to_github()

def eliminar_reporte(folio):
    """Elimina un reporte específico y todos sus archivos físicos asociados."""
    # 1. Eliminar archivos físicos
    folder_exp = os.path.join(EXPEDIENTES_DIR, folio)
    if os.path.exists(folder_exp):
        shutil.rmtree(folder_exp)
        
    # 2. Eliminar registro de la base de datos
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historial_reportes WHERE folio = ?", (folio,))
    conn.commit()
    conn.close()
    push_to_github()

def limpiar_base_datos():
    """Borra todos los registros, proveedores y archivos, reiniciando el sistema a su estado de fábrica."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historial_reportes")
    cursor.execute("DELETE FROM proveedores")
    
    # Reiniciar los contadores de IDs automáticos (para que vuelvan a empezar en 1)
    cursor.execute("DELETE FROM sqlite_sequence WHERE name='historial_reportes'")
    cursor.execute("DELETE FROM sqlite_sequence WHERE name='proveedores'")
    
    conn.commit()
    conn.close()
    
    # Reiniciar la estructura de expedientes (borrar todo físicamente)
    if os.path.exists(EXPEDIENTES_DIR):
        shutil.rmtree(EXPEDIENTES_DIR)
    
    # Recrear estructura y proveedores por defecto
    init_db()
    push_to_github()

def push_to_github():
    """Sincroniza la base de datos y los archivos de expedientes con el repositorio de GitHub."""
    import subprocess
    import streamlit as st
    try:
        if "GITHUB_TOKEN" not in st.secrets:
            return False
            
        token = st.secrets["GITHUB_TOKEN"]
        
        # Configurar identidad del commit temporal
        subprocess.run(["git", "config", "user.name", "SIGRAMA Auto-Sincronizador"], capture_output=True)
        subprocess.run(["git", "config", "user.email", "calidad@sigrama.com.mx"], capture_output=True)
        
        # Agregar base de datos y expedientes
        subprocess.run(["git", "add", "espesores_historial.db", "expedientes/"], capture_output=True)
        
        # Hacer commit
        res_commit = subprocess.run(["git", "commit", "-m", "Sincronizacion automatica de base de datos y expedientes [bot]"], capture_output=True, text=True)
        if "nothing to commit" in res_commit.stdout or "nothing added to commit" in res_commit.stdout:
            return True
            
        # Hacer push usando el token de acceso configurado
        repo_url = f"https://x-access-token:{token}@github.com/jesusalbertomoraleslopez-byte/control-espesores.git"
        res_push = subprocess.run(["git", "push", repo_url, "HEAD:main"], capture_output=True, text=True)
        
        return res_push.returncode == 0
    except Exception as e:
        print(f"Error al sincronizar con GitHub: {e}")
        return False

def enviar_correo_smtp(recipient, cc_recipients, subject, body, attachment_paths=None):
    """Envía un correo electrónico con múltiples archivos adjuntos usando credenciales SMTP en st.secrets."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    import streamlit as st
    import os
    
    try:
        smtp_server = st.secrets["SMTP_SERVER"]
        smtp_port = int(st.secrets["SMTP_PORT"])
        smtp_user = st.secrets["SMTP_USER"]
        smtp_password = st.secrets["SMTP_PASSWORD"]
        
        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = recipient
        if cc_recipients:
            msg['Cc'] = cc_recipients
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain'))
        
        if attachment_paths:
            for path in attachment_paths:
                if path and os.path.exists(path):
                    filename = os.path.basename(path)
                    with open(path, "rb") as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f"attachment; filename= {filename}")
                        msg.attach(part)
                        
        # Crear lista plana de todos los destinatarios (To + Cc) para el envío SMTP
        destinations = [x.strip() for x in recipient.replace(",", ";").split(";") if x.strip()]
        if cc_recipients:
            destinations += [x.strip() for x in cc_recipients.replace(",", ";").split(";") if x.strip()]
                
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, destinations, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Error al enviar correo SMTP: {e}")
        return False

# Inicializar al importar para asegurar que la tabla existe
init_db()
