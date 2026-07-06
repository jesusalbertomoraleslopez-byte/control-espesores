import sqlite3
import os
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

# Inicializar al importar para asegurar que la tabla existe
init_db()
