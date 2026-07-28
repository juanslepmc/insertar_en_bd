import mysql.connector
from mysql.connector import Error
import pandas as pd
import os

# --------------------------------------------------
# CONFIGURACIÓN DE CONEXIÓN INSTITUCIONAL
# --------------------------------------------------
def crear_conexion():
    return mysql.connector.connect(
        host='localhost',
        port=3306,
        user='root',
        password='123456789',
        database='bd_institucional_v2'
    )

# --------------------------------------------------
# FUNCIONES AUXILIARES DE ACONDICIONAMIENTO DE TIPOS
# --------------------------------------------------
def limpiar_decimal(valor):
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "sin información":
        return 0.0
    try:
        return float(valor)
    except:
        return 0.0

def limpiar_entero(valor):
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "sin información":
        return 0
    try:
        return int(float(valor))
    except:
        return 0

# --------------------------------------------------
# ETL COMPLETO: VISACIÓN COMPRAS P01
# --------------------------------------------------
def cargar_visacion_compras_p01(conexion, ruta_archivo):
    cursor = conexion.cursor()

    try:
        # Carga del set de datos original
        df = pd.read_excel(ruta_archivo)

        # Estandarización de cabeceras en minúsculas y sin espacios
        df.columns = df.columns.str.strip().str.lower()
        
        # Tratamiento preventivo general para valores nulos (NaN)
        df = df.fillna("Sin información")
        
        # Inicio de bloque transaccional
        conexion.start_transaction()
        
        # Limpieza previa de la tabla antes de inyectar el nuevo procesamiento
        cursor.execute("DELETE FROM Visacion_compras_P01")

        # Sentencia SQL estructurada exactamente en el orden de tu lista (48 atributos)
        sql = """
            INSERT INTO Visacion_compras_P01 (
                id_solicitud_compra, anio_solicitud_compra, folio_solicitud_compra,
                fecha_ingreso, fecha_lectura, fecha_revision, codigo_tarea,
                nombre_tarea, tipo_visacion, indice_en_ciclo, tiempo_lectura,
                tiempo_revision, tarea_valida, numero_ciclo, codigo_division,
                nombre_division, codigo_unidad, nombre_unidad, rol_usuario_asignado,
                login_usuario_asignado, tipo_usuario_asignado, nombre_usuario_asignado, login_usuario_finaliza_tarea,
                tipo_usuario_finaliza_tarea, nombre_usuario_finaliza_tarea, visacion_en_consultante, login_en_consultante,
                nombre_usuario_consultante, nombre_consulta, texto_consulta, tipo_consulta,
                tipo_respuesta_consulta, consulta_a_usuarios_especificos, login_usuario_consultado, tipo_usuario_consultado,
                nombre_usuario_consultado, rol_consultado, codigo_unidad_consulta, nombre_unidad_consulta,
                login_usuario_consultado_respuesta, tipo_usuario_consultado_respuesta, nombre_usuario_consultado_respuesta, fecha_ingreso_consulta,
                fecha_lectura_consulta, fecha_visacion_consulta, tiempo_lectura_consulta, tiempo_revision_consulta
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s
            )
        """
        
        lista_lote = []

        for index, row in df.iterrows():
            # Mapeo posicional indexado 1 a 1 resguardando la coherencia de datos
            valores = (
                str(row.get('id_solicitud_compra', '')),
                limpiar_entero(row.get('ano_solicitud_compra', 0)), # Match con 'ano_' del Excel
                str(row.get('folio_solicitud_compra', '')),
                str(row.get('fecha_ingreso', '')),
                str(row.get('fecha_lectura', '')),
                str(row.get('fecha_revision', '')),
                str(row.get('codigo_tarea', '')),
                str(row.get('nombre_tarea', '')),
                str(row.get('tipo_visacion', '')),
                limpiar_entero(row.get('indice_en_ciclo', 0)),
                limpiar_decimal(row.get('tiempo_lectura', 0.0)),
                limpiar_decimal(row.get('tiempo_revision', 0.0)),
                str(row.get('tarea_valida', '')),
                limpiar_entero(row.get('numero_ciclo', 0)),
                str(row.get('codigo_division', '')),
                str(row.get('nombre_division', '')),
                str(row.get('codigo_unidad', '')),
                str(row.get('nombre_unidad', '')),
                str(row.get('rol_usuario_asignado', '')),
                str(row.get('login_usuario_asignado', '')),
                str(row.get('tipo_usuario_asignado', '')),
                str(row.get('nombre_usuario_asignado', '')),
                str(row.get('login_usuario_finaliza_tarea', '')),
                str(row.get('tipo_usuario_finaliza_tarea', '')),
                str(row.get('nombre_usuario_finaliza_tarea', '')),
                str(row.get('visacion_en_consultas', '')),          # Match con 'visacion_en_consultas'
                str(row.get('login_usuario_consultante', '')),       # Match con 'login_usuario_consultante'
                str(row.get('nombre_usuario_consultante', '')),
                str(row.get('nombre_consulta', '')),
                str(row.get('texto_consulta', '')),
                str(row.get('tipo_consulta', '')),
                str(row.get('tipo_respuesta_consulta', '')),
                str(row.get('consulta_a_usuario_especifico', '')),   # Match con '_especifico'
                str(row.get('login_usuario_consultado', '')),
                str(row.get('tipo_usuario_consultado', '')),
                str(row.get('nombre_usuario_consultado', '')),
                str(row.get('rol_consultado', '')),
                str(row.get('codigo_unidad_consulta', '')),
                str(row.get('nombre_unidad_consulta', '')),
                str(row.get('login_usuario_consultado_respuesta', '')),
                str(row.get('tipo_usuario_consultado_respuesta', '')),
                str(row.get('nombre_usuario_consultado_respuesta', '')),
                str(row.get('fecha_ingreso_consulta', '')),
                str(row.get('fecha_lectura_consulta', '')),
                str(row.get('fecha_revision_consulta', '')),        # Match con '_revision_consulta'
                limpiar_decimal(row.get('tiempo_lectura_consulta', 0.0)),
                limpiar_decimal(row.get('tiempo_revision_consulta', 0.0))
            )
            lista_lote.append(valores)

        # 🚀 Envío masivo mediante bloque único binario
        if lista_lote:
            cursor.executemany(sql, lista_lote)
        
        conexion.commit()
        print(f"✅ Ingesta Exitosa: Se han cargado {len(lista_lote)} filas en la tabla Visacion_compras_P01.")

    except mysql.connector.Error as e:
        conexion.rollback()
        print(f"❌ Excepción MySQL detectada: {e} | Se ejecutó reversión (Rollback).")
    except Exception as e:
        conexion.rollback()
        print(f"❌ Falla general en ejecución: {e} | Se ejecutó reversión (Rollback).")
    finally:
        cursor.close()

# --------------------------------------------------
# PARTE EJECUTABLE
# --------------------------------------------------
if __name__ == '__main__':
    archivo = 'excel/solicitud_compra_slep_visacion.xlsx'
    conn = crear_conexion()
    if conn:
        cargar_visacion_compras_p01(conn, archivo)
        conn.close()