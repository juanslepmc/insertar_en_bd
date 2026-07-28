import mysql.connector
from mysql.connector import Error
import pandas as pd
import os

# --------------------------------------------------
# CONEXIÓN INSTITUCIONAL
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
# TRATAMIENTO CONTROLADO DE DATOS NUMÉRICOS
# --------------------------------------------------
def limpiar_entero(valor):
    """Maneja números enteros normales y de alta capacidad (BIGINT)"""
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "sin información":
        return 0
    try:
        return int(float(valor))
    except:
        return 0

# --------------------------------------------------
# PROCESO PRINCIPAL: ETL ESTABLECIMIENTOS (EE)
# --------------------------------------------------
def cargar_visacion_compras_ee(conexion, ruta_archivo):
    cursor = conexion.cursor()

    try:
        # Carga inicial desde archivo
        df = pd.read_excel(ruta_archivo)

        # Estandarización preventiva de cabeceras
        df.columns = df.columns.str.strip().str.lower()
        
        # Sanitización de campos nulos en texto
        df = df.fillna("Sin información")
        
        # Apertura de transacción
        conexion.start_transaction()
        
        # Vaciado controlado de la tabla destino
        cursor.execute("DELETE FROM Visacion_compras_por_EE")

        # Query estructurada respetando de forma estricta tu lista de 48 atributos
        sql = """
            INSERT INTO Visacion_compras_por_EE (
                id_solicitud_compra, anio_solicitud_compra, folio_solicitud_compra,
                codigo_tarea, nombre_tarea, tipo_visacion, indice_en_ciclo,
                fecha_ingreso, fecha_lectura, fecha_revision, tiempo_lectura,
                tiempo_revision, tarea_valida, numero_ciclo, codigo_division,
                nombre_division, codigo_unidad, nombre_unidad, rol_usuario_asignado,
                login_usuario_asignado, tipo_usuario_asignado, nombre_usuario_asignado, login_usuario_finaliza_tarea,
                tipo_usuario_finaliza_tarea, nombre_usuario_finaliza_tarea, login_en_consultante, visacion_en_consultante,
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
            # Mapeo posicional ordenado uno a uno con la tabla destino
            valores = (
                str(row.get('id_solicitud_compra', '')),
                limpiar_entero(row.get('ano_solicitud_compra', 0)),  # Excel: ano_
                str(row.get('folio_solicitud_compra', '')),
                str(row.get('codigo_tarea', '')),
                str(row.get('nombre_tarea', '')),
                str(row.get('tipo_visacion', '')),
                limpiar_entero(row.get('indice_en_ciclo', 0)),
                str(row.get('fecha_ingreso', '')),
                str(row.get('fecha_lectura', '')),
                str(row.get('fecha_revision', '')),
                limpiar_entero(row.get('tiempo_lectura', 0)),       # Tratado como BIGINT
                limpiar_entero(row.get('tiempo_revision', 0)),      # Tratado como BIGINT
                str(row.get('tarea_valida', '')),
                limpiar_entero(row.get('numero_ciclo', 0)),
                str(row.get('codigo_division', '')),
                str(row.get('nombre_division', '')),
                str(row.get('codigo_unidad', '')),
                str(row.get('nombre_unidad', '')),
                str(row.get('rol_usuario_asignado', '')),           # Excel: _asignado
                str(row.get('login_usuario_asignado', '')),
                str(row.get('tipo_usuario_asignado', '')),
                str(row.get('nombre_usuario_asignado', '')),
                str(row.get('login_usuario_finaliza_tarea', '')),
                str(row.get('tipo_usuario_finaliza_tarea', '')),
                str(row.get('nombre_usuario_finaliza_tarea', '')),
                str(row.get('login_usuario_consultante', '')),      # Excel: login_usuario_...
                str(row.get('visacion_en_consultas', '')),          # Excel: _en_consultas
                str(row.get('nombre_usuario_consultante', '')),
                str(row.get('nombre_consulta', '')),
                str(row.get('texto_consulta', '')),
                str(row.get('tipo_consulta', '')),
                str(row.get('tipo_respuesta_consulta', '')),
                str(row.get('consulta_a_usuario_especifico', '')),   # Excel: _especifico
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
                str(row.get('fecha_revision_consulta', '')),        # Excel: fecha_revision_...
                limpiar_entero(row.get('tiempo_lectura_consulta', 0)), # Tratado como BIGINT
                limpiar_entero(row.get('tiempo_revision_consulta', 0))  # Tratado como BIGINT
            )
            lista_lote.append(valores)

        # Inserción en masa de alto rendimiento
        if lista_lote:
            cursor.executemany(sql, lista_lote)
        
        conexion.commit()
        print(f"✅ Éxito absoluto: Se cargaron {len(lista_lote)} registros en Visacion_compras_por_EE sin errores.")

    except mysql.connector.Error as e:
        conexion.rollback()
        print(f"❌ Error crítico de MySQL: {e} | Rollback ejecutado de forma segura.")
    except Exception as e:
        conexion.rollback()
        print(f"❌ Fallo general en Python: {e} | Rollback ejecutado de forma segura.")
    finally:
        cursor.close()

# --------------------------------------------------
# PUNTO DE ARRANQUE
# --------------------------------------------------
if __name__ == '__main__':
    archivo = 'excel/solicitud_compra_rbd_visacion.xlsx'
    conn = crear_conexion()
    if conn:
        cargar_visacion_compras_ee(conn, archivo)
        conn.close()