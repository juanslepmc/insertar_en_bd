import mysql.connector
from mysql.connector import Error
import pandas as pd
import os

def crear_conexion():
    try:
        conexion = mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password='123456789', # Ajusta según tu configuración
            database='bd_institucional_v2'
        )
        return conexion
    except Error as e:
        print(f"❌ Error de conexión: {e}")
        return None

def limpiar_numero(valor):
    """ Convierte a decimal. Si falla o es vacío, devuelve 0.0 """
    if pd.isna(valor) or str(valor).strip() == "" or valor == "Sin información":
        return 0.0
    try:
        num = str(valor).replace('.', '').replace(',', '.')
        return float(num)
    except:
        return 0.0

def limpiar_entero(valor):
    """ Convierte a entero. Si falla o es vacío, devuelve 0 """
    if pd.isna(valor) or valor == "Sin información":
        return 0
    try:
        return int(float(valor))
    except:
        return 0

def cargar_adquisicion_p01(conexion, ruta_archivo):
    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ Error: No se encuentra el archivo en {ruta_archivo}")
            return

        print(f"🚀 Leyendo archivo: {ruta_archivo}")
        df = pd.read_excel(ruta_archivo)
        
        # Normalizamos encabezados a minúsculas y sin espacios
        df.columns = df.columns.str.strip().str.lower()
        df = df.fillna("Sin información")

        cursor = conexion.cursor()

        # 1. Limpieza de tabla
        print("🧹 Vaciando tabla adquisicion_P01...")
        cursor.execute("TRUNCATE TABLE adquisicion_P01")

        # 2. SQL de Inserción (Omitimos el campo 'id' para que sea autoincremental)
        # Nota: 'monto_total_documento_compra' no parece estar en el excel, se enviará 0.0
        sql = """
            INSERT INTO adquisicion_P01 
            (folio, anio, fecha_creacion, fecha_modificacion, nombre, descripcion, 
             fecha_esperada_compra, fecha_esperada_recepcion, fecha_firma_jefe_compras, 
             monto_total, monto_total_compra, tipo_documento_compra, numero_documento_compra, 
             fecha_documento_compra, monto_total_documento_compra, razon_social_proveedor, 
             rut_proveedor, estado, nombre_usuario_ejecutivo_compras, nombre_usuario_jefe_compras, 
             modalidad_compra, tipo_modalidad_compra, detalle_folio_solicitud_compra, 
             detalle_nombre_unidad, detalle_codigo_concepto_presupuestario, 
             detalle_nombre_concepto_presupuestario, detalle_nombre_materia_compra, 
             detalle_cantidad, detalle_valor_unitario_neto, detalle_valor_unitario_compra, 
             detalle_total_compra) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """

        stats = {"insertados": 0, "errores": 0}

        print("📥 Cargando datos en la base de datos...")
        for index, row in df.iterrows():
            valores = (
                str(row['folio']),
                limpiar_entero(row['ano']), # En excel es 'ano'
                str(row['fecha_creacion']),
                str(row['fecha_modificacion']),
                str(row['nombre']),
                str(row['descripcion']),
                str(row['fecha_esperada_compra']),
                str(row['fecha_esperada_recepcion']),
                str(row['fecha_firma_jefe_compras']),
                limpiar_numero(row['monto_total']),
                limpiar_numero(row['monto_total_compra']),
                str(row['tipo_documento_compra']),
                str(row['numero_documento_compra']),
                str(row['fecha_documento_compra']),
                0.0, # monto_total_documento_compra (no detectado en origen)
                str(row['razon_social_proveedor']),
                str(row['rut_proveedor']),
                str(row['estado']),
                str(row['nombre_usuario_ejecutivo_compras']),
                str(row['nombre_usuario_jefe_compras']),
                str(row['modalidad_compra']),
                str(row['tipo_modalidad_compra']),
                str(row['detalle_folio_solicitud_compra']),
                str(row['detalle_nombre_unidad']),
                str(row['detalle_codigo_concepto_presupuestario']),
                str(row['detalle_nombre_concepto_presupuestario']),
                str(row['detalle_nombre_materia_compra']),
                limpiar_entero(row['detalle_cantidad']),
                limpiar_numero(row['detalle_valor_unitario_neto']),
                limpiar_numero(row['detalle_valor_unitario_compra']),
                limpiar_numero(row['detalle_total_compra'])
            )

            try:
                cursor.execute(sql, valores)
                stats["insertados"] += 1
            except mysql.connector.Error as e:
                stats["errores"] += 1
                print(f"⚠️ Error en fila {index+2} (Folio {row['folio']}): {e.msg}")

        conexion.commit()
        print("-" * 50)
        print(f"✅ PROCESO FINALIZADO:")
        print(f"   Total registros insertados: {stats['insertados']}")
        print(f"   Errores encontrados: {stats['errores']}")

    except Exception as e:
        print(f"💥 Error crítico durante la carga: {e}")
    finally:
        if 'cursor' in locals(): cursor.close()

if __name__ == '__main__':
    # Ruta especificada: carpeta 'excel'
    #ruta = os.path.join('excel', 'adquisicion_slep.xlsx')
    archivo = 'excel/adquisicion_slep.xlsx'

    conn = crear_conexion()
    if conn:
        cargar_adquisicion_p01(conn, archivo)
        conn.close()