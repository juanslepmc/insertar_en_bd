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
            password='123456789', 
            database='bd_institucional_v2'
        )
        return conexion
    except Error as e:
        print(f"❌ Error de conexión: {e}")
        return None

def limpiar_numero(valor):
    if valor == "Sin información" or pd.isna(valor) or str(valor).strip() == "":
        return 0.0
    try:
        num = str(valor).replace('.', '').replace(',', '.')
        return float(num)
    except:
        return 0.0

def limpiar_entero(valor):
    if valor == "Sin información" or pd.isna(valor):
        return 0
    try:
        return int(float(valor))
    except:
        return 0

def cargar_compras_p01(conexion, ruta_archivo):
    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ Error: No existe el archivo {ruta_archivo}")
            return

        print(f"🚀 Procesando {ruta_archivo}...")
        df = pd.read_excel(ruta_archivo)
        
        # Normalizamos nombres de columnas (minúsculas y sin espacios)
        df.columns = df.columns.str.strip().str.lower()
        df = df.fillna("Sin información")

        cursor = conexion.cursor()

        # 1. Limpiamos la tabla
        print("🧹 Vaciando tabla y reiniciando contadores...")
        cursor.execute("TRUNCATE TABLE Compras_P01")

        # 2. SQL de Inserción (IMPORTANTE: No incluimos el campo 'id' para que sea AUTO_INCREMENT)
        sql = """
            INSERT INTO Compras_P01 
            (folio, anio, fecha_creacion, fecha_modificacion, nombre, descripcion, 
             fundamento, monto_estimado, tipo_usuario_requirente, nombre_usuario_requirente, 
             nombre_unidad_requirente, estado, nombre_materia_compra, nombre_fuente_financiamiento, 
             fecha_esperada_recepcion, fecha_esperada_compra, nombre_usuario_autorizador_unidad, 
             fecha_aprob_autorizador_unidad, nombre_usuario_ejecutivo_compras, fecha_aprobacion, 
             fecha_finalizacion, monto_total_compra, detalle_fecha_creacion, detalle_nombre, 
             detalle_concepto_presupuestario, detalle_concepto_presupuestario_codigo, 
             detalle_total_compra, detalle_fecha_aprobacion) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """

        stats = {"insertados": 0, "errores": 0}

        print("📥 Cargando datos...")
        for index, row in df.iterrows():
            # Mapeamos los valores (Omitimos el row['id'] del excel en la inserción del PK)
            valores = (
                str(row['folio']),
                limpiar_entero(row['ano']),
                str(row['fecha_creacion']),
                str(row['fecha_modificacion']),
                str(row['nombre']),
                str(row['descripcion']),
                str(row['fundamento']),
                limpiar_numero(row['monto_estimado']),
                str(row['tipo_usuario_requirente']),
                str(row['nombre_usuario_requirente']),
                str(row['nombre_unidad_requirente']),
                str(row['estado']),
                str(row['nombre_materia_compra']),
                str(row['nombre_fuente_financiamiento']),
                str(row['fecha_esperada_recepcion']),
                str(row['fecha_esperada_compra']),
                str(row['nombre_usuario_autorizador_unidad']),
                str(row['fecha_aprob_autorizador_unidad']),
                str(row['nombre_usuario_ejecutivo_compras']),
                str(row['fecha_aprobacion']),
                str(row['fecha_finalizacion']),
                limpiar_numero(row['monto_total_compra']),
                str(row['detalle_fecha_creacion']),
                str(row['detalle_nombre']),
                str(row['detalle_concepto_presupuestario']),
                str(row['detalle_concepto_presupuestario_codigo']),
                limpiar_numero(row['detalle_total_compra']),
                str(row['detalle_fecha_aprobacion'])
            )

            try:
                cursor.execute(sql, valores)
                stats["insertados"] += 1
            except mysql.connector.Error as e:
                stats["errores"] += 1
                print(f"⚠️ Error Fila {index+2}: {e.msg}")

        conexion.commit()
        print("-" * 50)
        print(f"✅ CARGA EXITOSA:")
        print(f"   Filas procesadas: {stats['insertados']}")
        print(f"   ID final generado: {stats['insertados']}")
        print(f"   Errores: {stats['errores']}")

    except Exception as e:
        print(f"💥 Error crítico: {e}")
    finally:
        if 'cursor' in locals(): cursor.close()

if __name__ == '__main__':
    archivo = 'excel/solicitud_compra_slep.xlsx'
    conn = crear_conexion()
    if conn:
        cargar_compras_p01(conn, archivo)
        conn.close()