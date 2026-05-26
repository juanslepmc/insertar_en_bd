import mysql.connector
from mysql.connector import Error
import pandas as pd
import os

# --------------------------------------------------
# 1. CONEXIÓN
# --------------------------------------------------
def crear_conexion():
    try:
        return mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password='123456789', 
            database='bd_institucional_v2'
        )
    except Error as e:
        print(f"❌ Error de conexión: {e}")
        return None

# --------------------------------------------------
# 2. UTILIDADES
# --------------------------------------------------
def limpiar_y_formatear_rut(run, dv):
    if pd.isna(run) or pd.isna(dv):
        return None
    # Limpiamos decimales y espacios
    run_str = str(run).split('.')[0].replace('.', '').replace(',', '').strip()
    dv_str = str(dv).strip().upper()
    return f"{run_str}-{dv_str}"

# --------------------------------------------------
# 3. LÓGICA DE CARGA HISTÓRICA
# --------------------------------------------------
def cargar_alumnos_historico(conexion, ruta_archivo):
    cursor = conexion.cursor()
    
    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ No existe el archivo: {ruta_archivo}")
            return

        # --- PASO 1: Obtener RUTs de la BD ---
        print("🔍 Consultando alumnos existentes en la base de datos...")
        cursor.execute("SELECT rut FROM Alumnos")
        ruts_existentes = {str(row[0]).strip().upper() for row in cursor.fetchall()}
        print(f"ℹ️ Se encontraron {len(ruts_existentes)} alumnos registrados actualmente.")

        # --- PASO 2: Leer el Excel ---
        print(f"🚀 Procesando archivo histórico: {ruta_archivo}...")
        df = pd.read_excel(ruta_archivo)
        
        # ⚠️ SOLUCIÓN 1: Pasar todas las columnas a minúsculas para evitar errores tipográficos
        df.columns = df.columns.str.strip().str.lower()
        df = df.fillna("Sin información")

        print(f"📊 Filas totales en el Excel a procesar: {len(df)}")

        # --- PASO 3: Preparar SQL ---
        sql = """
            INSERT INTO Alumnos 
            (rut, nombres, apellido_paterno, apellido_materno, comuna, estado) 
            VALUES (%s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                nombres = VALUES(nombres),
                apellido_paterno = VALUES(apellido_paterno),
                apellido_materno = VALUES(apellido_materno),
                comuna = VALUES(comuna),
                estado = IF(estado = 'Activo', 'Activo', VALUES(estado))
        """

        # ⚠️ SOLUCIÓN 2: Agregamos nuevos contadores para entender qué ignora el script
        stats = {
            "nuevos_inactivos": 0, 
            "actualizados": 0, 
            "sin_cambios": 0, 
            "omitidos_sin_rut": 0, 
            "errores": 0
        }

        # --- PASO 4: Procesar filas ---
        for index, row in df.iterrows():
            # Buscamos en minúsculas debido a la Solución 1
            rut_final = limpiar_y_formatear_rut(row.get('run'), row.get('dígito ver.'))
            
            if not rut_final or rut_final == "Sin información":
                stats["omitidos_sin_rut"] += 1
                continue

            # LÓGICA DE ESTADO
            if rut_final not in ruts_existentes:
                estado_final = 'Inactivo'
            else:
                estado_final = 'Activo' 

            valores = (
                rut_final,
                str(row.get('nombres', 'Sin información')).title(),
                str(row.get('apellido paterno', 'Sin información')).title(),
                str(row.get('apellido materno', 'Sin información')).title(),
                str(row.get('comuna residencia', 'Sin información')).capitalize(),
                estado_final
            )

            try:
                cursor.execute(sql, valores)
                
                # ⚠️ SOLUCIÓN 3: Control total del rowcount
                if cursor.rowcount == 1: 
                    stats["nuevos_inactivos"] += 1
                elif cursor.rowcount == 2: 
                    stats["actualizados"] += 1
                elif cursor.rowcount == 0:
                    stats["sin_cambios"] += 1 # La fila existe pero los datos eran idénticos
                    
            except mysql.connector.Error as e:
                stats["errores"] += 1
                print(f"⚠️ Error en RUT {rut_final}: {e.msg}")

        conexion.commit()

        print("-" * 50)
        print(f"📊 RESUMEN DE PROCESO:")
        print(f"✅ Nuevos insertados (Inactivos): {stats['nuevos_inactivos']}")
        print(f"🔄 Existentes actualizados: {stats['actualizados']}")
        print(f"⏸️ Existentes idénticos (sin cambios): {stats['sin_cambios']}")
        print(f"⏭️ Omitidos por falta de RUT: {stats['omitidos_sin_rut']}")
        print(f"❌ Errores en Base de Datos: {stats['errores']}")
        print("-" * 50)

    except Exception as e:
        conexion.rollback()
        print(f"💥 Error crítico: {e}")
    finally:
        cursor.close()

if __name__ == '__main__':
    ARCHIVO = 'excel/matricula/2025/12/Consolidado_sin_duplicados.xlsx'
    
    conn = crear_conexion()
    if conn:
        cargar_alumnos_historico(conn, ARCHIVO)
        conn.close()