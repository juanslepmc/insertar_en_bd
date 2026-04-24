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
        print(f"Error de conexión: {e}")
        return None

def limpiar_y_formatear_rut(run, dv):
    """ 
    Limpia puntos del RUN y lo concatena con el DV usando un guion.
    Ejemplo: 18.222.333 y K -> 18222333-K
    """
    if str(run) == "Sin información" or str(dv) == "Sin información":
        return "Sin información"
        
    run_limpio = str(run).replace('.', '').replace(',', '').strip()
    dv_limpio = str(dv).strip()
    
    return f"{run_limpio}-{dv_limpio}"

def cargar_alumnos(conexion, ruta_archivo):
    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ Error: No existe el archivo en la ruta: {ruta_archivo}")
            return

        print(f"🚀 Procesando archivo: {ruta_archivo}...")
        df = pd.read_excel(ruta_archivo)
        
        # 1. Rellenar celdas vacías con "Sin información"
        df = df.fillna("Sin información")

        # 2. Normalizar formatos de texto
        # Nombres y Apellidos: Formato Título (Juan Perez)
        df['Nombres'] = df['Nombres'].astype(str).str.title()
        df['Apellido Paterno'] = df['Apellido Paterno'].astype(str).str.title()
        df['Apellido Materno'] = df['Apellido Materno'].astype(str).str.title()
        
        # Comuna: Formato Oración (Santiago)
        df['Comuna Residencia'] = df['Comuna Residencia'].astype(str).str.capitalize()

        cursor = conexion.cursor()

        # SQL para la tabla Alumnos
        sql = """
            INSERT INTO Alumnos 
            (rut, nombres, apellido_paterno, apellido_materno, comuna, estado) 
            VALUES (%s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                nombres = VALUES(nombres),
                apellido_paterno = VALUES(apellido_paterno),
                apellido_materno = VALUES(apellido_materno),
                comuna = VALUES(comuna),
                estado = VALUES(estado)
        """

        stats = {"nuevos": 0, "actualizados": 0, "errores": 0}

        for index, row in df.iterrows():
            fila_excel = index + 2
            
            # Formatear el RUT combinando Run + Dígito Ver.
            rut_final = limpiar_y_formatear_rut(row['Run'], row['Dígito Ver.'])
            
            # Manejo del campo 'estado' (si no existe en Excel, usa "Sin información")
            estado_valor = row.get('estado', 'Sin información')

            valores = (
                rut_final,
                row['Nombres'],
                row['Apellido Paterno'],
                row['Apellido Materno'],
                row['Comuna Residencia'],
                estado_valor
            )

            try:
                cursor.execute(sql, valores)
                if cursor.rowcount == 1: stats["nuevos"] += 1
                elif cursor.rowcount == 2: stats["actualizados"] += 1
            except mysql.connector.Error as e:
                stats["errores"] += 1
                print(f"⚠️ Error Fila {fila_excel} (RUT: {rut_final}): {e.msg}")

        conexion.commit()
        print("-" * 50)
        print(f"RESUMEN DE PROCESO:")
        print(f"✅ Registros nuevos: {stats['nuevos']}")
        print(f"🔄 Registros actualizados: {stats['actualizados']}")
        print(f"❌ Errores encontrados: {stats['errores']}")

    except Exception as e:
        print(f"💥 Error crítico durante la carga: {e}")
    finally:
        if 'cursor' in locals(): cursor.close()

if __name__ == '__main__':
    # Configuración del nombre del archivo solicitado
    nombre_archivo = 'excel/Consolidado_sin_duplicados.xlsx'    
    # Si el archivo está dentro de una carpeta, agrégala aquí, ej: 'excel/Consolidado_sin_duplicados.xlsx'
    ruta_completa = nombre_archivo 
    
    conn = crear_conexion()
    if conn:
        cargar_alumnos(conn, ruta_completa)
        conn.close()