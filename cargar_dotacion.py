import mysql.connector
from mysql.connector import Error
import pandas as pd
import os
from datetime import datetime

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

def procesar_rut(rut_sucio):
    """ Quita los puntos del RUT. """
    if rut_sucio == "Sin información": 
        return rut_sucio
    return str(rut_sucio).replace('.', '').strip()

def transformar_nombre_completo(nombre_original):
    """ Transforma el nombre completo a formato Título. """
    if nombre_original == "Sin información" or ',' not in str(nombre_original):
        return nombre_original, "Sin información", "Sin información"

    partes = str(nombre_original).split(',')
    apellidos_raw = partes[0].strip()
    nombres_raw = partes[1].strip()

    lista_apellidos = apellidos_raw.split()
    apellido_paterno = lista_apellidos[0].title() if len(lista_apellidos) > 0 else "Sin información"
    apellido_materno = lista_apellidos[1].title() if len(lista_apellidos) > 1 else "Sin información"
    
    nombres = nombres_raw.title()

    return nombres, apellido_paterno, apellido_materno

def formatear_fecha(fecha_raw):
    """ Convierte fechas a formato dd/mm/yyyy. """
    if pd.isna(fecha_raw) or str(fecha_raw).strip() in ["Sin información", "", "nan", "None"]:
        return "Sin información"
    
    # Si Pandas leyó la celda directamente como Timestamp/datetime
    if isinstance(fecha_raw, (pd.Timestamp, datetime)):
        return fecha_raw.strftime('%d/%m/%Y')
    
    val_str = str(fecha_raw).strip()
    
    # Intentar parseo con pandas para formatos variados (ej: 1990-08-15)
    try:
        dt = pd.to_datetime(val_str, dayfirst=True, format='mixed', errors='coerce')
        if not pd.isna(dt):
            return dt.strftime('%d/%m/%Y')
    except Exception:
        pass

    # Si es una cadena directa como '15-08-1990', se reemplazan guiones por barras
    return val_str.replace('-', '/')

def cargar_nomina_funcionarios(conexion, ruta_archivo):
    try:
        if not os.path.exists(ruta_archivo):
            print(f"Error: No existe el archivo {ruta_archivo}")
            return

        print("Leyendo Excel y normalizando datos...")
        df = pd.read_excel(ruta_archivo)
        
        # 1. Rellenar celdas vacías con "Sin información"
        df = df.fillna("Sin información")

        # 2. Correo: Todo a minúsculas
        df['Email'] = df['Email'].astype(str).str.lower()

        # 3. Cargo: Solo la primera letra en Mayúscula (Sentence case)
        df['Cargo'] = df['Cargo'].astype(str).str.capitalize()
        
        cursor = conexion.cursor()

        sql = """
            INSERT INTO Dotacion_Docentes_AAEE 
            (rut, nombres, apellido_paterno, apellido_materno, sexo, email, 
             fecha_nacimiento, cargo, comuna, telefono, estado) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                nombres = VALUES(nombres),
                apellido_paterno = VALUES(apellido_paterno),
                apellido_materno = VALUES(apellido_materno),
                sexo = VALUES(sexo),
                email = VALUES(email),
                fecha_nacimiento = VALUES(fecha_nacimiento),
                cargo = VALUES(cargo),
                comuna = VALUES(comuna),
                telefono = VALUES(telefono),
                estado = VALUES(estado)
        """

        stats = {"nuevos": 0, "actualizados": 0, "errores": 0}

        for index, row in df.iterrows():
            fila_excel = index + 2
            
            # Transformaciones de lógica interna
            rut_limpio = procesar_rut(row['RUN'])
            nombres, ap_paterno, ap_materno = transformar_nombre_completo(row['Nombre'])
            fecha_nac_formateada = formatear_fecha(row['Fecha Nacimiento'])
            
            valores = (
                rut_limpio,
                nombres,
                ap_paterno,
                ap_materno,
                row['Sexo'],
                row['Email'],
                fecha_nac_formateada, 
                row['Cargo'],
                row['Comuna'],
                str(row['Teléfono']),
                row['Estado']
            )

            try:
                cursor.execute(sql, valores)
                if cursor.rowcount == 1: stats["nuevos"] += 1
                elif cursor.rowcount == 2: stats["actualizados"] += 1
            except mysql.connector.Error as e:
                stats["errores"] += 1
                print(f"❌ Error Fila {fila_excel} (RUT: {rut_limpio}): {e.msg}")

        conexion.commit()
        print("-" * 50)
        print(f"RESULTADO DE CARGA:")
        print(f"✅ Registros nuevos: {stats['nuevos']}")
        print(f"🔄 Registros actualizados: {stats['actualizados']}")
        print(f"⚠️ Errores encontrados: {stats['errores']}")

    except Exception as e:
        print(f"Error crítico: {e}")
    finally:
        if 'cursor' in locals(): cursor.close()

if __name__ == '__main__':
    archivo = 'excel/dotacion/dotacion_faltante_activos.xlsx'
    conn = crear_conexion()
    if conn:
        cargar_nomina_funcionarios(conn, archivo)
        conn.close()