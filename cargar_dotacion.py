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
            password='123456789', # Tu contraseña
            database='bd_institucional_v2'
        )
        return conexion
    except Error as e:
        print(f"Error de conexión: {e}")
        return None

def procesar_rut(rut_sucio):
    """ Quita los puntos del RUT: 17.187.030-5 -> 17187030-5 """
    if not rut_sucio: return None
    return str(rut_sucio).replace('.', '').strip()

def transformar_nombre_completo(nombre_original):
    """
    Transforma 'ABARZA ARELLANO, LEONARDO ESTEBAN' 
    en: nombres='Leonardo Esteban', paterno='Abarza', materno='Arellano'
    """
    if not nombre_original or pd.isna(nombre_original) or ',' not in str(nombre_original):
        return nombre_original, "", ""

    # 1. Separar por la coma
    partes = str(nombre_original).split(',')
    apellidos_raw = partes[0].strip() # ABARZA ARELLANO
    nombres_raw = partes[1].strip()   # LEONARDO ESTEBAN

    # 2. Separar apellidos
    lista_apellidos = apellidos_raw.split()
    apellido_paterno = lista_apellidos[0].title() if len(lista_apellidos) > 0 else ""
    apellido_materno = lista_apellidos[1].title() if len(lista_apellidos) > 1 else ""
    
    # 3. Formatear nombres
    nombres = nombres_raw.title()

    return nombres, apellido_paterno, apellido_materno

def cargar_nomina_funcionarios(conexion, ruta_archivo):
    try:
        if not os.path.exists(ruta_archivo):
            print(f"Error: No existe el archivo {ruta_archivo}")
            return

        print("Leyendo Excel y transformando datos...")
        df = pd.read_excel(ruta_archivo)
        
        # --- SOLUCIÓN AL ERROR DE 'nan' ---
        # 1. Convertimos todos los NaN de Pandas a None de Python (para que MySQL reciba NULL)
        df = df.where(pd.notnull(df), None)
        
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
            
            # Transformaciones
            rut_limpio = procesar_rut(row['RUN'])
            nombres, ap_paterno, ap_materno = transformar_nombre_completo(row['Nombre'])
            
            # --- PROTECCIÓN ADICIONAL PARA 'nan' ---
            # Evaluamos si el dato es None antes de forzarlo a string
            fecha_nac = str(row['Fecha Nacimiento']) if row['Fecha Nacimiento'] is not None else None
            fono = str(row['Teléfono']) if row['Teléfono'] is not None else None

            # Datos para el INSERT
            valores = (
                rut_limpio,
                nombres,
                ap_paterno,
                ap_materno,
                row['Sexo'],
                row['Email'],
                fecha_nac, 
                row['Cargo'],
                row['Comuna'],
                fono,
                row['estado']
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
    archivo = 'excel/NominaFuncionario.xlsx'
    conn = crear_conexion()
    if conn:
        cargar_nomina_funcionarios(conn, archivo)
        conn.close()