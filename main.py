import mysql.connector
from mysql.connector import Error
import pandas as pd
import os

def crear_conexion():
    """
    Establece y retorna la conexión a la base de datos MySQL local.
    """
    try:
        conexion = mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password='123456789', # Tu contraseña
            database='bd_institucional_v1'
        )

        if conexion.is_connected():
            version_servidor = conexion.get_server_info()
            print(f"¡Conexión exitosa a MySQL! (Versión: {version_servidor})\n")
            return conexion

    except Error as e:
        print(f"Error al intentar conectar a MySQL: {e}")
        return None


def insertar_desde_excel(conexion, ruta_archivo):
    """
    Lee un archivo Excel y carga sus datos masivamente en la tabla 'establecimientos'.
    """
    try:
        # 1. Verificar si el archivo existe
        if not os.path.exists(ruta_archivo):
            print(f"Error: No se encontró el archivo '{ruta_archivo}' en la carpeta actual.")
            return

        print(f"Leyendo el archivo '{ruta_archivo}'...")
        
        # 2. Leer el Excel usando pandas
        df = pd.read_excel(ruta_archivo)
        
        # 3. Limpieza: Pandas lee celdas vacías como 'NaN'. MySQL requiere que sean 'None' (NULL en SQL)
        df = df.where(pd.notnull(df), None)

        # 4. Convertir el DataFrame a una lista de tuplas para la inserción masiva
        datos_a_insertar = [tuple(x) for x in df.to_numpy()]

        cursor = conexion.cursor()
        
        # 5. Escribir la consulta SQL paramétrica
        # NOTA: Usamos INSERT IGNORE para que, si ejecutas el script 2 veces, 
        # no se caiga por intentar insertar un RBD (Primary Key) que ya existe.
        consulta = """
            INSERT IGNORE INTO establecimientos 
            (rbd, nombre, comuna, region, direccion, email, telefono) 
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """
        
        # 6. Ejecutar la inserción masiva
        cursor.executemany(consulta, datos_a_insertar)
        
        # 7. Confirmar los cambios en la base de datos (¡Muy importante para los INSERT!)
        conexion.commit()
        
        print(f"¡Proceso finalizado! Se insertaron {cursor.rowcount} registros nuevos en 'establecimientos'.\n")

    except Error as e:
        # Si hay un error de MySQL, deshacemos la transacción para no dejar datos a medias
        conexion.rollback()
        print(f"Error de base de datos durante la inserción: {e}")
    except Exception as e:
        print(f"Ocurrió un error inesperado al procesar el Excel: {e}")
    finally:
        if 'cursor' in locals():
            cursor.close()


def consultar_establecimientos(conexion):
    """
    Realiza un SELECT a la tabla establecimientos usando una conexión activa.
    """
    try:
        cursor = conexion.cursor(dictionary=True) 
        
        print("Consultando la tabla 'establecimientos' para verificar...\n")

        consulta = "SELECT * FROM establecimientos;"
        cursor.execute(consulta)
        registros = cursor.fetchall()
        
        if len(registros) == 0:
            print("La tabla 'establecimientos' actualmente está vacía.")
        else:
            print(f"Se encontraron {len(registros)} establecimientos en la base de datos:\n")
            for fila in registros:
                print(f"RBD: {fila['rbd']} | Nombre: {fila['nombre']} | Comuna: {fila['comuna']} | Teléfono: {fila['telefono']}")

    except Error as e:
        print(f"Error al ejecutar la consulta: {e}")
        
    finally:
        if 'cursor' in locals():
            cursor.close()


# ==========================================
# FLUJO PRINCIPAL DE EJECUCIÓN
# ==========================================
if __name__ == '__main__':
    # Nombre de tu archivo (asegúrate de que esté en la misma carpeta que este script)
    archivo_excel = 'excel/test_subida_ee.xlsx'
    
    # 1. Establecer conexión
    conexion_db = crear_conexion()
    
    if conexion_db:
        try:
            # 2. Primero: Insertar los datos del Excel
            insertar_desde_excel(conexion_db, archivo_excel)
            
            print("-" * 60)
            
            # 3. Segundo: Consultar la tabla para ver si se guardaron
            consultar_establecimientos(conexion_db)
            
        finally:
            # 4. Tercero: Cerrar la conexión
            if conexion_db.is_connected():
                conexion_db.close()
                print("\nLa conexión ha sido cerrada de forma segura.")