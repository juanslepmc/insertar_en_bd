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
            database='bd_institucional_v2'
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
    Lee un archivo Excel y carga sus datos de forma iterativa para reportar errores por fila.
    """
    try:
        # 1. Verificar si el archivo existe
        if not os.path.exists(ruta_archivo):
            print(f"Error: No se encontró el archivo '{ruta_archivo}' en la carpeta actual.")
            return

        print(f"Leyendo el archivo '{ruta_archivo}'...\n")
        
        # 2. Leer el Excel usando pandas
        df = pd.read_excel(ruta_archivo)
        
        # 3. Limpieza: Pandas lee celdas vacías como 'NaN'. MySQL requiere que sean 'None'
        df = df.where(pd.notnull(df), None)

        cursor = conexion.cursor()
        
        # 4. Escribir la consulta SQL paramétrica (Quitamos el IGNORE para forzar el error)
        consulta = """
            INSERT INTO establecimientos 
            (rbd, nombre, comuna, region, direccion, email, telefono, estado) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """
        
        insertados = 0
        fallidos = 0

        # 5. Iterar fila por fila usando pandas iterrows()
        print("Iniciando carga de datos...")
        for index, row in df.iterrows():
            # index empieza en 0, y el Excel suele tener cabecera en la fila 1, por lo que los datos parten en la 2
            fila_excel = index + 2 
            rbd_actual = row['rbd']
            
            try:
                # Intentamos insertar la fila actual
                cursor.execute(consulta, tuple(row))
                insertados += 1
                
            except mysql.connector.Error as e:
                fallidos += 1
                # El código de error 1062 en MySQL significa "Duplicate entry" (RBD ya existe)
                if e.errno == 1062:
                    print(f"⚠️ Omitido -> Fila Excel {fila_excel} | RBD: {rbd_actual} | Motivo: El RBD ya existe en la BD.")
                else:
                    # Capturamos cualquier otro tipo de error (ej: un string demasiado largo)
                    print(f"❌ Error -> Fila Excel {fila_excel} | RBD: {rbd_actual} | Motivo: {e.msg}")
        
        # 6. Confirmar los cambios en la base de datos
        conexion.commit()
        
        print("-" * 60)
        print(f"¡Proceso finalizado!")
        print(f"✅ Registros insertados exitosamente: {insertados}")
        print(f"⛔ Registros omitidos o fallidos: {fallidos}\n")

    except Exception as e:
        # Hacemos rollback solo si el error es general, no de una fila específica
        conexion.rollback()
        print(f"Ocurrió un error crítico inesperado al procesar el Excel: {e}")
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
                print(f"RBD: {fila['rbd']} | Nombre: {fila['nombre']} | Comuna: {fila['comuna']} | Teléfono: {fila['telefono']} | Estado: {fila['estado']}")

    except Error as e:
        print(f"Error al ejecutar la consulta: {e}")
        
    finally:
        if 'cursor' in locals():
            cursor.close()


# ==========================================
# FLUJO PRINCIPAL DE EJECUCIÓN
# ==========================================
if __name__ == '__main__':
    # Nombre de tu archivo
    archivo_excel = 'excel/establecimientos_carga.xlsx'
    
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