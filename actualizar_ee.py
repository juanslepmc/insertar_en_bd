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
            return conexion

    except Error as e:
        print(f"Error al intentar conectar a MySQL: {e}")
        return None


def cargar_o_actualizar_excel(conexion, ruta_archivo):
    """
    Lee un Excel y aplica 'Upsert': Inserta si no existe, actualiza si ya existe.
    """
    try:
        if not os.path.exists(ruta_archivo):
            print(f"Error: No se encontró el archivo '{ruta_archivo}'.")
            return

        print(f"Leyendo el archivo '{ruta_archivo}'...\n")
        
        df = pd.read_excel(ruta_archivo)
        df = df.where(pd.notnull(df), None)

        cursor = conexion.cursor()
        
        # CONSULTA MÁGICA: ON DUPLICATE KEY UPDATE
        # VALUES(columna) extrae el valor que se intentó insertar y lo usa para actualizar
        consulta_upsert = """
            INSERT INTO establecimientos 
            (rbd, nombre, comuna, region, direccion, email, telefono, estado) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                nombre = VALUES(nombre),
                comuna = VALUES(comuna),
                region = VALUES(region),
                direccion = VALUES(direccion),
                email = VALUES(email),
                telefono = VALUES(telefono),
                estado = VALUES(estado)
        """
        
        insertados = 0
        actualizados = 0
        sin_cambios = 0
        errores = 0

        print("Iniciando procesamiento de datos...")
        for index, row in df.iterrows():
            fila_excel = index + 2 
            rbd_actual = row['rbd']
            
            try:
                cursor.execute(consulta_upsert, tuple(row))
                
                # Evaluamos qué hizo MySQL exactamente con esta fila
                if cursor.rowcount == 1:
                    insertados += 1
                    print(f"✅ Fila {fila_excel} | RBD: {rbd_actual} -> NUEVO registro insertado.")
                elif cursor.rowcount == 2:
                    actualizados += 1
                    print(f"🔄 Fila {fila_excel} | RBD: {rbd_actual} -> Registro ACTUALIZADO.")
                elif cursor.rowcount == 0:
                    sin_cambios += 1
                    # Puedes comentar el print de abajo si no quieres ver los que no sufrieron cambios
                    print(f"⏩ Fila {fila_excel} | RBD: {rbd_actual} -> Sin cambios (los datos ya eran idénticos).")
                
            except mysql.connector.Error as e:
                errores += 1
                print(f"❌ Error -> Fila Excel {fila_excel} | RBD: {rbd_actual} | Motivo: {e.msg}")
        
        conexion.commit()
        
        print("-" * 60)
        print(f"¡Proceso finalizado con éxito!")
        print(f"✨ Nuevos creados: {insertados}")
        print(f"📝 Actualizados: {actualizados}")
        print(f"⏭️  Omitidos (Sin cambios): {sin_cambios}")
        print(f"⛔ Errores: {errores}\n")

    except Exception as e:
        conexion.rollback()
        print(f"Ocurrió un error crítico inesperado al procesar el Excel: {e}")
    finally:
        if 'cursor' in locals():
            cursor.close()

# ==========================================
# FLUJO PRINCIPAL DE EJECUCIÓN
# ==========================================
if __name__ == '__main__':
    archivo_excel = 'excel/establecimientos_carga.xlsx'
    
    conexion_db = crear_conexion()
    
    if conexion_db:
        try:
            cargar_o_actualizar_excel(conexion_db, archivo_excel)
        finally:
            if conexion_db.is_connected():
                conexion_db.close()
                print("La conexión ha sido cerrada de forma segura.")