import mysql.connector
from mysql.connector import Error
import pandas as pd
import os

# --------------------------------------------------
# CONFIGURACIÓN DE CONEXIÓN
# --------------------------------------------------
def crear_conexion():
    try:
        return mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password='123456789', # Ajusta tu contraseña
            database='bd_institucional_v2'
        )
    except Error as e:
        print(f"❌ Error de conexión: {e}")
        return None

# --------------------------------------------------
# UTILIDADES
# --------------------------------------------------
def formatear_rut(run, dv):
    """ Une el RUN y el DV formando el formato 12345678-K """
    try:
        if pd.isna(run) or pd.isna(dv):
            return None
        # Quitamos .0 si pandas leyó el número como float y limpiamos espacios
        run_str = str(run).split('.')[0].strip()
        dv_str = str(dv).strip().upper()
        return f"{run_str}-{dv_str}"
    except:
        return None

# --------------------------------------------------
# PROCESO DE CARGA
# --------------------------------------------------
def cargar_matricula_incremental(conexion, ruta_archivo):
    cursor = conexion.cursor()

    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ No se encontró el archivo: {ruta_archivo}")
            return

        print(f"📖 Leyendo archivo: {ruta_archivo}")
        # Forzamos lectura de rbd, run y dv como texto para no perder ceros
        df = pd.read_excel(
            ruta_archivo, 
            dtype={'RBD': str, 'Run': str, 'Dígito Ver.': str}
        )

        # Normalizar nombres de columnas
        df.columns = df.columns.str.strip().str.lower()

        if df.empty:
            print("⚠️ El archivo está vacío.")
            return

        # 1. Detectar periodo (Mes/Año) del archivo para el borrado selectivo
        # Suponemos que todo el archivo corresponde al mismo mes/año
        anio_carga = int(df.iloc[0]['año'])
        mes_carga = int(df.iloc[0]['mes'])

        # 2. Limpieza de datos (fillna selectivo)
        # No tocamos RBD ni RUN/DV para que fallen o se omitan si están vacíos
        cols_clave = ['rbd', 'run', 'dígito ver.']
        cols_rellenables = [c for c in df.columns if c not in cols_clave]
        df[cols_rellenables] = df[cols_rellenables].fillna("Sin información")

        # --- INICIO DE TRANSACCIÓN ---
        print(f"⚡ Iniciando carga para el periodo {mes_carga}-{anio_carga}...")
        conexion.start_transaction()

        # 3. BORRADO SELECTIVO
        # Borramos solo lo que ya existía para ese mes/año para evitar duplicados
        sql_delete = "DELETE FROM Matricula WHERE anio = %s AND mes = %s"
        cursor.execute(sql_delete, (anio_carga, mes_carga))
        print(f"🧹 Se eliminaron registros previos del mes {mes_carga} para reemplazarlos.")

        # 4. INSERCIÓN
        sql_insert = """
            INSERT INTO Matricula (
                rbd, rut_alumno, nivel, letra, anio, mes
            ) VALUES (%s, %s, %s, %s, %s, %s)
        """
        
        filas_ok = 0
        filas_omitidas = []

        for index, row in df.iterrows():
            rbd_val = str(row.get('rbd', '')).strip()
            rut_generado = formatear_rut(row.get('run'), row.get('dígito ver.'))

            # Validación de datos críticos
            if not rbd_val or rbd_val == 'nan' or not rut_generado:
                fila_err = row.to_dict()
                fila_err['motivo_error'] = "RBD o RUT ausente"
                filas_omitidas.append(fila_err)
                continue

            valores = (
                rbd_val,
                rut_generado,
                str(row.get('desc grado', 'Sin información')),
                str(row.get('letra curso', 'Sin información')),
                anio_carga,
                mes_carga
            )

            try:
                cursor.execute(sql_insert, valores)
                filas_ok += 1
            except mysql.connector.Error as e:
                # Si hay error de Llave Foránea (FK) o cualquier otro
                fila_err = row.to_dict()
                fila_err['rut_generado'] = rut_generado
                
                if e.errno == 1452:
                    msg_error = str(e).lower()
                    if 'rbd' in msg_error:
                        fila_err['motivo_error'] = f"RBD {rbd_val} no existe en Establecimientos"
                    else:
                        fila_err['motivo_error'] = f"RUT {rut_generado} no existe en Alumnos"
                else:
                    fila_err['motivo_error'] = f"Error MySQL: {e.msg}"
                
                filas_omitidas.append(fila_err)

        # Finalizar transacción
        conexion.commit()

        # 5. GENERAR LOG DE ERRORES
        if filas_omitidas:
            if not os.path.exists('excel'): os.makedirs('excel')
            df_err = pd.DataFrame(filas_omitidas)
            ruta_log = 'excel/log_errores_matricula.xlsx'
            df_err.to_excel(ruta_log, index=False)
            print(f"📄 Se generó el archivo de registros omitidos/errores: {ruta_log}")

        print("-" * 50)
        print(f"✅ CARGA FINALIZADA PARA EL MES {mes_carga}/{anio_carga}")
        print(f"   Insertados: {filas_ok}")
        print(f"   Omitidos: {len(filas_omitidas)}")
        print("-" * 50)

    except Exception as e:
        conexion.rollback()
        print(f"❌ PROCESO FALLIDO: {e}")
    finally:
        cursor.close()

# --------------------------------------------------
# EJECUCIÓN
# --------------------------------------------------
if __name__ == '__main__':
    archivo_input = 'excel/matricula/2025/12/Consolidado_sin_duplicados.xlsx'
    conn = crear_conexion()
    
    if conn:
        cargar_matricula_incremental(conn, archivo_input)
        conn.close()