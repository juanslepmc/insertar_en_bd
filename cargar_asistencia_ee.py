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
            password='123456789', 
            database='bd_institucional_v2'
        )
    except Error as e:
        print(f"❌ Error de conexión: {e}")
        return None

# --------------------------------------------------
# UTILIDADES DE LIMPIEZA
# --------------------------------------------------
def formatear_rut(run, dv):
    """ Une el RUN_ALU y el DGV_ALU formando el formato 12345678-K """
    try:
        if pd.isna(run) or pd.isna(dv):
            return None
        run_str = str(run).split('.')[0].strip()
        dv_str = str(dv).strip().upper()
        return f"{run_str}-{dv_str}"
    except:
        return None

def limpiar_decimal_asistencia(valor):
    """ Convierte de forma segura valores a Float, manejando NaN """
    try:
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
            return 0.0
        val_str = str(valor).replace(',', '.')
        return float(val_str)
    except:
        return 0.0

def limpiar_entero(valor):
    """ 
    🛡️ EVITA EL ERROR DE FLOAT NaN TO INTEGER
    Convierte de forma segura a entero, manejando NaN, vacíos y flotantes como '22.0'
    """
    try:
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
            return 0
        # Convertimos primero a float y luego a int por si viene como '22.0'
        return int(float(valor))
    except:
        return 0

# --------------------------------------------------
# PROCESO DE CARGA
# --------------------------------------------------
def cargar_asistencia_incremental(conexion, ruta_archivo):
    cursor = conexion.cursor()

    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ No se encontró el archivo: {ruta_archivo}")
            return

        df = pd.read_excel(
            ruta_archivo, 
            dtype={'RBD': str, 'RUN_ALU': str, 'DGV_ALU': str}
        )

        df.columns = df.columns.str.strip().str.lower()

        if df.empty:
            print("⚠️ Archivo vacío.")
            return

        anio_carga = int(df.iloc[0]['agno'])
        mes_carga = int(df.iloc[0]['mes_escolar'])

        # --- TRANSACCIÓN ---
        print(f"⚡ Iniciando carga de Asistencia: {mes_carga}-{anio_carga}")
        conexion.start_transaction()

        cursor.execute("DELETE FROM Asistencia WHERE anio = %s AND mes = %s", (anio_carga, mes_carga))

        sql_insert = """
            INSERT INTO Asistencia (
                anio, mes, rbd, rut_alumno, dias_asistidos, 
                dias_trabajados, asistencia_promedio
            ) VALUES (%s, %s, %s, %s, %s, %s, %s)
        """
        
        filas_ok = 0
        filas_omitidas = []

        for index, row in df.iterrows():
            rbd_val = str(row.get('rbd', '')).strip()
            rut_generado = formatear_rut(row.get('run_alu'), row.get('dgv_alu'))

            if not rbd_val or rbd_val == 'nan' or not rut_generado:
                fila_err = row.to_dict()
                fila_err['motivo_error'] = "Faltan datos de identidad"
                filas_omitidas.append(fila_err)
                continue

            # 🛠️ Aplicamos las funciones de limpieza para blindar los números
            valores = (
                anio_carga,
                mes_carga,
                rbd_val,
                rut_generado,
                limpiar_entero(row.get('dias_asistidos')),
                limpiar_entero(row.get('dias_matriculados')), 
                limpiar_decimal_asistencia(row.get('pct_asist_sobre_matricula')) 
            )

            try:
                cursor.execute(sql_insert, valores)
                filas_ok += 1
            except mysql.connector.Error as e:
                fila_err = row.to_dict()
                fila_err['rut_procesado'] = rut_generado
                if e.errno == 1452:
                    error_msg = str(e).lower()
                    if 'rbd' in error_msg:
                        fila_err['motivo_error'] = f"RBD {rbd_val} no existe"
                    else:
                        fila_err['motivo_error'] = "Alumno no registrado en la base de datos (Histórico)"
                else:
                    fila_err['motivo_error'] = f"Error MySQL: {e.msg}"
                
                filas_omitidas.append(fila_err)

        conexion.commit()

        if filas_omitidas:
            if not os.path.exists('excel'): os.makedirs('excel')
            df_err = pd.DataFrame(filas_omitidas)
            ruta_log = f'excel/asistencia_omitidos_{mes_carga}_{anio_carga}.xlsx'
            df_err.to_excel(ruta_log, index=False)
            print(f"📄 Log de alumnos no procesados generado: {ruta_log}")

        print(f"✅ CARGA FINALIZADA: {filas_ok} registros insertados.")

    except Exception as e:
        conexion.rollback()
        print(f"❌ ERROR CRÍTICO: {e}")
    finally:
        cursor.close()

if __name__ == '__main__':
    ARCHIVO = 'excel/asistencia/2025/12/12-SLEP-MAULE-COSTA-DICIEMBRE-2025.xlsx' 
    conn = crear_conexion()
    if conn:
        cargar_asistencia_incremental(conn, ARCHIVO)
        conn.close()