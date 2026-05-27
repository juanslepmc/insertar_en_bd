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
# UTILIDADES DE SOPORTE Y LIMPIEZA
# --------------------------------------------------
def obtener_valor(row, *nombres_columnas):
    """ 
    Busca de forma flexible entre múltiples nombres de columnas 
    para dar soporte a variaciones entre Hoja1 y Hoja2.
    """
    for col in nombres_columnas:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != "":
            return row[col]
    return None

def limpiar_texto(valor):
    """ Reemplaza nulos o celdas vacías por 'Sin información' """
    if valor is None or pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
        return "Sin información"
    return str(valor).strip()

def limpiar_entero(valor):
    """ Garantiza un valor entero, si es NaN o vacío devuelve 0 """
    try:
        if valor is None or pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
            return 0
        return int(float(valor))
    except:
        return 0

def limpiar_fecha(valor):
    """ Formatea fechas de manera segura a YYYY-MM-DD """
    if valor is None or pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
        return "Sin información"
    try:
        val_str = str(valor).strip().split(" ")[0]
        fecha_dt = pd.to_datetime(val_str)
        return fecha_dt.strftime('%Y-%m-%d')
    except:
        return "Sin información"

def extraer_rbd(valor):
    """ Toma '3172 - Liceo...' y extrae únicamente '3172' """
    if valor is None or pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
        return "Sin información"
    try:
        partes = str(valor).split('-')
        rbd_candidato = partes[0].strip()
        if rbd_candidato.isdigit():
            return rbd_candidato
        return "Sin información"
    except:
        return "Sin información"

# --------------------------------------------------
# PROCESO PRINCIPAL DE CARGA: ASISTENTES
# --------------------------------------------------
def cargar_asistentes_historicos(conexion, ruta_archivo):
    cursor = conexion.cursor()
    
    if not os.path.exists(ruta_archivo):
        print(f"❌ No se encontró el archivo Excel en: {ruta_archivo}")
        return

    print(f"🔍 Leyendo el archivo de asistentes de la educación...")
    # Puedes cambiar sheet_name=0 por sheet_name=1 si prefieres cargar la Hoja 2 específicamente
    df = pd.read_excel(ruta_archivo, sheet_name=0)

    if df.empty:
        print("⚠️ El archivo Excel no contiene datos.")
        return

    print(f"⚡ Iniciando proceso de inserción para {len(df)} filas de asistentes...")
    conexion.start_transaction()

    sql_insert = """
        INSERT INTO reemplazos (
            fecha_ingreso_solicitud, rbd, tipo_solicitud, motivo, 
            fecha_inicio, fecha_termino, nombre_docente_aaee_reemplazar, rut_docente_aaee_reemplazar, 
            funcion_profesional_docente_aaee_reemplazar, nombre_completo_reemplazante, rut_docente_aaee_reemplazante, 
            comuna, funcion_profesional_reemplazante, cantidad_horas_totales, subv_general, 
            sep, pie, vtf, docente_horas, fecha_visacion_utp, 
            fecha_visacion_gdp, fecha_ingreso_funcionario, tipo_horas_docente_aprobadas, fecha_inicio_reposo, 
            fecha_termino_reposo, estado_uatp, estado_gdp, tipo_de_profesional
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """

    filas_exitosas = 0
    filas_erroneas = []

    for index, row in df.iterrows():
        try:
            # 1. Extracción con alias dinámicos (Soporta Hoja 1 y Hoja 2)
            fecha_ingreso_solicitud = limpiar_fecha(obtener_valor(row, 'Fecha ingreso solicitud', 'Fecha Ingreso Solicitud'))
            rbd = extraer_rbd(obtener_valor(row, 'Establecimiento'))
            tipo_solicitud = limpiar_texto(obtener_valor(row, 'Tipo solicitud'))
            motivo = limpiar_texto(obtener_valor(row, 'Motivo'))
            fecha_inicio = limpiar_fecha(obtener_valor(row, 'Fecha inicio'))
            fecha_termino = limpiar_fecha(obtener_valor(row, 'Fecha término', 'Fecha termino'))
            
            nombre_docente_aaee_reemplazar = limpiar_texto(obtener_valor(row, 'Nombre docente o asistente a reemplazar', 'Nombre docente aaee a reemplazar'))
            rut_docente_aaee_reemplazar = limpiar_texto(obtener_valor(row, 'RUT'))
            funcion_profesional_docente_aaee_reemplazar = limpiar_texto(obtener_valor(row, 'Función profesional'))
            
            nombre_completo_reemplazante = limpiar_texto(obtener_valor(row, 'Nombre completo del reemplazante', 'Nombre Reemplazante'))
            rut_docente_aaee_reemplazante = limpiar_texto(obtener_valor(row, 'RUT2', 'rut reemplazante'))
            comuna = limpiar_texto(obtener_valor(row, 'COMUNA'))
            funcion_profesional_reemplazante = limpiar_texto(obtener_valor(row, 'Función profesional2'))
            
            # Valores numéricos
            cantidad_horas_totales = limpiar_entero(obtener_valor(row, 'Cantidad de horas totales'))
            subv_general = limpiar_entero(obtener_valor(row, 'Subv Genral', 'Subv General'))
            sep = limpiar_entero(obtener_valor(row, 'SEP'))
            pie = limpiar_entero(obtener_valor(row, 'PIE'))
            vtf = limpiar_entero(obtener_valor(row, 'VTF'))
            
            # REQUERIMIENTO: Guardar la columna 'Asistente Estamento' en el campo 'docente_horas'
            docente_horas = limpiar_texto(obtener_valor(row, 'Asistente Estamento'))
            
            fecha_visacion_utp = limpiar_fecha(obtener_valor(row, 'Fecha Estado UATP'))
            fecha_inicio_reposo = limpiar_fecha(obtener_valor(row, 'Fecha de Inicio Reposo'))
            fecha_termino_reposo = limpiar_fecha(obtener_valor(row, 'Fecha de Termino Reposo'))
            estado_uatp = limpiar_texto(obtener_valor(row, 'V°B° UATP'))

            # 2. COLUMNAS NO PRESENTES EN EL EXCEL (Regla 'Sin información')
            fecha_visacion_gdp = "Sin información"
            fecha_ingreso_funcionario = "Sin información"
            tipo_horas_docente_aprobadas = "Sin información"
            estado_gdp = "Sin información"
            
            # REQUERIMIENTO: Tipo de profesional estático para este grupo
            tipo_de_profesional = "Asistente de la educación"

            # 3. Empaquetado ordenado para la consulta SQL (Sin id_folio ya que es AUTO_INCREMENT)
            valores = (
                fecha_ingreso_solicitud, rbd, tipo_solicitud, motivo,
                fecha_inicio, fecha_termino, nombre_docente_aaee_reemplazar, rut_docente_aaee_reemplazar,
                funcion_profesional_docente_aaee_reemplazar, nombre_completo_reemplazante, rut_docente_aaee_reemplazante,
                comuna, funcion_profesional_reemplazante, cantidad_horas_totales, subv_general,
                sep, pie, vtf, docente_horas, fecha_visacion_utp,
                fecha_visacion_gdp, fecha_ingreso_funcionario, tipo_horas_docente_aprobadas, fecha_inicio_reposo,
                fecha_termino_reposo, estado_uatp, estado_gdp, tipo_de_profesional
            )

            # 4. Ejecución del Insert
            cursor.execute(sql_insert, valores)
            filas_exitosas += 1

        except Exception as e:
            fila_con_error = row.to_dict()
            fila_con_error['MOTIVO_ERROR_PROCESO'] = str(e)
            filas_erroneas.append(fila_con_error)

    conexion.commit()

    # Generación de log si hay caídas
    if filas_erroneas:
        if not os.path.exists('excel'): 
            os.makedirs('excel')
        df_err = pd.DataFrame(filas_erroneas)
        ruta_log = 'excel/reemplazos/asistentes-sin-ingresar.xlsx'
        df_err.to_excel(ruta_log, index=False)
        print(f"⚠️ Alerta: {len(filas_erroneas)} filas de asistentes no pudieron ser ingresadas.")
        print(f"📄 Registro detallado generado en: {ruta_log}")

    print(f"✅ PROCESO CONCLUIDO: {filas_exitosas} registros de Asistentes guardados exitosamente.")
    cursor.close()

if __name__ == '__main__':
    # Configura aquí el nombre de tu archivo de asistentes
    ARCHIVO = 'excel/reemplazos/sol-asistentes-historico.xlsx' 
    conn = crear_conexion()
    if conn:
        cargar_asistentes_historicos(conn, ARCHIVO)
        conn.close()