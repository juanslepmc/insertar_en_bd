import mysql.connector
from mysql.connector import Error
import pandas as pd
import os

# --------------------------------------------------
# CONFIGURACIÓN GLOBAL DE CONTROL DE TIPOS
# --------------------------------------------------
# Si las columnas de tipo FECHA en tu MySQL son estrictamente de tipo 'DATE', 
# usar el texto "Sin información" romperá la base de datos. 
# - True: Convierte fechas faltantes a NULL (None) para que MySQL las acepte de forma nativa.
# - False: Convierte fechas faltantes al texto "Sin información" (Úsalo solo si tus columnas son VARCHAR).
FECHAS_COMO_NULL = True

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
# UTILIDADES DE LIMPIEZA Y ESCUDOS ANTI-NAN
# --------------------------------------------------
def limpiar_texto(valor):
    """ Reemplaza nulos o celdas vacías por 'Sin información' """
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
        return "Sin información"
    return str(valor).strip()

def limpiar_entero(valor):
    """ Garantiza un valor entero, si es NaN o vacío devuelve 0 """
    try:
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
            return 0
        return int(float(valor))
    except:
        return 0

def limpiar_fecha(valor):
    """ Formatea fechas de manera segura a YYYY-MM-DD """
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
        return None if FECHAS_COMO_NULL else "Sin información"
    try:
        val_str = str(valor).strip().split(" ")[0]
        fecha_dt = pd.to_datetime(val_str)
        return fecha_dt.strftime('%Y-%m-%d')
    except:
        return None if FECHAS_COMO_NULL else "Sin información"

def extraer_rbd(valor):
    """ Toma '3172 - Liceo...' y extrae únicamente '3172' """
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
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
# PROCESO PRINCIPAL DE CARGA
# --------------------------------------------------
def cargar_reemplazos_historicos(conexion, ruta_archivo):
    cursor = conexion.cursor()
    
    if not os.path.exists(ruta_archivo):
        print(f"❌ No se encontró el archivo Excel en: {ruta_archivo}")
        return

    print(f"🔍 Leyendo el archivo de reemplazos...")
    df = pd.read_excel(ruta_archivo)

    if df.empty:
        print("⚠️ El archivo Excel no contiene datos.")
        return

    print(f"⚡ Iniciando proceso de inserción para {len(df)} filas...")
    conexion.start_transaction()

    sql_insert = """
        INSERT INTO reemplazos (
            id_folio, fecha_ingreso_solicitud, rbd, tipo_solicitud, motivo, 
            fecha_inicio, fecha_termino, nombre_docente_aaee_reemplazar, rut_docente_aaee_reemplazar, 
            funcion_profesional_docente_aaee_reemplazar, nombre_completo_reemplazante, rut_docente_aaee_reemplazante, 
            comuna, funcion_profesional_reemplazante, cantidad_horas_totales, subv_general, 
            sep, pie, vtf, docente_horas, fecha_visacion_utp, 
            fecha_visacion_gdp, fecha_ingreso_funcionario, tipo_horas_docente_aprobadas, fecha_inicio_reposo, 
            fecha_termino_reposo, estado_uatp, estado_gdp, tipo_de_profesional
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """

    filas_exitosas = 0
    filas_erroneas = []

    for index, row in df.iterrows():
        try:
            # 1. Extracción y Limpieza de columnas del Excel
            id_folio = limpiar_entero(row.get('Folio'))
            fecha_ingreso_solicitud = limpiar_fecha(row.get('Fecha Ingreso Solicitud'))
            rbd = extraer_rbd(row.get('Establecimiento'))
            tipo_solicitud = limpiar_texto(row.get('Tipo solicitud'))
            motivo = limpiar_texto(row.get('Motivo'))
            fecha_inicio = limpiar_fecha(row.get('Fecha inicio'))
            fecha_termino = limpiar_fecha(row.get('Fecha término'))
            nombre_docente_aaee_reemplazar = limpiar_texto(row.get('Nombre docente aaee a reemplazar'))
            rut_docente_aaee_reemplazar = limpiar_texto(row.get('RUT'))
            funcion_profesional_docente_aaee_reemplazar = limpiar_texto(row.get('Función profesional'))
            nombre_completo_reemplazante = limpiar_texto(row.get('Nombre completo del reemplazante'))
            rut_docente_aaee_reemplazante = limpiar_texto(row.get('rut reemplazante'))
            comuna = limpiar_texto(row.get('COMUNA'))
            funcion_profesional_reemplazante = limpiar_texto(row.get('Función profesional2'))
            
            # Valores numéricos blindados (vacíos pasan como 0)
            cantidad_horas_totales = limpiar_entero(row.get('Cantidad de horas totales'))
            subv_general = limpiar_entero(row.get('Subv Genral'))
            sep = limpiar_entero(row.get('SEP'))
            pie = limpiar_entero(row.get('PIE'))
            vtf = limpiar_entero(row.get('VTF'))
            
            docente_horas = limpiar_texto(row.get('Docente horas'))
            fecha_visacion_utp = limpiar_fecha(row.get('Fecha Estado UATP'))
            fecha_inicio_reposo = limpiar_fecha(row.get('Fecha de Inicio Reposo'))
            fecha_termino_reposo = limpiar_fecha(row.get('Fecha de Termino Reposo'))
            estado_uatp = limpiar_texto(row.get('V°B° UATP'))

            # 2. COLUMNAS QUE NO VIENEN EN EL EXCEL (Sujetas a regla 'Sin información')
            fecha_visacion_gdp = None if FECHAS_COMO_NULL else "Sin información"
            fecha_ingreso_funcionario = None if FECHAS_COMO_NULL else "Sin información"
            tipo_horas_docente_aprobadas = "Sin información"
            estado_gdp = "Sin información"
            
            # Atributo con valor por defecto obligatorio solicitado previamente
            tipo_de_profesional = "Docente"

            # 3. Empaquetado de parámetros ordenados para MySQL
            valores = (
                id_folio, fecha_ingreso_solicitud, rbd, tipo_solicitud, motivo,
                fecha_inicio, fecha_termino, nombre_docente_aaee_reemplazar, rut_docente_aaee_reemplazar,
                funcion_profesional_docente_aaee_reemplazar, nombre_completo_reemplazante, rut_docente_aaee_reemplazante,
                comuna, funcion_profesional_reemplazante, cantidad_horas_totales, subv_general,
                sep, pie, vtf, docente_horas, fecha_visacion_utp,
                fecha_visacion_gdp, fecha_ingreso_funcionario, tipo_horas_docente_aprobadas, fecha_inicio_reposo,
                fecha_termino_reposo, estado_uatp, estado_gdp, tipo_de_profesional
            )

            # 4. Inserción
            cursor.execute(sql_insert, valores)
            filas_exitosas += 1

        except Exception as e:
            # Captura de filas con errores estructurales o de claves duplicadas
            fila_con_error = row.to_dict()
            fila_con_error['MOTIVO_ERROR_PROCESO'] = str(e)
            filas_erroneas.append(fila_con_error)

    # Confirmar cambios exitosos en la BD
    conexion.commit()

    # 5. Reporte de Log de errores si corresponde
    if filas_erroneas:
        if not os.path.exists('excel'): 
            os.makedirs('excel')
        df_err = pd.DataFrame(filas_erroneas)
        ruta_log = 'excel/reemplazos/reemplazos-sin-ingresar.xlsx'
        df_err.to_excel(ruta_log, index=False)
        print(f"⚠️ Alerta: {len(filas_erroneas)} filas no pudieron ser ingresadas.")
        print(f"📄 Registro detallado generado en: {ruta_log}")

    print(f"✅ PROCESO CONCLUIDO: {filas_exitosas} registros de reemplazo almacenados correctamente.")
    cursor.close()

if __name__ == '__main__':
    ARCHIVO = 'excel/reemplazos/reemplazos-docente-historico.xlsx' 
    conn = crear_conexion()
    if conn:
        cargar_reemplazos_historicos(conn, ARCHIVO)
        conn.close()