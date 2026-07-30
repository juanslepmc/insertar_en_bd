import mysql.connector
from mysql.connector import Error
import pandas as pd
import os
import unicodedata
import re

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
# UTILIDADES DE SOPORTE Y LIMPIEZA BLINDADAS
# --------------------------------------------------
def normalizar_cadena(texto):
    """ Remueve acentos, mayúsculas, espacios y caracteres especiales de las cabeceras """
    if texto is None or pd.isna(texto):
        return ""
    # Quitar acentos (Ej: Término -> Termino)
    texto = ''.join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn')
    # Convertir a minúsculas y remover todo lo que no sea letras o números (quita espacios, _, V°B°, etc.)
    texto = texto.lower().strip()
    texto = re.sub(r'[^a-z0-9]', '', texto)
    return texto

def obtener_valor(row, *nombres_columnas):
    """ Busca de forma flexible limpiando tanto la cabecera buscada como la del Excel """
    for col in nombres_columnas:
        clean_col = normalizar_cadena(col)
        if clean_col in row and pd.notna(row[clean_col]) and str(row[clean_col]).strip() != "":
            return row[clean_col]
    return None

def limpiar_texto(valor):
    if valor is None or pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
        return "Sin información"
    return str(valor).strip()

def limpiar_entero(valor):
    try:
        if valor is None or pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
            return 0
        return int(float(valor))
    except:
        return 0

def limpiar_fecha(valor):
    if valor is None or pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan" or str(valor) == "Sin información":
        return "Sin información"
    try:
        val_str = str(valor).strip().split(" ")[0]
        fecha_dt = pd.to_datetime(val_str, errors='coerce')
        if pd.isna(fecha_dt):
            return str(valor).strip()
        return fecha_dt.strftime('%Y-%m-%d')
    except:
        return str(valor).strip()

def extraer_rbd(valor):
    if valor is None or pd.isna(valor) or str(valor).strip() == "" or str(valor).lower() == "nan":
        return "Sin información"
    try:
        partes = str(valor).split('-')
        rbd_candidato = partes[0].strip()
        if rbd_candidato.isdigit():
            return rbd_candidato
        return str(valor).strip()
    except:
        return str(valor).strip()

# --------------------------------------------------
# PROCESO PRINCIPAL DE CARGA
# --------------------------------------------------
def cargar_reemplazos_historicos(conexion, ruta_archivo):
    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ Error: No existe el archivo {ruta_archivo}")
            return

        print(f"🚀 Procesando {ruta_archivo}...")
        df = pd.read_excel(ruta_archivo)

        if df.empty:
            print("⚠️ El archivo Excel no contiene datos.")
            return

        # 🔍 DIAGNÓSTICO: Imprimir las primeras cabeceras reales detectadas antes de limpiar
        print("\n=== DETECCIÓN DE CABECERAS EN EXCEL ===")
        print(f"Columnas encontradas originalmente en tu archivo:\n{list(df.columns)[:8]}... (y más)")
        print("=========================================\n")

        # 🔥 NORMALIZACIÓN AGRESIVA DE LAS COLUMNAS DEL EXCEL
        df.columns = [normalizar_cadena(col) for col in df.columns]

        cursor = conexion.cursor()

        print("🧹 Vaciando tabla y reiniciando contadores...")
        cursor.execute("TRUNCATE TABLE reemplazos")

        # Sentencia SQL con el orden de inserción exacto de tus atributos
        sql_insert = """
            INSERT INTO reemplazos (
                fecha_solicitud, rbd, tipo_de_profesional, solicitud, motivo, 
                nombre_reemplazo, rut_reemplazo, funcion_profesional_reemplazo, cantidad_de_horas, subv_general, 
                horas_sep, horas_pie, horas_vtf, tipo_de_horas, observaciones_solicitud, 
                nombre_reemplazante, rut_reemplazante, telefono_reemplazante, email_reemplazante, funcion_profesional_reemplazante, 
                estado, fecha_visacion_uatp, observacion_uatp, fecha_visacion_gdp, observacion_gdp, 
                fecha_ingreso, licencia_medica, fecha_inicio_de_reposo, fecha_termino_reposo, fecha_inicio_reemplazo, 
                fecha_termino_de_reemplazo, id_simple, nombre_solicitante, rut_solicitante, email_solicitante, telefono_solicitante
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """

        stats = {"insertados": 0, "errores": 0}
        filas_erroneas = []

        print("📥 Cargando datos en la base de datos...")
        for index, row in df.iterrows():
            try:
                # 💥 MAPEO MULTI-ALIAS COMBINADO (Busca nombres cortos y largos de ambas listas)
                fecha_solicitud = limpiar_fecha(obtener_valor(row, 'fechasolicitud', 'Fecha Solicitud', 'Fecha Ingreso Solicitud'))
                rbd = extraer_rbd(obtener_valor(row, 'establecimiento', 'Establecimiento'))
                tipo_de_profesional = limpiar_texto(obtener_valor(row, 'tipodeprofesional', 'Tipo de Profesional'))
                solicitud = limpiar_texto(obtener_valor(row, 'solicitud', 'Solicitud', 'Tipo solicitud'))
                motivo = limpiar_texto(obtener_valor(row, 'motivo', 'Motivo'))
                
                nombre_reemplazo = limpiar_texto(obtener_valor(row, 'nombrereemplazo', 'Nombre Reemplazo', 'Nombre docente aaee a reemplazar', 'Nombre docente o asistente a reemplazar'))
                rut_reemplazo = limpiar_texto(obtener_valor(row, 'rutreemplazo', 'Rut Reemplazo', 'RUT'))
                funcion_profesional_reemplazo = limpiar_texto(obtener_valor(row, 'funcionprofesionalreemplazo', 'Funcion Profesional Reemplazo', 'Función profesional'))
                
                cantidad_de_horas = limpiar_entero(obtener_valor(row, 'cantidaddehoras', 'Cantidad de Horas', 'Cantidad de horas totales'))
                subv_general = limpiar_entero(obtener_valor(row, 'subvgeneral', 'Subv General', 'Subv Genral'))
                horas_sep = limpiar_entero(obtener_valor(row, 'horassep', 'Horas SEP', 'SEP'))
                horas_pie = limpiar_entero(obtener_valor(row, 'horaspie', 'Horas PIE', 'PIE'))
                horas_vtf = limpiar_entero(obtener_valor(row, 'horasvtf', 'Horas VTF', 'VTF'))
                
                tipo_de_horas = limpiar_texto(obtener_valor(row, 'tipodehoras', 'Tipo de Horas'))
                observaciones_solicitud = limpiar_texto(obtener_valor(row, 'observacionessolicitud', 'Observaciones Solicitud'))
                
                nombre_reemplazante = limpiar_texto(obtener_valor(row, 'nombrereemplazante', 'Nombre Reemplazante', 'Nombre completo del reemplazante'))
                rut_reemplazante = limpiar_texto(obtener_valor(row, 'rutreemplazante', 'Rut Reemplazante', 'rut reemplazante', 'RUT2'))
                telefono_reemplazante = limpiar_texto(obtener_valor(row, 'telefonoreemplazante', 'Telefono Reemplazante'))
                email_reemplazante = limpiar_texto(obtener_valor(row, 'emailreemplazante', 'Email Reemplazante'))
                funcion_profesional_reemplazante = limpiar_texto(obtener_valor(row, 'funcionprofesionalreemplazante', 'Funcion Profesional Reemplazante', 'Función profesional2'))
                
                estado = limpiar_texto(obtener_valor(row, 'estado', 'Estado'))
                
                fecha_visacion_uatp = limpiar_fecha(obtener_valor(row, 'fechavisacionuatp', 'Fecha Visacion UATP', 'Fecha Estado UATP'))
                observacion_uatp = limpiar_texto(obtener_valor(row, 'observacionuatp', 'Observacion UATP', 'V°B° UATP'))
                fecha_visacion_gdp = limpiar_fecha(obtener_valor(row, 'fechavisaciongdp', 'Fecha Visacion GDP'))
                observacion_gdp = limpiar_texto(obtener_valor(row, 'observaciongdp', 'Observacion GDP', 'Observaciongdp'))
                
                fecha_ingreso = limpiar_fecha(obtener_valor(row, 'fechaingreso', 'Fecha Ingreso'))
                licencia_medica = limpiar_texto(obtener_valor(row, 'licenciamedica', 'Licencia Medica'))
                fecha_inicio_de_reposo = limpiar_fecha(obtener_valor(row, 'fechainiciodereposo', 'Fecha Inicio Reposo', 'Fecha de Inicio Reposo'))
                fecha_termino_reposo = limpiar_fecha(obtener_valor(row, 'fechaterminoreposo', 'Fecha Termino Reposo', 'Fecha de Termino Reposo'))
                fecha_inicio_reemplazo = limpiar_fecha(obtener_valor(row, 'fechainicioreemplazo', 'Fecha Inicio Reemplazo', 'Fecha inicio'))
                fecha_termino_de_reemplazo = limpiar_fecha(obtener_valor(row, 'fechadeterminodereemplazo', 'Fecha Termino Reemplazo', 'Fecha término', 'Fecha termino'))
                
                id_simple = limpiar_texto(obtener_valor(row, 'idsimple', 'Id Simple'))
                nombre_solicitante = limpiar_texto(obtener_valor(row, 'nombresolicitante', 'Nombre Solicitante'))
                rut_solicitante = limpiar_texto(obtener_valor(row, 'rutsolicitante', 'Rut Solicitante'))
                email_solicitante = limpiar_texto(obtener_valor(row, 'emailsolicitante', 'Email Solicitante'))
                telefono_solicitante = limpiar_texto(obtener_valor(row, 'telefonosolicitante', 'Telefono Solicitante'))

                # Tupla ordenada de manera idéntica al INSERT
                valores = (
                    fecha_solicitud, rbd, tipo_de_profesional, solicitud, motivo,
                    nombre_reemplazo, rut_reemplazo, funcion_profesional_reemplazo, cantidad_de_horas, subv_general,
                    horas_sep, horas_pie, horas_vtf, tipo_de_horas, observaciones_solicitud,
                    nombre_reemplazante, rut_reemplazante, telefono_reemplazante, email_reemplazante, funcion_profesional_reemplazante,
                    estado, fecha_visacion_uatp, observacion_uatp, fecha_visacion_gdp, observacion_gdp,
                    fecha_ingreso, licencia_medica, fecha_inicio_de_reposo, fecha_termino_reposo, fecha_inicio_reemplazo,
                    fecha_termino_de_reemplazo, id_simple, nombre_solicitante, rut_solicitante, email_solicitante, telefono_solicitante
                )

                cursor.execute(sql_insert, valores)
                stats["insertados"] += 1

            except Exception as e:
                stats["errores"] += 1
                mensaje_error = getattr(e, 'msg', str(e))
                print(f"⚠️ Error Fila {index+2}: {mensaje_error}")
                
                fila_con_error = row.to_dict()
                fila_con_error['ERROR_BD_DETALLE'] = mensaje_error
                filas_erroneas.append(fila_con_error)

        conexion.commit()
        print("-" * 50)
        print(f"📊 RESUMEN DE EJECUCIÓN:")
        print(f"   Filas insertadas: {stats['insertados']}")
        print(f"   Filas con error: {stats['errores']}")

        if filas_erroneas:
            ruta_log = 'excel/reemplazos/reemplazos-sin-ingresar.xlsx'
            os.makedirs(os.path.dirname(ruta_log), exist_ok=True)
            df_err = pd.DataFrame(filas_erroneas)
            df_err.to_excel(ruta_log, index=False)
            print(f"📄 Archivo de errores generado en: {ruta_log}")

    except Exception as e:
        print(f"💥 Error crítico general: {e}")
    finally:
        if 'cursor' in locals(): cursor.close()

if __name__ == '__main__':
    ARCHIVO = 'excel/reemplazos/solicituddereemplazos.xlsx' 
    conn = crear_conexion()
    if conn:
        cargar_reemplazos_historicos(conn, ARCHIVO)
        conn.close()