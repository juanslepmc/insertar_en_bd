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
            password='123456789', 
            database='bd_institucional_v2'
        )
        return conexion
    except Error as e:
        print(f"❌ Error de conexión: {e}")
        return None

def extraer_datos_fila(row, columns_map):
    """
    Extrae y normaliza los 8 campos dinámicamente sin importar si las 
    columnas están en mayúsculas, minúsculas o divididas (Run + DV).
    """
    col_map = {str(c).strip().lower(): c for c in columns_map}

    # 1. RUT (Directo o combinado Run + DV)
    rut_final = "Sin información"
    if 'rut' in col_map and pd.notna(row[col_map['rut']]):
        r_val = str(row[col_map['rut']]).strip()
        rut_final = r_val[:-2] if r_val.endswith('.0') else r_val
    elif 'run' in col_map and pd.notna(row[col_map['run']]):
        run_str = str(row[col_map['run']]).strip().replace('.', '').replace(',', '')
        if run_str.endswith('.0'): run_str = run_str[:-2]
        dv_col = col_map.get('dígito ver.') or col_map.get('digito ver') or col_map.get('dv')
        dv_str = str(row[dv_col]).strip().upper() if dv_col and pd.notna(row[dv_col]) else ''
        if dv_str.endswith('.0'): dv_str = dv_str[:-2]
        rut_final = f"{run_str}-{dv_str}" if dv_str else run_str

    # Helper para textos
    def get_text(keys, fmt='title'):
        for k in keys:
            if k in col_map and pd.notna(row[col_map[k]]):
                val = str(row[col_map[k]]).strip()
                if val not in ['', 'nan', 'None']:
                    if fmt == 'title': return val.title()
                    if fmt == 'capitalize': return val.capitalize()
                    return val
        return "Sin información"

    # 2. Nombres, Apellidos, Comuna y Estado
    nombres = get_text(['nombres', 'nombre'], 'title')
    ape_pat = get_text(['apellido_paterno', 'apellido paterno', 'paterno'], 'title')
    ape_mat = get_text(['apellido_materno', 'apellido materno', 'materno'], 'title')
    if ape_mat == "Sin información": 
        ape_mat = None
        
    comuna = get_text(['comuna', 'comuna residencia', 'comuna_residencia'], 'capitalize')
    estado = get_text(['estado'], 'raw')
    if estado == "Sin información": 
        estado = "Activo"

    # 3. Género
    genero = None
    for k in ['genero', 'género', 'sexo']:
        if k in col_map and pd.notna(row[col_map[k]]):
            v_raw = str(row[col_map[k]]).strip().upper()
            if v_raw in ['F', 'FEMENINO']: 
                genero = 'Femenino'
            elif v_raw in ['M', 'MASCULINO']: 
                genero = 'Masculino'
            elif v_raw not in ['', 'NAN', 'NONE']: 
                genero = v_raw.title()
            break

    # 4. Fecha de Nacimiento
    fecha_nac = None
    for k in ['fecha_nacimiento', 'fecha nacimiento', 'fecha_nac']:
        if k in col_map and pd.notna(row[col_map[k]]):
            raw_f = str(row[col_map[k]]).strip()
            if raw_f not in ['', 'nan', 'None']:
                dt = pd.to_datetime(raw_f, errors='coerce')
                fecha_nac = dt.strftime('%Y-%m-%d') if pd.notna(dt) else raw_f
            break

    return (rut_final, nombres, ape_pat, ape_mat, comuna, estado, genero, fecha_nac)

def cargar_alumnos(conexion, ruta_archivo):
    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ Error: No existe el archivo en: {ruta_archivo}")
            return

        print(f"🚀 Procesando archivo: {ruta_archivo}...")
        df = pd.read_excel(ruta_archivo)
        cursor = conexion.cursor()

        # SQL alineado con las 8 columnas del modelo
        sql = """
            INSERT INTO Alumnos 
            (rut, nombres, apellido_paterno, apellido_materno, comuna, estado, genero, fecha_nacimiento) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                nombres = VALUES(nombres),
                apellido_paterno = VALUES(apellido_paterno),
                apellido_materno = VALUES(apellido_materno),
                comuna = VALUES(comuna),
                estado = VALUES(estado),
                genero = VALUES(genero),
                fecha_nacimiento = VALUES(fecha_nacimiento)
        """

        stats = {"nuevos": 0, "actualizados": 0, "errores": 0}

        for index, row in df.iterrows():
            fila_num = index + 2
            valores = extraer_datos_fila(row, df.columns)

            try:
                cursor.execute(sql, valores)
                if cursor.rowcount == 1:
                    stats["nuevos"] += 1
                elif cursor.rowcount == 2:
                    stats["actualizados"] += 1
            except mysql.connector.Error as e:
                stats["errores"] += 1
                print(f"⚠️ Error Fila {fila_num} (RUT: {valores[0]}): {e.msg}")

        conexion.commit()
        print("-" * 50)
        print("RESUMEN DE CARGA EN BASE DE DATOS:")
        print(f"✅ Registros nuevos: {stats['nuevos']}")
        print(f"🔄 Registros actualizados: {stats['actualizados']}")
        print(f"❌ Errores encontrados: {stats['errores']}")

    except Exception as e:
        print(f"💥 Error crítico durante la carga: {e}")
    finally:
        if 'cursor' in locals():
            cursor.close()

if __name__ == '__main__':
    ruta_archivo = 'excel/alumnos/cargar/cargar_generico_2.xlsx'
    
    conn = crear_conexion()
    if conn:
        cargar_alumnos(conn, ruta_archivo)
        conn.close()