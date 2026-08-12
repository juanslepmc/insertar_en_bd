import mysql.connector
from mysql.connector import Error
import pandas as pd
import os
from datetime import datetime

# --------------------------------------------------
# DICCIONARIOS DE DATOS (CLAVE - VALOR)
# --------------------------------------------------
MESES_MAPA = {
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4,
    'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8,
    'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
}

# --------------------------------------------------
# DICCIONARIO DE MAPEO ESTABLECIMIENTOS -> RBD
# --------------------------------------------------
DICCIONARIO_RBD = {
    'Centro Educ. Constitución -Maestra M. Fresia Hormazábal L.': 3170,
    'Centro Educacional Constitución': 3170,
    'Colegio Blanco Encalada': 3544,
    'Escuela Adolfo Quiroz Hernández': 3552,
    'Escuela Antonieta Leon Verdugo': 3632,
    'Escuela Antonieta León Verdugo': 3632,
    'Escuela Aníbal Pinto': 3541,
    'Escuela Artística Barrio Estación': 3545,
    'Escuela Barranquillas': 3179,
    'Escuela Benito Mancilla Pérez': 3616,
    'Escuela Blanca Bustos Castillo': 3611,
    'Escuela Cabrería': 3604,
    'Escuela Carreras Cortas': 3623,
    'Escuela Carrizalillo': 3187,
    'Escuela Cerro Alto Jose Opazo Diaz': 3169,
    'Escuela Cerro Alto José Opazo Diaz': 3169,
    'Escuela Chacarillas': 16506,
    'Escuela Chorrillos': 3586,
    'Escuela Clorindo Alvear': 3554,
    'Escuela Costa Blanca': 3185,
    'Escuela Ecológica Rosita O\'Higgins': 3548,
    'Escuela Eduardo Machado Corsi': 3198,
    'Escuela El Sauce': 3629,
    'Escuela El Trozo': 3602,
    'Escuela Ema Vasquez Gutierrez': 3631,
    'Escuela Ema Vásquez Gutiérrez': 3631,
    'Escuela Enrique Donn Muller': 3166,
    'Escuela Enrique Donn Müller': 3166,
    'Escuela Escritora Marcela Paz': 3615,
    'Escuela Especial De Lenguaje Polhuín': 20342,
    'Escuela Especial Horizonte': 3537,
    'Escuela Especial de Lenguaje Polhuin': 20342,
    'Escuela Ester Urrutia De Ruiz': 3588,
    'Escuela Ester Urrutia de Ruiz': 3588,
    'Escuela Gabriela Mistral': 3625,
    'Escuela Gladys Canales Paredes': 3610,
    'Escuela Héctor Silvestre Paiva Manríquez': 3564,
    'Escuela Independencia': 3546,
    'Escuela Javiera Carrera': 3569,
    'Escuela Jose Rivas Hernandez': 3614,
    'Escuela José Benito Mancilla': 3616,
    'Escuela José Dolores Muñoz Labra - Ex Rincón de Pilen': 3579,
    'Escuela José Rivas Hernández de Cardonal': 3614,
    'Escuela Junquillar': 3176,
    'Escuela La Capilla De Pilen Alto': 3596,
    'Escuela La Capilla de Pilen Alto': 3596,
    'Escuela La Patagua': 3601,
    'Escuela Las Corrientes': 3174,
    'Escuela Leontina Letelier': 3196,
    'Escuela Loanco': 3636,
    'Escuela Los Conquistadores': 3551,
    'Escuela Los Héroes': 3619,
    'Escuela Los Peumos': 3624,
    'Escuela Maria Inés Maromillas': 3177,
    'Escuela Maria Olga Vega': 3134,
    'Escuela Mariano Latorre': 3535,
    'Escuela María Inés Maromillas': 3177,
    'Escuela María Olga Vega Vega': 3583,
    'Escuela Miguel Faundez Morales': 3186,
    'Escuela Miguel Faúndez Morales': 3186,
    'Escuela Mixta Atenea': 3599,
    'Escuela Octavio Palma Pérez': 3555,
    'Escuela Pahuil': 3622,
    'Escuela Pedernales': 3580,
    'Escuela Pedro Antonio Tejos': 3130,
    'Escuela Pedro Antonio Tejos Tejos': 3130,
    'Escuela Pedro De Valdivia': 3573,
    'Escuela Pedro de Valdivia': 3573,
    'Escuela Penitenciaria Las Dunas': 16454,
    'Escuela Penitenciaria Mariano Latorre': 3535,
    'Escuela Porongo': 3549,
    'Escuela Purísima Concepción De Pocillas': 3550,
    'Escuela Purísima Concepción de Pocillas': 3550,
    'Escuela Quiñipato': 3633,
    'Escuela Reloca': 3626,
    'Escuela Ricardo Salgado': 3630,
    'Escuela Rincón De Pilen': 3579,
    'Escuela Rural Quebrada Verde': 3190,
    'Escuela San Alfonso De Canelillo': 3612,
    'Escuela San Alfonso de Canelillo': 3612,
    'Escuela San Ambrosio': 3621,
    'Escuela Santa Aurora De Carrizal': 3191,
    'Escuela Santa Aurora de Carrizal': 3191,
    'Escuela Superior Nueva Bilbao': 3168,
    'Escuela Teresa Consuelo': 3189,
    'Jardín Infantil Mi Pequeño Mundo': 7104009,
    'Jardín Infantil y Sala Cuna Caracolitos': 7102014,
    'Jardín Infantil y Sala Cuna Personitas De Santa Olga': 7102010,
    'Liceo Antonio Varas': 3538,
    'Liceo Bicentenario De Cauquenes': 16751,
    'Liceo Bicentenario de Cauquenes': 16751,
    'Liceo Bicentenario de Excelencia Técnico Profesional': 3172,
    'Liceo Claudina Urrutia De Lavín': 3539,
    'Liceo Constitución': 3165,
    'Liceo Federico Albert Faupp': 3618,
    'Liceo Pelluhue': 3609,
    'Liceo Politécnico Pedro Aguirre Cerda': 3540,
    'Liceo Rural Enrique Mac Iver': 3173,
    'Liceo San Ignacio': 3128,
    'Liceo de Anticipación Claudina Urrutia de Lavín': 3539,
    'Sala Cuna Jardin Infantil Carita de Angel': 7202003,
    'Sala Cuna Jardín Carita de Ángel': 7202003,
    'Sala Cuna Jardín Infantil Abejita Dul': 7102015,
    'Sala Cuna Jardín Infantil Antonio Varas': 7201013,
    'Sala Cuna Jardín Infantil Bellavista': 7201005,
    'Sala Cuna Jardín Infantil Caracolitos': 7102014,
    'Sala Cuna Jardín Infantil Claro de Luna': 7201009,
    'Sala Cuna Jardín Infantil La Casita en el Bosque': 7203001,
    'Sala Cuna Jardín Infantil Los Grillitos de Porongo': 7201011,
    'Sala Cuna Jardín Infantil Lucerito de Esperanza': 7201010,
    'Sala Cuna Jardín Infantil Mi Pequeño Mundo': 7104009,
    'Sala Cuna Jardín Infantil Personitas': 7102010,
    'Sala Cuna Jardín Infantil Sol de Esperanza': 7203003,
    'Sala Cuna y Jardín Infantil Antonio Varas': 7201013,
    'Sala Cuna y Jardín Infantil Bellavista': 7201005,
    'Sala Cuna y Jardín Infantil Claro De Luna': 7201009,
    'Sala Cuna y Jardín Infantil La Casita En El Bosque': 7203001,
    'Sala Cuna y Jardín Infantil Los Grillitos': 7201011,
    'Sala Cuna y Jardín Infantil Lucerito De Esperanza': 7201010,
    'Sala Cuna y Jardín Infantil Sol De Esperanza': 7203003
}

# --------------------------------------------------
# FUNCIONES DE BÚSQUEDA Y LIMPIEZA
# --------------------------------------------------
def obtener_rbd(nombre_centro):
    if pd.isna(nombre_centro) or str(nombre_centro).strip() == "":
        return 0
    nombre_limpio = str(nombre_centro).strip()
    return DICCIONARIO_RBD.get(nombre_limpio, 0)

def obtener_mes_num(mes_texto):
    if pd.isna(mes_texto) or str(mes_texto).strip() == "":
        return 0
    return MESES_MAPA.get(str(mes_texto).strip().lower(), 0)

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

def procesar_rut(rut_sucio):
    if rut_sucio == "Sin información" or pd.isna(rut_sucio): 
        return rut_sucio
    return str(rut_sucio).replace('.', '').strip()

def limpiar_decimal(valor):
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "sin información":
        return 0.0
    try:
        return float(valor)
    except (ValueError, TypeError):
        return 0.0

def limpiar_fecha(valor):
    """
    Estandariza cualquier fecha al formato string 'YYYY-MM-DD', 
    preservando valores cero ('00') originales provenientes de la fuente.
    
    Ejemplos:
      '00-01-1900' -> '1900-01-00'
      '01-03-2023' -> '2023-03-01'
      '2023-03-01' -> '2023-03-01'
      '1900-01-00' -> '1900-01-00'
    """
    # 1. Filtro de nulos o vacíos reales
    if valor is None or pd.isna(valor):
        return "Sin información"


    
    txt = str(valor).strip().split(" ")[0] # Quita la hora si viene incluida
    txt = txt.replace('/', '-')            # Normaliza separadores a '-'
    
    # 2. Control de cadenas vacías o explícitamente nulas
    if txt.lower() in ["", "nan", "none", "null", "sin información"]:
        return "Sin información"
    
    partes = txt.split('-')
    
    # Debe tener 3 partes (día, mes, año o año, mes, día)
    if len(partes) != 3:
        return "Sin información"
    
    p1, p2, p3 = partes[0].strip(), partes[1].strip(), partes[2].strip()
    
    # Caso A: Formato Latino 'DD-MM-YYYY' (El año viene al final, 4 dígitos)
    # Ej: '00-01-1900' -> p1='00', p2='01', p3='1900'
    if len(p3) == 4 and len(p1) <= 2:
        dia = p1.zfill(2)
        mes = p2.zfill(2)
        anio = p3
        return f"{anio}-{mes}-{dia}"
        
    # Caso B: Formato ISO 'YYYY-MM-DD' (El año viene al inicio, 4 dígitos)
    # Ej: '1900-01-00' -> p1='1900', p2='01', p3='00'
    elif len(p1) == 4 and len(p3) <= 2:
        anio = p1
        mes = p2.zfill(2)
        dia = p3.zfill(2)
        return f"{anio}-{mes}-{dia}"
    
    # Si no cumple ninguna de las estructuras de fecha esperadas
    return "Sin información"
# --------------------------------------------------
# PROCESO PRINCIPAL DE CARGA
# --------------------------------------------------
def cargar_dotacion_en_ee(conexion, ruta_archivo, anio_proceso):
    cursor = conexion.cursor()

    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ Error: No existe el archivo de dotación {ruta_archivo}")
            return

        print("Leyendo Excel de dotación y normalizando datos...")
        
        df = pd.read_excel(ruta_archivo, sheet_name='BDMensual') 
        df.columns = df.columns.str.strip().str.lower()

        sql = """
            INSERT INTO dotacion_en_ee 
            (rut, rbd, horas_contratadas, anio, mes, subvencion, tipo_contrato, tipo_hora, fecha_inicio, fecha_termino) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                horas_contratadas = VALUES(horas_contratadas),
                subvencion = VALUES(subvencion),
                tipo_contrato = VALUES(tipo_contrato),
                tipo_hora = VALUES(tipo_hora),
                fecha_inicio = VALUES(fecha_inicio),
                fecha_termino = VALUES(fecha_termino)
        """

        filas_ok = 0
        filas_omitidas = []

        for index, row in df.iterrows():
            fila_excel = index + 2  
            
            nombre_centro = row.get('centro de costo', '')
            rbd_val = obtener_rbd(nombre_centro)
            
            mes_texto = row.get('periodo', '')
            mes_num = obtener_mes_num(mes_texto)

            rut_limpio = procesar_rut(row.get('r.u.n.', row.get('rut', '')))
            horas_contratadas = limpiar_decimal(row.get('n° horas', 0))
            subvencion = str(row.get('subvencion', 'Sin información')).strip().capitalize()
            tipo_contrato = str(row.get('tipo contrato', 'Sin información')).strip().capitalize()
            tipo_hora = str(row.get('tipo hora', 'Sin información')).strip().capitalize()

            if tipo_hora.lower() == 'nan':
                tipo_hora = 'Sin información'

            # --- NUEVA LÓGICA: LECTURA Y FORMATO DE FECHAS ---
            fecha_inicio_raw = row.get('inicio', None)
            fecha_termino_raw = row.get('término', row.get('termino', None))

            fecha_inicio = limpiar_fecha(fecha_inicio_raw)
            fecha_termino = limpiar_fecha(fecha_termino_raw)

            # --- VALIDACIONES ---
            motivos = []
            if not rut_limpio or rut_limpio == "Sin información":
                motivos.append("RUT nulo, vacío o no válido")
            if rbd_val == 0:
                motivos.append(f"RBD no encontrado para el Centro de Costo: '{nombre_centro}'")
            if mes_num == 0:
                motivos.append(f"Mes no reconocido: '{mes_texto}'")

            if motivos:
                fila_error = row.to_dict()
                fila_error['fila_excel_origen'] = fila_excel
                fila_error['motivo_omision'] = " | ".join(motivos)
                filas_omitidas.append(fila_error)
                continue

            valores = (
                rut_limpio,
                rbd_val,
                horas_contratadas,
                anio_proceso,
                mes_num,
                subvencion,
                tipo_contrato,
                tipo_hora,
                fecha_inicio,
                fecha_termino
            )

            # --- INSERCIÓN MYSQL ---
            try:
                cursor.execute(sql, valores)
                filas_ok += 1

            except mysql.connector.Error as e:
                if e.errno == 1452:
                    motivo = f"Error FK (1452): El RUT '{rut_limpio}' o el RBD '{rbd_val}' no existe en las tablas maestras."
                elif e.errno in (1406, 1265):
                    motivo = f"Error Cadena (1406/1265): El texto excede el largo del campo VARCHAR asignado."
                elif e.errno == 1048:
                    motivo = f"Error NOT NULL (1048): Campo obligatorio es nulo en la BD."
                else:
                    motivo = f"Error MySQL [{e.errno}]: {e.msg}"

                fila_error = row.to_dict()
                fila_error['fila_excel_origen'] = fila_excel
                fila_error['motivo_omision'] = motivo
                filas_omitidas.append(fila_error)

        conexion.commit()

        # --- EXPORTACIÓN DE OMITIDOS (2 PESTAÑAS) ---
        if filas_omitidas:
            df_omitidas = pd.DataFrame(filas_omitidas)
            
            rut_col = next((c for c in ['r.u.n.', 'rut', 'run'] if c in df_omitidas.columns), 'r.u.n.')
            nombre_col = next((c for c in ['nombre', 'nombre_funcionario', 'funcionario'] if c in df_omitidas.columns), 'nombre')
            periodo_col = next((c for c in ['periodo', 'mes'] if c in df_omitidas.columns), 'periodo')

            df_omitidas['periodo_norm'] = df_omitidas[periodo_col].astype(str).str.strip().str.capitalize()
            df_omitidas['mes_num'] = df_omitidas[periodo_col].astype(str).str.strip().str.lower().map(MESES_MAPA).fillna(0).astype(int)

            # --- PESTAÑA 1: DETALLE DE OMISIONES ---
            cols_detalle = ['fila_excel_origen', 'motivo_omision'] + [
                c for c in df_omitidas.columns if c not in ['fila_excel_origen', 'motivo_omision', 'periodo_norm', 'mes_num']
            ]
            df_detalle = df_omitidas[cols_detalle].copy()

            # --- PESTAÑA 2: ÚLTIMO MES REGISTRADO POR FUNCIONARIO (RUT ÚNICO) ---
            df_sorted = df_omitidas.sort_values(by=[rut_col, 'mes_num'])
            group_cols = [c for c in [rut_col, nombre_col] if c in df_omitidas.columns]
            
            df_resumen_rut = df_sorted.groupby(group_cols).agg(
                ultimo_mes_registrado=('periodo_norm', 'last'),
                ultimo_mes_num=('mes_num', 'last'),
                total_meses_omitido=('mes_num', 'count'),
                ultimo_motivo_omision=('motivo_omision', 'last')
            ).reset_index()

            df_resumen_rut = df_resumen_rut.sort_values(by=['ultimo_mes_num', rut_col], ascending=[False, True])

            # Exportar a Excel
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output = f'excel/dotacion/dotacion_en_ee/dotacion_omitidas_{timestamp}.xlsx'
            os.makedirs(os.path.dirname(output), exist_ok=True)

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_detalle.to_excel(writer, sheet_name='Detalle_Omitidos', index=False)
                df_resumen_rut.to_excel(writer, sheet_name='Resumen_Ultimo_Mes_RUT', index=False)

            print(f"📄 Archivo consolidado de omitidos generado correctamente: {output}")
            print(f"👥 Total de funcionarios únicos (RUTs) omitidos: {len(df_resumen_rut)}")

        print("-" * 50)
        print("RESULTADO DE CARGA:")
        print(f"✅ Registros cargados/actualizados correctamente: {filas_ok}")
        print(f"⚠️ Filas omitidas/rechazadas en total: {len(filas_omitidas)}")

    except Exception as e:
        conexion.rollback()
        print(f"❌ Error crítico en la carga: {e}")
        print("🔄 Rollback ejecutado")

    finally:
        if 'cursor' in locals():
            cursor.close()

if __name__ == '__main__':
    archivo_dotacion = 'excel/dotacion/dotacion_en_ee/TD Netcore.xlsx'
    anio_proceso = 2026

    conn = crear_conexion()
    if conn:
        cargar_dotacion_en_ee(conn, archivo_dotacion, anio_proceso)
        conn.close()