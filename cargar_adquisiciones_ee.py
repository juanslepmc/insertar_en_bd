import mysql.connector
from mysql.connector import Error
import pandas as pd
import os

# --------------------------------------------------
# CONEXIÓN
# --------------------------------------------------
def crear_conexion():
    try:
        return mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password='123456789', # Ajusta si es necesario
            database='bd_institucional_v2'
        )
    except Error as e:
        print(f"❌ Error de conexión: {e}")
        return None

# --------------------------------------------------
# UTILIDADES
# --------------------------------------------------
def limpiar_decimal(valor):
    if pd.isna(valor) or str(valor).strip() == "" or valor == "Sin información":
        return 0.0
    try:
        # Reemplaza comas por puntos si el Excel viene con formato regional
        val_str = str(valor).replace(',', '.')
        return float(val_str)
    except:
        return 0.0

# --------------------------------------------------
# CARGA DE ADQUISICIONES
# --------------------------------------------------
def cargar_adquisicion_completo(conexion, ruta_archivo):
    cursor = conexion.cursor()

    try:
        if not os.path.exists(ruta_archivo):
            print(f"❌ No se encontró el archivo: {ruta_archivo}")
            return

        # Cargamos el nuevo archivo forzando el RBD como texto
        df = pd.read_excel(
            ruta_archivo,
            dtype={'detalle_numero_rbd': str}
        )

        df.columns = df.columns.str.strip().str.lower()
        #df = df.fillna("Sin información")
        # 1. Identificamos qué columnas NO son el RBD
        cols_para_limpiar = [c for c in df.columns if c != 'detalle_numero_rbd']

        # 2. Llenamos solo esas con "Sin información"
        df[cols_para_limpiar] = df[cols_para_limpiar].fillna("Sin información")


        # --- INICIO DE TRANSACCIÓN ---
        print("⚡ Iniciando transacción segura para Adquisiciones...")
        conexion.start_transaction()

        # Limpiamos la tabla de adquisiciones
        cursor.execute("DELETE FROM Adquisicion_por_EE")

        sql = """
            INSERT INTO Adquisicion_por_EE (
                folio, anio, fecha_creacion, fecha_modificacion, nombre,
                descripcion, fecha_esperada_compra, fecha_esperada_recepcion,
                fecha_firma_jefe_compras, monto_total, monto_total_compra,
                tipo_documento_compra, numero_documento_compra, fecha_documento_compra,
                razon_social_proveedor, rut_proveedor, estado,
                nombre_usuario_ejecutivo_compras, nombre_usuario_jefe_compras,
                modalidad_compra, tipo_modalidad_compra, detalle_folio_solicitud_compra,
                detalle_nombre_materia_compra, detalle_nombre_fuente_financiamiento,
                detalle_tipo_gasto_pac, detalle_codigo_concepto_presupuestario,
                detalle_nombre_concepto_presupuestario, detalle_cantidad,
                detalle_valor_unitario_neto, detalle_valor_unitario_compra,
                detalle_total_compra, rbd
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
        """
        
        filas_ok = 0
        filas_omitidas = []

        for index, row in df.iterrows():
            rbd_raw = row.get('detalle_numero_rbd')
            rbd_str = "" if pd.isna(rbd_raw) else str(rbd_raw).strip()

            # ✅ OMISIÓN CONTROLADA (Si el RBD es nulo, vacío o "0")
            if rbd_str == "" or rbd_str == "0" or rbd_str == "nan":
                fila_error = row.to_dict()
                fila_error['fila_excel_origen'] = index + 2
                fila_error['motivo_omision'] = 'RBD ausente o cero (Sin asignación)'
                filas_omitidas.append(fila_error)
                continue

            # Mapeo de valores (Ajustado a los nombres del Excel Adquisiciones)
            valores = (
                str(row.get('folio', '')),
                int(row.get('ano', 0)), # En excel es 'ano', en BD 'anio'
                str(row.get('fecha_creacion', '')),
                str(row.get('fecha_modificacion', '')),
                str(row.get('nombre', '')),
                str(row.get('descripcion', '')),
                str(row.get('fecha_esperada_compra', '')),
                str(row.get('fecha_esperada_recepcion', '')),
                str(row.get('fecha_firma_jefe_compras', '')),
                limpiar_decimal(row.get('monto_total', 0)),
                limpiar_decimal(row.get('monto_total_compra', 0)),
                str(row.get('tipo_documento_compra', '')),
                str(row.get('numero_documento_compra', '')),
                str(row.get('fecha_documento_compra', '')),
                str(row.get('razon_social_proveedor', '')),
                str(row.get('rut_proveedor', '')),
                str(row.get('estado', '')),
                str(row.get('nombre_usuario_ejecutivo_compras', '')),
                str(row.get('nombre_usuario_jefe_compras', '')),
                str(row.get('modalidad_compra', '')),
                str(row.get('tipo_modalidad_compra', '')),
                str(row.get('detalle_folio_solicitud_compra', '')),
                str(row.get('detalle_nombre_materia_compra', '')),
                str(row.get('detalle_nombre_fuente_financiamiento', '')),
                str(row.get('detalle_tipo_gasto_pac', '')),
                str(row.get('detalle_codigo_concepto_presupuestario', '')),
                str(row.get('detalle_nombre_concepto_presupuestario', '')),
                int(row.get('detalle_cantidad', 0)) if str(row.get('detalle_cantidad', '')).isnumeric() else 0,
                limpiar_decimal(row.get('detalle_valor_unitario_neto', 0)),
                limpiar_decimal(row.get('detalle_valor_unitario_compra', 0)),
                limpiar_decimal(row.get('detalle_total_compra', 0)),
                rbd_str
            )

            try:
                cursor.execute(sql, valores)
                filas_ok += 1
            except mysql.connector.Error as e:
                if e.errno == 1452:
                    print(f"⚠️ Error FK: El RBD '{rbd_str}' no existe en Establecimientos | Fila: {index + 2}")
                else:
                    print(f"❌ Error MySQL en fila {index + 2}: {e}")
                raise # Provoca el rollback global

        conexion.commit()

        # Exportar logs de omitidos
        if filas_omitidas:
            df_omitidas = pd.DataFrame(filas_omitidas)
            output = 'excel/adquisiciones_omitidas_sin_rbd.xlsx'
            df_omitidas.to_excel(output, index=False)
            print(f"📄 Archivo de registros omitidos generado: {output}")

        print("-" * 50)
        print(f"✅ CARGA FINALIZADA EXITOSAMENTE")
        print(f"   Registros insertados: {filas_ok}")
        print(f"   Registros omitidos: {len(filas_omitidas)}")

    except Exception as e:
        conexion.rollback()
        print(f"❌ PROCESO CANCELADO: {e}")
        print("🔄 Rollback ejecutado: La tabla Adquisicion_por_EE no sufrió cambios.")

    finally:
        cursor.close()

# --------------------------------------------------
# EJECUCIÓN
# --------------------------------------------------
if __name__ == '__main__':
    archivo_input = 'excel/adquisicion_rbd.xlsx'
    conn = crear_conexion()
    
    if conn:
        cargar_adquisicion_completo(conn, archivo_input)
        conn.close()