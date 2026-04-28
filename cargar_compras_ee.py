import mysql.connector
from mysql.connector import Error
import pandas as pd
import os

# --------------------------------------------------
# CONEXIÓN
# --------------------------------------------------
def crear_conexion():
    return mysql.connector.connect(
        host='localhost',
        port=3306,
        user='root',
        password='123456789',
        database='bd_institucional_v2'
    )

# --------------------------------------------------
# UTILIDADES
# --------------------------------------------------
def limpiar_decimal(valor):
    if pd.isna(valor) or str(valor).strip() == "" or valor == "Sin información":
        return 0.0
    try:
        return float(valor)
    except:
        return 0.0

# --------------------------------------------------
# CARGA DE COMPRAS
# --------------------------------------------------
def cargar_compras_completo(conexion, ruta_archivo):
    cursor = conexion.cursor()

    try:
        df = pd.read_excel(
            ruta_archivo,
            dtype={'detalle_numero_rbd': str}
        )

        df.columns = df.columns.str.strip().str.lower()

        conexion.start_transaction()
        cursor.execute("DELETE FROM Compras_por_EE")

        sql = """
            INSERT INTO Compras_por_EE (
                folio, anio, fecha_creacion, fecha_modificacion, nombre,
                descripcion, fundamento, informacion_extra, monto_estimado,
                tipo_usuario_requirente, nombre_usuario_requirente, unidad_requirente,
                estado, materia_compra, fuente_financiamiento, tipo_gasto_pac,
                fecha_esperada_compra, nombre_usuario_autorizador_rbd,
                fecha_aprob_autorizador_rbd, nombre_usuario_ejecutivo_compras,
                fecha_aprobacion, fecha_finalizacion, monto_total_compra,
                detalle_nombre, detalle_concepto_presupuesto,
                detalle_concepto_presupuestario_codigo, detalle_total_compra,
                detalle_fecha_aprobacion, rbd
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s
            )
        """
        filas_procesadas = 0

        filas_ok = 0
        filas_omitidas = []

        for index, row in df.iterrows():
            rbd = row.get('detalle_numero_rbd')

            rbd_str = "" if pd.isna(rbd) else str(rbd).strip()


            # ✅ OMISIÓN CONTROLADA (NULL, vacío o 0)
            if rbd_str == "" or rbd_str == "0":
                fila_error = row.to_dict()
                fila_error['fila_excel_origen'] = index + 2
                fila_error['motivo_omision'] = 'RBD nulo / 0 (solicitud borrador)'
                filas_omitidas.append(fila_error)
                continue


            #rbd = rbd.strip()

            valores = (
                str(row.get('folio', '')),
                int(row.get('ano', 0)),
                str(row.get('fecha_creacion', '')),
                str(row.get('fecha_modificacion', '')),
                str(row.get('nombre', '')),
                str(row.get('descripcion', '')),
                str(row.get('fundamento', '')),
                str(row.get('informacion_extra', '')),
                limpiar_decimal(row.get('monto_estimado', 0)),
                str(row.get('tipo_usuario_requirente', '')),
                str(row.get('nombre_usuario_requirente', '')),
                str(row.get('unidad_requirente', '')),
                str(row.get('estado', '')),
                str(row.get('materia_compra', '')),
                str(row.get('fuente_financiamiento', '')),
                str(row.get('tipo_gasto_pac', '')),
                str(row.get('fecha_esperada_compra', '')),
                str(row.get('nombre_usuario_autorizador_rbd', '')),
                str(row.get('fecha_aprob_autorizador_rbd', '')),
                str(row.get('nombre_usuario_ejecutivo_compras', '')),
                str(row.get('fecha_aprobacion', '')),
                str(row.get('fecha_finalizacion', '')),
                limpiar_decimal(row.get('monto_total_compra', 0)),
                str(row.get('detalle_nombre', '')),
                str(row.get('detalle_concepto_presupuestario', '')),
                str(row.get('detalle_concepto_presupuestario_codigo', '')),
                limpiar_decimal(row.get('detalle_total_compra', 0)),
                str(row.get('detalle_fecha_aprobacion', '')),
                rbd
            )

            try:
                cursor.execute(sql, valores)
                filas_ok += 1
                filas_procesadas += 1

            except mysql.connector.Error as e:
                if e.errno == 1452:
                    print(
                        f"⚠️ Error FK: El RBD '{rbd}' no existe en Establecimientos | "
                        f"Fila Excel: {index + 2}"
                    )
                else:
                    print(
                        f"❌ Error MySQL en fila {index + 2}: {e}"
                    )
                raise  # fuerza rollback
        conexion.commit()

        # --------------------------------------------------
        # EXPORTAR OMITIDAS
        # --------------------------------------------------
        if filas_omitidas:
            df_omitidas = pd.DataFrame(filas_omitidas)
            output = 'excel/compras_omitidas_sin_rbd.xlsx'
            df_omitidas.to_excel(output, index=False)
            print(f"📄 Archivo generado: {output}")

        print(f"✅ Registros cargados correctamente: {filas_ok}")
        print(f"⚠️ Registros omitidos (borradores): {len(filas_omitidas)}")

    except Exception as e:
        conexion.rollback()
        print(f"❌ Error: {e}")
        print("🔄 Rollback ejecutado")

    finally:
        cursor.close()

# --------------------------------------------------
# MAIN
# --------------------------------------------------
if __name__ == '__main__':
    archivo = 'excel/solicitud_compra_rbd_completa.xlsx'
    conn = crear_conexion()
    if conn:
        cargar_compras_completo(conn, archivo)
        conn.close()