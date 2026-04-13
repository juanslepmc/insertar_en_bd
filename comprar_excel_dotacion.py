import pandas as pd

# 1. Definir los nombres de los archivos
archivo_1 = 'liquidaciones-netcore-marzo2026.xlsx'
archivo_2 = 'NominaFuncionario.xlsx'
archivo_salida = 'dotacion_funcionarios.xlsx'

# 2. Cargar los archivos Excel
print("Cargando archivos...")
df_liquidacion = pd.read_excel(archivo_1)
df_nomina = pd.read_excel(archivo_2)

# Aseguramos que la columna RUN sea tratada como texto para evitar errores de formato
df_liquidacion['RUN'] = df_liquidacion['RUN'].astype(str).str.strip()
df_nomina['RUN'] = df_nomina['RUN'].astype(str).str.strip()

# 3. Realizar el MATCH (Inner Join)
# Solo los que están en AMBOS archivos
df_match = pd.merge(df_liquidacion, df_nomina, on='RUN', how='inner')

# 4. Realizar el NO MATCH
# Usamos un 'outer join' con un indicador para saber dónde falta el dato
df_all = pd.merge(df_liquidacion, df_nomina, on='RUN', how='outer', indicator=True)

# Filtramos los que NO están en ambos (solo en izquierda o solo en derecha)
df_no_match = df_all[df_all['_merge'] != 'both']

# Opcional: Eliminar la columna auxiliar '_merge' para que el Excel quede limpio
df_no_match = df_no_match.drop(columns=['_merge'])

# 5. Generar el archivo final con dos pestañas
print(f"Generando archivo: {archivo_salida}...")
with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
    df_match.to_excel(writer, sheet_name='Match', index=False)
    df_no_match.to_excel(writer, sheet_name='No Match', index=False)

print("¡Proceso completado con éxito!")