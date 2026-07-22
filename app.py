import pandas as pd
import re

# 1. Cargar el archivo inicial subido
archivo_inicial = 'CARGA PALERMO 2026 (1).xlsx'
xls = pd.ExcelFile(archivo_inicial)

registros_procesados = []

# 2. Recorrer las pestañas de categorías
for categoria in xls.sheet_names:
    df_hoja = pd.read_excel(archivo_inicial, sheet_name=categoria)
    
    for _, fila in df_hoja.iterrows():
        codigo = str(fila.iloc[0]).strip() if pd.notna(fila.iloc[0]) else ""
        descripcion = str(fila.iloc[1]).strip() if pd.notna(fila.iloc[1]) else ""
        
        # Validar que sea un código de producto válido (ej. AC1180, T1895, etc.)
        if re.match(r'^[A-Z]{1,3}\d+', codigo, re.IGNORECASE):
            
            # --- CONCATENACIÓN EN EL ARCHIVO INICIAL ---
            descripcion_concatenada = f"{codigo} - {descripcion}"
            
            costo = fila.iloc[2] if len(fila) > 2 and pd.notna(fila.iloc[2]) else 0
            precio_tc = fila.iloc[3] if len(fila) > 3 and pd.notna(fila.iloc[3]) else 0
            
            # Guardar el producto con la descripción ya concatenada
            registros_procesados.append({
                "Categoría": categoria,
                "Código Origen": codigo,
                "Descripción Origen": descripcion,
                "Descripción Completa (Concatenada)": descripcion_concatenada,
                "Costo": costo,
                "Precio Venta": precio_tc
            })

# 3. Crear el DataFrame resultante y exportar a Excel
df_final = pd.DataFrame(registros_procesados)
df_final.to_excel("Resultado_Concatenacion_Inicial.xlsx", index=False)

print(f"Procesados {len(df_final)} productos correctamente.")