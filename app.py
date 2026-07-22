import pandas as pd
import re

def procesar_archivo_inicial(path_archivo_inicial, path_salida="Matriz_Carga_Web.xlsx"):
    xls = pd.ExcelFile(path_archivo_inicial)
    productos_procesados = []
    codigo_padre_counter = 503500

    for hoja in xls.sheet_names:
        df = pd.read_excel(path_archivo_inicial, sheet_name=hoja)
        if df.empty:
            continue

        # Detectar la fila de encabezados de color/talle (suele estar en las primeras 3 filas)
        header_idx = None
        for i in range(min(5, len(df))):
            row_vals = [str(x).upper() for x in df.iloc[i].values]
            if "COSTO" in row_vals or "TARJETA" in row_vals or "TC" in row_vals:
                header_idx = i
                break

        if header_idx is None:
            continue

        headers = df.iloc[header_idx].values
        df_datos = df.iloc[header_idx + 1 :].reset_index(drop=True)

        for _, fila in df_datos.iterrows():
            cod_raw = str(fila.iloc[0]).strip() if pd.notna(fila.iloc[0]) else ""
            desc_raw = str(fila.iloc[1]).strip() if pd.notna(fila.iloc[1]) else ""

            # Filtramos filas que no sean productos (fechas, vacíos, notas)
            if not cod_raw or cod_raw.upper() in ["REPO", "FECHA", "NAN"] or "FECHA" in desc_raw.upper():
                continue

            # Validar patrón de código (ej. AC1180, T1895, B600, SH353, etc.)
            if re.match(r"^[A-Z]{1,3}\d+", cod_raw, re.IGNORECASE):
                
                # ========================================================
                # 1. CONCATENACIÓN EN EL ARCHIVO INICIAL (Código + Nombre)
                # ========================================================
                nombre_concatenado_inicial = f"{cod_raw} {desc_raw}"

                costo = fila.iloc[2] if len(fila) > 2 and pd.notna(fila.iloc[2]) else 0
                precio = fila.iloc[3] if len(fila) > 3 and pd.notna(fila.iloc[3]) else 0

                # Desglose de variantes (Talle / Color / Stock)
                col_idx = 5
                mientras_variantes = True
                variante_num = 1

                while col_idx < len(fila):
                    col_header = str(headers[col_idx]).strip() if pd.notna(headers[col_idx]) else ""
                    val_talle_col = fila.iloc[col_idx - 1] if col_idx - 1 >= 5 else "U"
                    val_stock = fila.iloc[col_idx]

                    # Si hay stock registrado
                    if pd.notna(val_stock) and str(val_stock).strip().isdigit() and int(val_stock) > 0:
                        talle = str(val_talle_col).strip() if pd.notna(val_talle_col) else "U"
                        color = col_header if col_header and "UNNAMED" not in col_header.upper() else "ÚNICO"

                        codigo_padre_str = f"{codigo_padre_counter} {variante_num:03d}"

                        # ========================================================
                        # 2. SE GUARDA EN LA MATRIZ USA LA CONCATENACIÓN PREVIA
                        # ========================================================
                        productos_procesados.append({
                            "Codigo padre": codigo_padre_str,
                            "Hijo": None,
                            "Descripción / Nombre": nombre_concatenado_inicial,  # <--- Concatenado desde el origen
                            "Categoría": hoja.upper(),
                            "Proveedor": "1704",
                            "Costo": costo,
                            "Precio": precio,
                            "Talle": talle,
                            "Color": color.upper(),
                            "Stock": int(val_stock),
                            "Año": "1704"
                        })
                        variante_num += 1

                    col_idx += 2

                codigo_padre_counter += 1

    df_matriz = pd.DataFrame(productos_procesados)
    df_matriz.to_excel(path_salida, index=False)
    print(f" Proceso finalizado con éxito: {len(df_matriz)} filas generadas en '{path_salida}'.")
    return df_matriz

# Ejemplo de ejecución:
# df_resultado = procesar_archivo_inicial("CARGA PALERMO 2026 (1).xlsx")