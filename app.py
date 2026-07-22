import os
import openpyxl
import pandas as pd


def procesar_carga_palermo(
    path_carga_palermo: str,
    path_modelo: str = "Modelo 2 palermo 1407.xlsx",
    path_salida: str = "Carga_Web_Final.xlsx",
    proveedor_id: str = "1407",
):
    """Procesa el archivo 'Carga palermo' filtrando celdas amarillas,

    realiza la concatenación A + B en la columna C y genera la matriz
    con la estructura exacta del 'Modelo 2'.
    """
    if not os.path.exists(path_carga_palermo):
        print(f"❌ Error: El archivo '{path_carga_palermo}' no existe.")
        return

    # 1. Cargar libro con openpyxl para detectar relleno amarillo
    wb = openpyxl.load_workbook(path_carga_palermo, data_only=True)
    filas_nuevas = []

    print("🔍 Buscando artículos nuevos (relleno amarillo)...")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        for r in range(1, ws.max_row + 1):
            cell_a = ws.cell(row=r, column=1)

            # Verificar si la celda de la columna A tiene fondo amarillo
            is_yellow = False
            if cell_a.fill and cell_a.fill.start_color:
                rgb = str(cell_a.fill.start_color.rgb).upper()
                if "FFFF00" in rgb or "FFFF0000" in rgb:
                    is_yellow = True

            if is_yellow and cell_a.value is not None:
                cod_a = str(cell_a.value).strip()
                desc_b = (
                    str(ws.cell(row=r, column=2).value).strip()
                    if ws.cell(row=r, column=2).value
                    else ""
                )

                # ========================================================
                # PASO 1: Concatenar A y B con espacio (" ") -> Columna C
                # ========================================================
                col_c_concatenada = f"{cod_a} {desc_b}"

                # Tomar Costo (Col C origen) y Precio (Col D u E origen)
                costo_d = ws.cell(row=r, column=3).value or 0
                precio_e = ws.cell(row=r, column=4).value or 0

                # Desglose de talles/colores a partir de la columna F en adelante
                # (Recorremos las columnas de variantes del Excel original)
                for c_idx in range(6, ws.max_column + 1, 2):
                    talle_val = ws.cell(row=r, column=c_idx - 1).value or "U"
                    stock_val = ws.cell(row=r, column=c_idx).value
                    header_color = ws.cell(row=2, column=c_idx).value or "ÚNICO"

                    if (
                        stock_val is not None
                        and str(stock_val).isdigit()
                        and int(stock_val) > 0
                    ):
                        filas_nuevas.append(
                            {
                                "descripcion_concatenada": col_c_concatenada,
                                "categoria": sheet_name.upper(),
                                "costo": costo_d,
                                "precio": precio_e,
                                "talle": str(talle_val).strip(),
                                "color": str(header_color).strip().upper(),
                                "stock": int(stock_val),
                            }
                        )

    print(
        f"✅ Se detectaron {len(filas_nuevas)} variantes de productos nuevos."
    )

    # 2. Construir el DataFrame final con el formato del Modelo 2
    registros_finales = []
    codigo_padre = 503823  # Autoincremental base de código padre

    for idx, item in enumerate(filas_nuevas):
        registros_finales.append(
            {
                "Codigo padre": codigo_padre + idx,
                "hijo": None,
                # Columna C concatenada -> 'Descripción/ Nombre'
                "Descripción/ Nombre": item["descripcion_concatenada"],
                "Categoria": item["categoria"],
                "Proveedor": proveedor_id,
                # Columnas C y D (Origen) -> Columnas F y G (Destino)
                "Costo": item["costo"],
                "Precio": item["precio"],
                # Mapeo a Columnas H, I, J (Talle, Color, Stock)
                "Talle": item["talle"],
                "Color": item["color"],
                "Stock": item["stock"],
                "Año": proveedor_id,
            }
        )

    df_resultado = pd.DataFrame(registros_finales)

    # 3. Exportar al Excel definitivo
    df_resultado.to_excel(path_salida, index=False)
    print(f"🚀 Archivo procesado con éxito y guardado en: '{path_salida}'")


if __name__ == "__main__":
    procesar_carga_palermo(
        path_carga_palermo="CARGA PALERMO 2026 (1)_2.xlsx",
        path_modelo="Modelo 2 palermo 1407.xlsx",
        path_salida="Modelo_2_Palermo_Generado.xlsx",
    )