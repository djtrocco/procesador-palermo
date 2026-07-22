import os
import openpyxl
import pandas as pd


def procesar_carga_palermo(
    archivo_entrada: str,
    archivo_salida: str = "Modelo_2_Palermo_Generado.xlsx",
    proveedor_id: str = "1407",
    codigo_padre_inicial: int = 503823,
):
    """Procesa el archivo 'Carga palermo':

    1. Filtra filas con relleno amarillo (productos nuevos) en cualquier pestaña.
    2. Concatena Col A y Col B -> Col C ('CÓDIGO DESCRIPCIÓN').
    3. Si la fila tiene múltiples columnas de colores/talles con valores (stock > 0),
       crea una NUEVA FILA duplicando de la col C a la G (Descripción, Categoría, Proveedor, Costo, Precio).
    """
    if not os.path.exists(archivo_entrada):
        print(
            f"❌ Error: No se encontró el archivo de entrada '{archivo_entrada}'."
        )
        return

    print(f"🔄 Leyendo archivo inicial: {archivo_entrada} ...")

    wb = openpyxl.load_workbook(archivo_entrada, data_only=True)
    filas_destino = []
    codigo_padre_actual = codigo_padre_inicial

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # Obtener los encabezados de colores/talles de la fila 1 (si existen)
        encabezados_col = {}
        for col_idx in range(5, ws.max_column + 1):
            val_header = ws.cell(row=1, column=col_idx).value
            if val_header is not None and str(val_header).strip() != "":
                encabezados_col[col_idx] = str(val_header).strip().upper()

        # Recorrer filas desde la fila 2
        for r in range(2, ws.max_row + 1):
            cell_a = ws.cell(row=r, column=1)

            # Detectar relleno amarillo en la Columna A (#FFFF00)
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
                # PASO 1: CONCATENACIÓN "=A & ' ' & B" -> Columna C Destino
                # ========================================================
                descripcion_concatenada = f"{cod_a} {desc_b}"

                # Columnas C y D origen -> Costo y Precio
                costo_val = ws.cell(row=r, column=3).value or 0
                precio_val = ws.cell(row=r, column=4).value or 0
                categoria_val = sheet_name.upper()

                # ========================================================
                # PASO 2: EVALUAR SI HAY MÚLTIPLES VALORES (VARIANTES/STOCK)
                # ========================================================
                variantes_encontradas = []

                for col_idx, col_nombre in encabezados_col.items():
                    val_stock = ws.cell(row=r, column=col_idx).value

                    # Si el valor de stock es un número válido y mayor a 0
                    if (
                        val_stock is not None
                        and isinstance(val_stock, (int, float))
                        and val_stock > 0
                    ):
                        variantes_encontradas.append(
                            {
                                "color": col_nombre,
                                "talle": "U",  # Talle por defecto si es único
                                "stock": int(val_stock),
                            }
                        )

                # Si no se halló stock en las columnas mapeadas pero es amarillo
                if not variantes_encontradas:
                    variantes_encontradas.append(
                        {"color": "ÚNICO", "talle": "U", "stock": 1}
                    )

                # ========================================================
                # PASO 3: DUPLICAR COLUMNAS C A G EN CADAS FILA NUEVA
                # ========================================================
                for var in variantes_encontradas:
                    filas_destino.append(
                        {
                            "Codigo padre": codigo_padre_actual,  # Col A destino
                            "hijo": None,  # Col B destino
                            # --- COPIADO C A G (Se repite para cada nueva fila creada) ---
                            "Descripción/ Nombre": descripcion_concatenada,  # Col C
                            "Categoria": categoria_val,  # Col D
                            "Proveedor": proveedor_id,  # Col E
                            "Costo": costo_val,  # Col F
                            "Precio": precio_val,  # Col G
                            # -------------------------------------------------------------
                            "Talle": var["talle"],  # Col H
                            "Color": var["color"],  # Col I
                            "Stock": var["stock"],  # Col J
                            "Año": proveedor_id,  # Col K
                        }
                    )
                    # Incrementar el código padre por fila única asignada
                    codigo_padre_actual += 1

        print(f"  ✓ Pestaña '{sheet_name}' procesada.")

    # Guardar en DataFrame y exportar al formato exacto del Modelo 2
    df_resultado = pd.DataFrame(filas_destino)

    if not df_resultado.empty:
        df_resultado.to_excel(archivo_salida, index=False)
        print(f"\n✅ ¡Proceso completado con éxito!")
        print(f"📊 Filas totales generadas: {len(df_resultado)}")
        print(f"📁 Archivo final: '{archivo_salida}'")
    else:
        print("\n⚠️ No se encontraron filas amarillas en el archivo cargado.")


if __name__ == "__main__":
    NOMBRE_INPUT = "CARGA PALERMO 2026 (1).xlsx"
    NOMBRE_OUTPUT = "Modelo_2_Palermo_Generado.xlsx"

    procesar_carga_palermo(archivo_entrada=NOMBRE_INPUT, archivo_salida=NOMBRE_OUTPUT)