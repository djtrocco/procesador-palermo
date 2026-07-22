# Mapeo inicial/dinámico de colores por columna
current_colors = {}

for row in range(1, sheet.max_row + 1):
    val_b = str(sheet.cell(row=row, column=2).value or "").upper()
    val_c = str(sheet.cell(row=row, column=3).value or "").upper()

    # Si la fila es un encabezado de bloque, actualizamos la lista de colores según la columna
    if "FECHA" in val_b or "COSTO" in val_c:
        for col in range(1, sheet.max_column + 1):
            col_header = sheet.cell(row=row, column=col).value
            if col_header and str(col_header).strip():
                # Guardamos el nombre del color para esa columna específica
                current_colors[col] = str(col_header).strip()
        continue

    # Al procesar el producto (ej. V699)
    codigo = sheet.cell(row=row, column=1).value
    if codigo == "V699":
        # Para la columna 13 con cantidad 3:
        # Usamos current_colors[13], que para el bloque FECHA 1407 devolverá 'PLATA'
        color = current_colors.get(13, "DESCONOCIDO")
        # Resultado -> 'PLATA'