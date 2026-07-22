from io import BytesIO
import openpyxl
import streamlit as st

st.set_page_config(
    page_title="Procesador de Cargas - Palermo", page_icon="📦", layout="centered"
)

st.title("📦 Procesador de Archivos Palermo")
st.markdown(
    "Sube tu archivo de carga original (`.xlsx`). La plataforma filtrará las filas amarillas y generará la matriz con el formato correcto."
)

archivo_subido = st.file_uploader(
    "Arrastra o selecciona el archivo Excel de origen", type=["xlsx"]
)


def procesar_excel(file):
    wb_origen = openpyxl.load_workbook(file, data_only=True)
    sheet_origen = wb_origen.active

    wb_salida = openpyxl.Workbook()
    sheet_salida = wb_salida.active
    sheet_salida.title = "Matriz Procesada"

    headers = [
        "Codigo padre",
        "hijo",
        "Descripción/ Nombre",
        "Categoria",
        "Proveedor",
        "Costo",
        "Precio",
        "Talle",
        "Color",
        "Stock",
        "Año",
    ]
    sheet_salida.append(headers)

    articulos_procesados = 0

    for fila in range(1, sheet_origen.max_row + 1):
        celda_ref = sheet_origen.cell(row=fila, column=1)
        fill = celda_ref.fill

        es_amarillo = False
        if fill and fill.fill_type and fill.start_color and fill.start_color.rgb:
            color_hex = str(fill.start_color.rgb).upper()
            # Detección de variantes de amarillo de Excel
            if any(
                y in color_hex
                for y in [
                    "FFFF00",
                    "FFF200",
                    "FFFF0000",
                    "FFFFCC00",
                    "FFFFE599",
                ]
            ):
                es_amarillo = True

        if es_amarillo:
            val_a = sheet_origen.cell(row=fila, column=1).value or ""
            val_b = sheet_origen.cell(row=fila, column=2).value or ""
            val_c = sheet_origen.cell(row=fila, column=3).value or ""
            val_d = sheet_origen.cell(row=fila, column=4).value or ""
            val_f = sheet_origen.cell(row=fila, column=6).value or ""
            val_g = sheet_origen.cell(row=fila, column=7).value or ""

            # Regla 1: Concatenar A y B
            concatenado_c = f"{str(val_a).strip()} {str(val_b).strip()}".strip()

            # Mapeo según tus reglas
            nueva_fila = [None] * 11
            nueva_fila[2] = concatenado_c  # Columna C -> Concatenado A + B
            nueva_fila[5] = val_c  # Columna F -> C del origen
            nueva_fila[6] = val_d  # Columna G -> D del origen
            nueva_fila[7] = val_f  # Columna H -> F del origen
            nueva_fila[9] = val_g  # Columna J -> G del origen

            sheet_salida.append(nueva_fila)
            articulos_procesados += 1

    output = BytesIO()
    wb_salida.save(output)
    output.seek(0)
    return output, articulos_procesados


if archivo_subido is not None:
    if st.button("🚀 Procesar Archivo", type="primary"):
        with st.spinner("Procesando artículos amarillos..."):
            excel_procesado, cantidad = procesar_excel(archivo_subido)

            if cantidad > 0:
                st.success(
                    f"¡Listo! Se detectaron y procesaron **{cantidad}** artículos nuevos."
                )
                st.download_button(
                    label="📥 Descargar Excel Final",
                    data=excel_procesado,
                    file_name="Modelo_2_Palermo_Procesado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning(
                    "No se encontraron celdas con relleno amarillo en la columna A del archivo subido."
                )