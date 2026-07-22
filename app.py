import io
import openpyxl
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Carga de Productos Web", page_icon="📦", layout="wide"
)

st.title("📦 Procesador de Productos Nuevos - Palermo")
st.markdown(
    "Sube tu archivo inicial Excel para detectar los productos nuevos (en **amarillo**), "
    "concatenar código/descripción y armar el archivo final formateado."
)

uploaded_file = st.file_uploader(
    "Selecciona el archivo Excel inicial (ej. 'Carga Palermo')",
    type=["xlsx", "xls"],
)


def procesar_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

    registros_salida = []
    codigo_padre = 503823  # Autoincremental inicial

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]

        for r in range(1, ws.max_row + 1):
            cell_a = ws.cell(row=r, column=1)

            # Detectar relleno amarillo (#FFFF00) en la Columna A
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

                # Omitir encabezados
                if "FECHA" in desc_b or "FECHA" in cod_a or cod_a == "None":
                    continue

                # REGLA: Concatenar (A + B) en la Columna C de salida
                col_c_concatenada = f"{cod_a} {desc_b}"

                costo = ws.cell(row=r, column=3).value or 0
                precio = ws.cell(row=r, column=4).value or 0
                categoria = sheetname.upper()

                # NUEVA REGLA: Copiar el valor original de la Columna F (Columna 6 del inicial) para el Talle
                val_talle_orig = ws.cell(row=r, column=6).value
                talle_final = (
                    str(val_talle_orig).strip()
                    if val_talle_orig is not None
                    else "U"
                )

                # Detectar variantes de stock en columnas posteriores (G en adelante)
                variantes = []
                for c in range(7, ws.max_column + 1):
                    val = ws.cell(row=r, column=c).value
                    if (
                        isinstance(val, (int, float))
                        and val > 0
                        and not isinstance(val, bool)
                    ):
                        # Detectar nombre del color en encabezados
                        h1 = ws.cell(row=1, column=c).value
                        h2 = ws.cell(row=2, column=c).value
                        h3 = ws.cell(row=3, column=c).value

                        color_name = "NEGRO"
                        for h in [h3, h2, h1]:
                            if (
                                h
                                and isinstance(h, str)
                                and not h.startswith("FECHA")
                                and h.strip().upper()
                                not in [
                                    "COSTO",
                                    "TC",
                                    "EF",
                                    "NEG",
                                    "TALLE",
                                    "TARJETA",
                                ]
                            ):
                                color_name = h.strip().upper()
                                break
                            elif (
                                h
                                and isinstance(h, str)
                                and h.strip().upper() == "NEG"
                            ):
                                color_name = "NEGRO"
                                break

                        variantes.append((color_name, int(val)))

                # Si no hay variantes registradas en columnas posteriores
                if not variantes:
                    registros_salida.append(
                        {
                            "Codigo padre": codigo_padre,
                            "hijo": None,
                            "Descripción/ Nombre": col_c_concatenada,
                            "Categoria": categoria,
                            "Proveedor": "1407",
                            "Costo": costo,
                            "Precio": precio,
                            "Talle": talle_final,  # Columna H toma el valor de F
                            "Color": "NEGRO",
                            "Stock": 1,
                            "Año": "1407",
                        }
                    )
                    codigo_padre += 1

                elif len(variantes) == 1:
                    color_var, stock_var = variantes[0]
                    registros_salida.append(
                        {
                            "Codigo padre": codigo_padre,
                            "hijo": None,
                            "Descripción/ Nombre": col_c_concatenada,
                            "Categoria": categoria,
                            "Proveedor": "1407",
                            "Costo": costo,
                            "Precio": precio,
                            "Talle": talle_final,  # Columna H toma el valor de F
                            "Color": color_var,
                            "Stock": stock_var,
                            "Año": "1407",
                        }
                    )
                    codigo_padre += 1

                else:
                    for color_var, stock_var in variantes:
                        registros_salida.append(
                            {
                                "Codigo padre": codigo_padre,
                                "hijo": None,
                                "Descripción/ Nombre": col_c_concatenada,
                                "Categoria": categoria,
                                "Proveedor": "1407",
                                "Costo": costo,
                                "Precio": precio,
                                "Talle": talle_final,  # Columna H toma el valor de F
                                "Color": color_var,
                                "Stock": stock_var,
                                "Año": "1407",
                            }
                        )
                        codigo_padre += 1

    return pd.DataFrame(registros_salida)


if uploaded_file is not None:
    st.info("🔄 Procesando archivo...")
    bytes_data = uploaded_file.read()

    df_resultado = procesar_excel(bytes_data)

    if not df_resultado.empty:
        st.success(
            f"✅ ¡Proceso finalizado! Se generaron **{len(df_resultado)} filas** de productos."
        )

        st.subheader("Vista previa del archivo generado:")
        st.dataframe(df_resultado.head(20), use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False, sheet_name="Hoja1")

        st.download_button(
            label="📥 Descargar Excel Procesado (Modelo 2)",
            data=buffer.getvalue(),
            file_name="Modelo_2_Palermo_Generado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning(
            "⚠️ No se detectaron celdas rellenas en color amarillo en el archivo subido."
        )