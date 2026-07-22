import io
import openpyxl
import pandas as pd
import streamlit as st

# Configuración de página
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


def obtener_nombre_color(ws, col_idx):
    """Obtiene el nombre limpio del color desde los encabezados (filas 3, 2 o 1)."""
    for r in [3, 2, 1]:
        val = ws.cell(row=r, column=col_idx).value
        if val and isinstance(val, str):
            clean = val.strip().upper()
            if clean not in [
                "COSTO",
                "TC",
                "EF",
                "TALLE",
                "TARJETA",
                "FECHA",
            ] and not clean.startswith("FECHA"):
                if clean == "NEG":
                    return "NEGRO"
                return clean
    return None


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

                # Omitir encabezados accidentales
                if "FECHA" in desc_b or "FECHA" in cod_a or cod_a == "None":
                    continue

                # Concatenar (A + B) en la Columna C de salida
                col_c_concatenada = f"{cod_a} {desc_b}"

                costo = ws.cell(row=r, column=3).value or 0
                precio = ws.cell(row=r, column=4).value or 0
                categoria = sheetname.upper()

                # Talle general por defecto en Columna 6 (F)
                val_f = ws.cell(row=r, column=6).value
                talle_col_f = str(val_f).strip() if val_f is not None else ""

                # Buscar variantes reales de stock recorriendo desde la columna 6 en adelante
                variantes = []
                for c in range(6, ws.max_column + 1):
                    val = ws.cell(row=r, column=c).value

                    # Verificar si la celda contiene una cantidad numérica válida de stock
                    if (
                        isinstance(val, (int, float))
                        and val > 0
                        and not isinstance(val, bool)
                    ):
                        color = obtener_nombre_color(ws, c)

                        # Si la columna tiene número de stock pero no tiene nombre de color en el header,
                        # miramos la columna previa si tenía el nombre
                        if not color and c > 1:
                            color = obtener_nombre_color(ws, c - 1)

                        if not color:
                            color = "NEGRO"

                        # Buscar el talle asociado a este stock:
                        # Se busca en la celda anterior (columna de talle del par) o en la Columna F
                        val_talle_par = ws.cell(row=r, column=c - 1).value
                        if (
                            val_talle_par is not None
                            and str(val_talle_par).strip() != ""
                            and not isinstance(val_talle_par, (int, float))
                        ):
                            talle_var = str(val_talle_par).strip()
                        elif talle_col_f != "" and not isinstance(
                            val_f, (int, float)
                        ):
                            talle_var = talle_col_f
                        else:
                            talle_var = "U"

                        variantes.append((color, talle_var, int(val)))

                # Si la fila estaba en amarillo pero no tenía valores de stock numéricos
                if not variantes:
                    talle_def = talle_col_f if talle_col_f != "" else "U"
                    registros_salida.append(
                        {
                            "Codigo padre": codigo_padre,
                            "hijo": None,
                            "Descripción/ Nombre": col_c_concatenada,
                            "Categoria": categoria,
                            "Proveedor": "1407",
                            "Costo": costo,
                            "Precio": precio,
                            "Talle": talle_def,
                            "Color": "NEGRO",
                            "Stock": 1,
                            "Año": "1407",
                        }
                    )
                    codigo_padre += 1

                # Si solo tiene 1 variante (ej. V699 solo en PLATA)
                elif len(variantes) == 1:
                    color_var, talle_var, stock_var = variantes[0]
                    registros_salida.append(
                        {
                            "Codigo padre": codigo_padre,
                            "hijo": None,
                            "Descripción/ Nombre": col_c_concatenada,
                            "Categoria": categoria,
                            "Proveedor": "1407",
                            "Costo": costo,
                            "Precio": precio,
                            "Talle": talle_var,
                            "Color": color_var,  # Ej: PLATA
                            "Stock": stock_var,
                            "Año": "1407",
                        }
                    )
                    codigo_padre += 1

                # Si tiene múltiples variantes con stock
                else:
                    for color_var, talle_var, stock_var in variantes:
                        registros_salida.append(
                            {
                                "Codigo padre": codigo_padre,
                                "hijo": None,
                                "Descripción/ Nombre": col_c_concatenada,
                                "Categoria": categoria,
                                "Proveedor": "1407",
                                "Costo": costo,
                                "Precio": precio,
                                "Talle": talle_var,
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
        st.dataframe(df_resultado.head(25), use_container_width=True)

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