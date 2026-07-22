import os
import pandas as pd


def procesar_concatenacion_inicial(
    archivo_entrada: str,
    archivo_salida: str = "Resultado_Concatenado.xlsx",
    columna_a: int = 0,
    columna_b: int = 1,
    separador: str = " - ",
):
    """Lee un archivo Excel y realiza la concatenación directamente sobre las columnas

    iniciales de cada pestaña antes de realizar cualquier otro proceso.
    """
    if not os.path.exists(archivo_entrada):
        print(f"❌ Error: El archivo '{archivo_entrada}' no existe en este directorio.")
        return

    print(f"🔄 Cargando archivo: {archivo_entrada} ...")
    xls = pd.ExcelFile(archivo_entrada)
    writer = pd.ExcelWriter(archivo_salida, engine="openpyxl")

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(archivo_entrada, sheet_name=sheet_name)

        # Si la pestaña está vacía o no tiene suficientes columnas, se pasa tal cual
        if df.empty or df.shape[1] < 2:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            continue

        # Identificar las columnas por su posición (0 = Columna A, 1 = Columna B)
        col1_name = df.columns[columna_a] if isinstance(columna_a, int) else columna_a
        col2_name = df.columns[columna_b] if isinstance(columna_b, int) else columna_b

        # --- CONCATENACIÓN EN EL ARCHIVO INICIAL ---
        col1_limpia = df[col1_name].fillna("").astype(str).str.strip()
        col2_limpia = df[col2_name].fillna("").astype(str).str.strip()

        # Generar texto concatenado
        concatenacion_inicial = col1_limpia + separador + col2_limpia

        # Insertar la concatenación como la PRIMERA columna (posición 0) de la tabla inicial
        df.insert(0, "CONCATENACION_INICIAL", concatenacion_inicial)

        # Guardar en la pestaña correspondiente
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"  ✓ Pestaña '{sheet_name}' procesada ({len(df)} filas).")

    writer.close()
    print(f"\n✅ ¡Proceso completado con éxito!")
    print(f"📁 Archivo generado: '{archivo_salida}'")


if __name__ == "__main__":
    # Nombre de tu archivo de entrada (asegúrate de que esté en la misma carpeta)
    NOMBRE_ARCHIVO_INPUT = "CARGA PALERMO 2026 (1).xlsx"

    # Nombre del archivo concatenado final que generará
    NOMBRE_ARCHIVO_OUTPUT = "CARGA_PALERMO_CONCATENADO.xlsx"

    # Ejecutar la función
    procesar_concatenacion_inicial(
        archivo_entrada=NOMBRE_ARCHIVO_INPUT,
        archivo_salida=NOMBRE_ARCHIVO_OUTPUT,
        columna_a=0,  # Columna A (Código)
        columna_b=1,  # Columna B (Descripción)
        separador=" - ",  # Formato: "CÓDIGO - DESCRIPCIÓN"
    )