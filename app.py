import pandas as pd
import os

def procesar_concatenacion_inicial(
    archivo_entrada: str, 
    archivo_salida: str = "Resultado_Concatenado.xlsx",
    columna_a: int = 0, 
    columna_b: int = 1, 
    separador: str = " - "
):
    """
    Lee cualquier archivo Excel, aplica la concatenación sobre las columnas de origen
    directamente en los datos iniciales y guarda el resultado manteniendo todas las pestañas.
    
    :param archivo_entrada: Ruta del archivo Excel (.xlsx o .xls).
    :param archivo_salida: Nombre o ruta del archivo generado.
    :param columna_a: Índice (0 para Columna A) o nombre de la primera columna a concatenar.
    :param columna_b: Índice (1 para Columna B) o nombre de la segunda columna a concatenar.
    :param separador: Carácter o texto de unión (por defecto ' - ').
    """
    if not os.path.exists(archivo_entrada):
        raise FileNotFoundError(f"El archivo '{archivo_entrada}' no existe.")

    # 1. Cargar el libro Excel completo
    xls = pd.ExcelFile(archivo_entrada)
    writer = pd.ExcelWriter(archivo_salida, engine='openpyxl')
    
    print(f"Procesando el archivo: {archivo_entrada}")

    for sheet_name in xls.sheet_names:
        # Leer datos crudos de cada pestaña
        df = pd.read_excel(archivo_entrada, sheet_name=sheet_name)

        if df.empty or df.shape[1] < 2:
            # Si la hoja está vacía o tiene menos de 2 columnas, se escribe tal cual
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            continue

        # 2. Identificar dinámicamente el nombre de las columnas a concatenar
        col1_name = df.columns[columna_a] if isinstance(columna_a, int) else columna_a
        col2_name = df.columns[columna_b] if isinstance(columna_b, int) else columna_b

        # 3. REALIZAR LA CONCATENACIÓN EN EL ARCHIVO INICIAL
        # Convertimos a string y limpiamos espacios vacíos (NaN)
        col1_limpia = df[col1_name].fillna('').astype(str).str.strip()
        col2_limpia = df[col2_name].fillna('').astype(str).str.strip()

        # Creamos la serie concatenada
        concatenacion_inicial = col1_limpia + separador + col2_limpia

        # 4. Insertar la columna concatenada al inicio de la tabla (Posición 0)
        df.insert(0, 'CONCATENACION_INICIAL', concatenacion_inicial)

        # 5. Guardar la hoja procesada en el nuevo archivo Excel
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f" Pestaña '{sheet_name}' procesada ({len(df)} filas).")

    # Guardar cambios
    writer.close()
    print(f"\n Archivo final generado con éxito: '{archivo_salida}'")


# ==============================================================================
# EJECUCIÓN DEL SCRIPT
# ==============================================================================
if __name__ == "__main__":
    # Nombre del archivo recibido
    archivo_subido = "CARGA PALERMO 2026 (1).xlsx"
    
    # Ejecutar la función
    procesar_concatenacion_inicial(
        archivo_entrada=archivo_subido,
        archivo_salida="CARGA_PALERMO_CONCATENADO.xlsx",
        columna_a=0,        # Primera columna (Columna A / Código)
        columna_b=1,        # Segunda columna (Columna B / Descripción)
        separador=" - "     # Formato: "CÓDIGO - DESCRIPCIÓN"
    )