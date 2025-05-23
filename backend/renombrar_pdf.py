import os

def renombrar_pdfs_en_mayusculas(carpeta):
    """
    Renombra todos los archivos PDF en una carpeta específica, convirtiendo sus nombres a mayúsculas.
    """
    try:
        # Verificar si la carpeta existe
        if not os.path.isdir(carpeta):
            print(f"La carpeta especificada no existe: {carpeta}")
            return

        # Listar todos los archivos en la carpeta
        archivos = os.listdir(carpeta)

        # Procesar solo los archivos PDF
        for archivo in archivos:
            if archivo.lower().endswith(".pdf"):
                ruta_actual = os.path.join(carpeta, archivo)  # Ruta completa del archivo actual
                nombre_nuevo = archivo.upper()  # Convertir el nombre a mayúsculas
                ruta_nueva = os.path.join(carpeta, nombre_nuevo)  # Nueva ruta con el nombre en mayúsculas

                # Renombrar el archivo
                os.rename(ruta_actual, ruta_nueva)
                print(f"Renombrado: {archivo} -> {nombre_nuevo}")

        print("Renombrado completado.")
    except Exception as e:
        print(f"Se produjo un error: {e}")

# Solicitar al usuario la carpeta a procesar
if __name__ == "__main__":
    carpeta = input("Ingrese la ruta completa de la carpeta con los PDFs: ").strip()
    renombrar_pdfs_en_mayusculas(carpeta)