import pandas as pd
from tkinter import Tk, Label, Button, filedialog, messagebox
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def cargar_datos(ruta_excel):
    """
    Carga los datos del archivo Excel y valida que tenga las columnas necesarias.
    """
    try:
        datos = pd.read_excel(ruta_excel)
        columnas_requeridas = ["Nombre", "Documentos_Faltantes", "Nombre_Tecnologo", "Correo"]
        if not all(col in datos.columns for col in columnas_requeridas):
            messagebox.showerror("Error", f"El archivo debe contener las columnas: {', '.join(columnas_requeridas)}")
            return None
        return datos
    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")
        return None


def obtener_saludo():
    """
    Determina el saludo según la hora del día.
    """
    hora_actual = datetime.now().hour
    if 5 <= hora_actual < 12:
        return "Buenos días"
    elif 12 <= hora_actual < 18:
        return "Buenas tardes"
    else:
        return "Buenas noches"


def generar_mensaje(nombre, documentos_faltantes, nombre_tecnologo):
    """
    Genera un mensaje personalizado para el estudiante.
    """
    saludo = obtener_saludo()
    documentos_lista = [doc.strip() for doc in str(documentos_faltantes).split(',')]
    documentos_formateados = "\n".join([f"{i+1}. {doc}" for i, doc in enumerate(documentos_lista)])
    mensaje = (f"{saludo} {nombre}, escribe Marlon Mosquera del Centro Agropecuario La Granja del Espinal. "
               f"Estamos validando los documentos de matrícula para la tecnología en {nombre_tecnologo} y "
               f"encontramos que se deben enviar nuevamente los siguientes documentos:\n\n"
               f"{documentos_formateados}\n\n"
               "Agradecemos el ajuste a los documentos y VOLVER A CARGAR TODO EL PAQUETE AL LINK DE MATRICULA. "
               "Esto debe quedar listo en lo posible hoy mismo.")
    return mensaje


def enviar_correo(correo, mensaje):
    """
    Envía un correo electrónico al destinatario.
    """
    try:
        remitente = ""  # Cambiar por tu correo
        clave = ""  # Cambiar por tu contraseña o token de aplicación
        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(remitente, clave)

        # Crear el correo
        correo_msg = MIMEMultipart()
        correo_msg["From"] = remitente
        correo_msg["To"] = correo
        correo_msg["Subject"] = "Validación de Documentos"
        correo_msg.attach(MIMEText(mensaje, "plain"))

        # Enviar el correo
        servidor.sendmail(remitente, correo, correo_msg.as_string())
        servidor.quit()
        print(f"Correo enviado correctamente a {correo}")
        return "Exitoso"
    except Exception as e:
        print(f"Error al enviar el correo a {correo}: {e}")
        return "Fallido"


def procesar_archivo(ruta_excel):
    """
    Procesa el archivo Excel y envía correos según los datos disponibles.
    """
    datos = cargar_datos(ruta_excel)
    if datos is None:
        return

    registros = []

    for _, fila in datos.iterrows():
        nombre = fila.get("Nombre")
        documentos_faltantes = fila.get("Documentos_Faltantes")
        nombre_tecnologo = fila.get("Nombre_Tecnologo")
        correo = fila.get("Correo")

        if pd.isna(nombre) or pd.isna(documentos_faltantes) or pd.isna(nombre_tecnologo) or pd.isna(correo):
            registros.append({"Correo": correo, "Estado": "Fallido", "Motivo": "Datos incompletos"})
            continue

        mensaje = generar_mensaje(nombre, documentos_faltantes, nombre_tecnologo)
        estado = enviar_correo(correo, mensaje)
        registros.append({"Correo": correo, "Estado": estado})

    # Guardar registro en un archivo CSV
    ruta_registro = "registro_envios.csv"
    pd.DataFrame(registros).to_csv(ruta_registro, index=False)
    messagebox.showinfo("Finalizado", f"El proceso ha finalizado. Registro guardado en {ruta_registro}.")


def seleccionar_archivo():
    """
    Permite al usuario seleccionar un archivo Excel y procesarlo.
    """
    ruta_excel = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if ruta_excel:
        procesar_archivo(ruta_excel)


def iniciar_interfaz():
    """
    Inicia la interfaz gráfica.
    """
    ventana = Tk()
    ventana.title("Envío de Correos Automático")
    ventana.geometry("400x200")

    Label(ventana, text="Envío Automático de Correos", font=("Arial", 16)).pack(pady=10)
    Button(ventana, text="Seleccionar Archivo Excel", command=seleccionar_archivo, width=25).pack(pady=20)
    Button(ventana, text="Salir", command=ventana.quit, width=25).pack(pady=10)

    ventana.mainloop()


if __name__ == "__main__":
    iniciar_interfaz()
