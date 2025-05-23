import pandas as pd
import time
from tkinter import Tk, Label, Button, filedialog, messagebox
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import quote
from webdriver_manager.chrome import ChromeDriverManager

def cargar_datos(ruta_excel):
    """
    Carga los datos del archivo Excel y retorna un DataFrame.
    """
    try:
        datos = pd.read_excel(ruta_excel)
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
    mensaje = mensaje = (f"{saludo} {nombre},\n\n"
           "Escribe Marlon Mosquera del Centro Agropecuario La Granja del Espinal. "
           f"Estamos validando los documentos de matrícula para la tecnología en {nombre_tecnologo}. "
           "Hemos identificado que es necesario enviar nuevamente los siguientes documentos:\n\n"
           f"{documentos_formateados}\n\n"
           "Por favor, realiza las correcciones indicadas y carga nuevamente todo el paquete completo, "
           "incluyendo los documentos anteriormente mencionados, en el enlace de matrícula.\n\n"
           "IMPORTANTE:\n"
           "- En caso de faltar un documento, se debe enviar todo el paquete completo, no solo el documento faltante.\n"
           "- Los reportes de EPS deben obtenerse directamente desde la página oficial de ADRES.\n"
           "- Los resultados del ICFES deben generarse en formato PDF desde la página oficial (NO capturas de pantalla).\n\n"
           "Te solicitamos realizar este proceso lo antes posible, preferiblemente el día de hoy.\n\n"
           "Agradecemos tu colaboración.")

    return mensaje

def enviar_mensaje(numero, mensaje, driver):
    try:
        mensaje_codificado = quote(mensaje)
        url = f"https://web.whatsapp.com/send?phone={numero}&text={mensaje_codificado}"
        driver.get(url)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
        )     
        send_button = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
        send_button.click()
        time.sleep(3)
    except Exception as e:
        print(f"Error al enviar el mensaje a {numero}: {e}")

def procesar_archivo(ruta_excel):
    """
    Procesa el archivo Excel y envía los mensajes.
    """
    datos = cargar_datos(ruta_excel)
    if datos is None:
        return

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-extensions")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    # messagebox.showinfo("WhatsApp Web", "Por favor, escanee el código QR en WhatsApp Web.")
    driver.get("https://web.whatsapp.com/")
    time.sleep(4)
    WebDriverWait(driver, 40).until(
        EC.presence_of_element_located((By.XPATH, '//canvas[@aria-label="Scan this QR code to link a device!"]'))
    )
    time.sleep(5)

    for _, fila in datos.iterrows():
        nombre = fila.get("Nombre")
        documentos_faltantes = fila.get("Documentos_Faltantes")
        nombre_tecnologo = fila.get("Nombre_Tecnologo")
        numero_telefono = fila.get("Numero_Telefono")

        if pd.isna(nombre) or pd.isna(documentos_faltantes) or pd.isna(nombre_tecnologo) or pd.isna(numero_telefono):
            print(f"Fila con datos incompletos: {fila}")
            continue

        mensaje = generar_mensaje(nombre, documentos_faltantes, nombre_tecnologo)
        enviar_mensaje(f"+57{int(numero_telefono)}", mensaje, driver)

    messagebox.showinfo("Finalizado", "Los mensajes se han enviado con exito.")
    driver.quit()

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
    ventana.title("Envío de Mensajes WhatsApp")
    ventana.geometry("400x200")

    Label(ventana, text="Envío Automático de Mensajes", font=("Arial", 16)).pack(pady=10)
    Button(ventana, text="Seleccionar Archivo Excel", command=seleccionar_archivo, width=25).pack(pady=20)
    Button(ventana, text="Salir", command=ventana.quit, width=25).pack(pady=10)

    ventana.mainloop()

if __name__ == "__main__":
    iniciar_interfaz()
