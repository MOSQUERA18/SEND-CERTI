import pandas as pd
import time
import threading
from tkinter import Tk, Label, Button, filedialog, messagebox, Toplevel,PhotoImage
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import quote
from webdriver_manager.chrome import ChromeDriverManager
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import tkinter.font as tkFont



def obtener_ruta_descargas():
    """ Obtiene la ruta de la carpeta de Descargas del usuario """
    if os.name == "nt":  # Windows
        return os.path.join(os.environ["USERPROFILE"], "Downloads")
    else:  # MacOS y Linux
        return os.path.join(os.path.expanduser("~"), "Downloads")

def descargar_plantilla(tipo):
    """ Descarga una plantilla de Excel dependiendo del tipo de mensaje en la carpeta de Descargas del usuario """
    ruta_descargas = obtener_ruta_descargas()
    os.makedirs(ruta_descargas, exist_ok=True)  # Asegurar que la carpeta existe

    if tipo == "whatsapp":
        ruta_guardado = os.path.join(ruta_descargas, "plantilla_whatsapp.xlsx")
        columnas = ["Nombre", "Documentos_Faltantes", "Nombre_Tecnologo", "Numero_Telefono"]
    else:
        ruta_guardado = os.path.join(ruta_descargas, "plantilla_correo.xlsx")
        columnas = ["Nombre", "Documentos_Faltantes", "Nombre_Tecnologo", "Correo"]

    df = pd.DataFrame(columns=columnas)
    df.to_excel(ruta_guardado, index=False)

    # Mensaje de confirmación con la ruta
    from tkinter import messagebox
    messagebox.showinfo("Éxito", f"Plantilla de {tipo.capitalize()} guardada en: {ruta_guardado}")
def mostrar_cargando(ventana):
    """ Muestra una ventana emergente de carga """
    global ventana_carga
    ventana_carga = Toplevel(ventana)
    ventana_carga.title("Cargando...")
    ventana_carga.geometry("250x100")
    ventana_carga.resizable(False, False)
    
    Label(ventana_carga, text="Procesando, por favor espere...", font=("Arial", 10)).pack(pady=10)
    
    # Deshabilitar la ventana principal mientras se carga
    ventana_carga.grab_set()
    ventana.update()

def ocultar_cargando():
    """ Oculta la ventana emergente de carga """
    ventana_carga.destroy()

def cargar_datos(ruta_excel):
    try:
        datos = pd.read_excel(ruta_excel)
        return datos
    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")
        return None

def obtener_saludo():
    hora_actual = datetime.now().hour
    if 5 <= hora_actual < 12:
        return "Buenos días"
    elif 12 <= hora_actual < 18:
        return "Buenas tardes"
    else:
        return "Buenas noches"

def generar_mensaje(nombre, documentos_faltantes, nombre_tecnologo):
    saludo = obtener_saludo()
    documentos_lista = [doc.strip() for doc in str(documentos_faltantes).split(',')]
    documentos_formateados = "\n".join([f"{i+1}. {doc}" for i, doc in enumerate(documentos_lista)])
    mensaje = (f"{saludo} {nombre},\n\n"
               "Escribe Marlon Mosquera del Centro Agropecuario La Granja del Espinal. "
               f"Estamos validando los documentos de matrícula para la tecnología en {nombre_tecnologo}. "
               "Hemos identificado que es necesario enviar nuevamente los siguientes documentos:\n\n"
               f"{documentos_formateados}\n\n"
               "Por favor, realiza las correcciones indicadas y carga nuevamente todo el paquete completo.")
    return mensaje

def enviar_mensajes_whatsapp(ruta_excel, ventana):
    mostrar_cargando(ventana)
    
    datos = cargar_datos(ruta_excel)
    if datos is None:
        ocultar_cargando()
        return
    
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get("https://web.whatsapp.com/")
    
    messagebox.showinfo("WhatsApp Web", "Inicie sesión en WhatsApp Web y presione Aceptar cuando los chats estén cargados.")
    
    for _, fila in datos.iterrows():
        nombre = fila.get("Nombre")
        documentos_faltantes = fila.get("Documentos_Faltantes")
        nombre_tecnologo = fila.get("Nombre_Tecnologo")
        numero_telefono = fila.get("Numero_Telefono")

        if pd.isna(nombre) or pd.isna(documentos_faltantes) or pd.isna(nombre_tecnologo) or pd.isna(numero_telefono):
            continue

        mensaje = generar_mensaje(nombre, documentos_faltantes, nombre_tecnologo)
        mensaje_codificado = quote(mensaje)
        url = f"https://web.whatsapp.com/send?phone=+57{int(numero_telefono)}&text={mensaje_codificado}"
        driver.get(url)
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[@data-icon="send"]'))
            )
            send_button = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
            send_button.click()
            time.sleep(3)
        except Exception as e:
            print(f"Error al enviar mensaje a {numero_telefono}: {e}")
    
    driver.quit()
    ocultar_cargando()
    messagebox.showinfo("Finalizado", "Mensajes enviados correctamente.")

def seleccionar_archivo(opcion, ventana):
    ruta_excel = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
    if ruta_excel:
        if opcion == 1:
            threading.Thread(target=enviar_mensajes_whatsapp, args=(ruta_excel, ventana)).start()
        elif opcion == 2:
            threading.Thread(target=procesar_correo, args=(ruta_excel, ventana)).start()

def procesar_correo(ruta_excel, ventana):
    mostrar_cargando(ventana)
    
    datos = cargar_datos(ruta_excel)
    if datos is None:
        ocultar_cargando()
        return
    
    for _, fila in datos.iterrows():
        nombre = fila.get("Nombre")
        documentos_faltantes = fila.get("Documentos_Faltantes")
        nombre_tecnologo = fila.get("Nombre_Tecnologo")
        correo = fila.get("Correo")
        
        if pd.isna(nombre) or pd.isna(documentos_faltantes) or pd.isna(nombre_tecnologo) or pd.isna(correo):
            continue
        
        mensaje = generar_mensaje(nombre, documentos_faltantes, nombre_tecnologo)
        enviar_correo(correo, mensaje)
    
    ocultar_cargando()
    messagebox.showinfo("Finalizado", "Correos enviados correctamente.")

def enviar_correo(correo, mensaje):
    try:
        remitente = "matriculaslagranja2@gmail.com"
        clave = "umji dkpr fytx gndj"
        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(remitente, clave)
        
        correo_msg = MIMEMultipart()
        correo_msg["From"] = remitente
        correo_msg["To"] = correo
        correo_msg["Subject"] = "Validación de Documentos"
        correo_msg.attach(MIMEText(mensaje, "plain"))
        
        servidor.sendmail(remitente, correo, correo_msg.as_string())
        servidor.quit()
        print(f"Correo enviado correctamente a {correo}")
    except Exception as e:
        print(f"Error al enviar el correo a {correo}: {e}")
        


def obtener_tiempo():
    hora = time.strftime('%H:%M:%S')
    fecha = time.strftime('%A %d %B %Y')
    texto_hora.config(text=hora)
    texto_fecha.config(text=fecha)
    texto_hora.after(1000, obtener_tiempo)

def agregar_reloj(ventana):
    global texto_hora, texto_fecha
    reloj_frame = Toplevel(ventana)
    reloj_frame.geometry("1300x100")  # Ajusta la posición en la esquina superior derecha
    reloj_frame.overrideredirect(True)  # Eliminar barra de título
    reloj_frame.config(bg='gray')
    reloj_frame.wm_attributes('-transparentcolor', 'gray')
    
    texto_hora = Label(reloj_frame, fg='white', bg='black', font=('Arial', 20, 'bold'))
    texto_hora.pack()
    texto_fecha = Label(reloj_frame, fg='white', bg='black', font=('Arial', 12))
    texto_fecha.pack()
    
    obtener_tiempo()


def iniciar_interfaz():
    global ventana
    
    ventana = Tk()
    
    # Colores suaves para el fondo
    ventana.title("Envío de Mensajes Masivos a WhatsApp y Gmail")
    ventana.geometry("500x400")
    ventana.resizable(False, False)
    ventana.config(bg='#f4f4f9')  # Fondo de color suave (blanco sucio)

    # Establecer fuente personalizada
    font_button = tkFont.Font(family="Segoe UI", size=12, weight="bold")
    font_label = tkFont.Font(family="Segoe UI", size=16, weight="bold")

    # Cambiar íconos
    ventana.iconbitmap("backend/iconos/send.ico") 

    # Cargar iconos y ajustar tamaño
    icono_whatsapp = PhotoImage(file="backend\iconos\wasap.png").subsample(8, 8)
    icono_gmail = PhotoImage(file="backend\iconos\gmaili.png").subsample(8, 8)
    icono_exit = PhotoImage(file="backend\iconos\exit.png").subsample(11, 11)
    
    agregar_reloj(ventana)  # Agregar el reloj en la esquina superior derecha

    # Label con color de fondo claro
    Label(ventana, text="Seleccione una opción", font=font_label, bg='#f4f4f9').pack(pady=10)
    
    # Botones con colores suaves, bordes redondeados y sombras
    Button(ventana, text="Enviar mensajes a WhatsApp", command=lambda: seleccionar_archivo(1, ventana), 
           width=30, height=2, font=font_button, relief="flat", bg='#3498db', fg='white', 
           activebackground='#2980b9').pack(pady=5)
    Button(ventana, text="Enviar mensajes por Gmail", command=lambda: seleccionar_archivo(2, ventana), 
           width=30, height=2, font=font_button, relief="flat", bg='#e74c3c', fg='white', 
           activebackground='#c0392b').pack(pady=5)

    Button(ventana, text="  ""Descargar plantilla WhatsApp", image=icono_whatsapp, compound="left", 
           command=lambda: descargar_plantilla("whatsapp"), width=300, height=50, font=font_button, 
           relief="flat", bg='#2ecc71', fg='white', activebackground='#27ae60').pack(pady=5)
    
    Button(ventana, text="  ""Descargar plantilla Gmail", image=icono_gmail, compound="left", 
           command=lambda: descargar_plantilla("gmail"), width=300, height=50, font=font_button, 
           relief="flat", bg='#f39c12', fg='white', activebackground='#e67e22').pack(pady=5)
    
    Button(ventana, text="Salir", image=icono_exit, compound="left", 
        command=ventana.quit, width=285, height=50, font=font_button, 
        relief="flat", bg='#7f8c8d', fg='white', activebackground='#95a5a6',
        padx=10).pack(pady=2)


    ventana.mainloop()
if __name__ == "__main__":
    iniciar_interfaz()