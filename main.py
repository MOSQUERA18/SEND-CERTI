import pandas as pd
import time
import tkinter as tk
from tkinter import ttk
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
from dotenv import load_dotenv

load_dotenv()
REMITENTE = os.getenv("EMAIL_USER")
CLAVE = os.getenv("PASSWORD")


# Obtener la ruta de Descargas del usuario
def obtener_ruta_descargas():
    if os.name == "nt":  # Windows
        return os.path.join(os.environ["USERPROFILE"], "Downloads")

# Crear carpeta "Reporte de envíos" dentro de Descargas
RUTA_REPORTE = os.path.join(obtener_ruta_descargas(), "Reporte de envíos")
os.makedirs(RUTA_REPORTE, exist_ok=True)

# Archivos de historial
HISTORIAL_WHATSAPP = os.path.join(RUTA_REPORTE, "historial_whatsapp.txt")
HISTORIAL_CORREOS = os.path.join(RUTA_REPORTE, "historial_correos.txt")

def registrar_historial(archivo, destinatario, estado):
    """ Registra el estado del envío en un archivo de historial. """
    with open(archivo, "a", encoding="utf-8") as file:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        file.write(f"{timestamp}, {destinatario}, {estado}\n")



def mostrar_vista_previa(df, ruta_excel, opcion, ventana):
    """ Muestra una vista previa del archivo Excel antes de enviarlo. """
    
    # Crear ventana emergente
    vista_previa = tk.Toplevel(ventana)
    vista_previa.title("Vista previa del archivo")
    vista_previa.geometry("600x400")

    # Etiqueta de instrucciones
    label = tk.Label(vista_previa, text="Verifica los datos antes de enviarlos:", font=("Arial", 12, "bold"))
    label.pack(pady=5)

    # Crear un frame con scroll para mostrar la tabla
    frame_tabla = tk.Frame(vista_previa)
    frame_tabla.pack(fill=tk.BOTH, expand=True)

    # Crear tabla con ttk.Treeview
    tree = ttk.Treeview(frame_tabla, columns=list(df.columns), show="headings")

    # Agregar encabezados de columnas
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)

    # Agregar filas de la tabla (solo las primeras 10 filas para no saturar la ventana)
    for _, row in df.head(10).iterrows():
        tree.insert("", tk.END, values=list(row))

    tree.pack(fill=tk.BOTH, expand=True)

    # Botones para aceptar o cancelar el envío
    def aceptar_envio():
        vista_previa.destroy()
        if opcion == 1:
            threading.Thread(target=enviar_mensajes_whatsapp, args=(ruta_excel, ventana)).start()
        elif opcion == 2:
            threading.Thread(target=procesar_correo, args=(ruta_excel, ventana)).start()

    tk.Button(vista_previa, text="Aceptar y Enviar", command=aceptar_envio, bg="#2ecc71", fg="white").pack(side=tk.LEFT, padx=20, pady=10)
    tk.Button(vista_previa, text="Cancelar", command=vista_previa.destroy, bg="#e74c3c", fg="white").pack(side=tk.RIGHT, padx=20, pady=10)




def descargar_plantilla(tipo):
    """ Descarga una plantilla de Excel en la carpeta de Descargas del usuario """
    ruta_descargas = obtener_ruta_descargas()
    os.makedirs(ruta_descargas, exist_ok=True)

    if tipo == "whatsapp":
        ruta_guardado = os.path.join(ruta_descargas, "plantilla_whatsapp.xlsx")
        columnas = ["Nombre", "Numero_Telefono", "Remitente", "Mensaje"]
    else:
        ruta_guardado = os.path.join(ruta_descargas, "plantilla_correo.xlsx")
        columnas = [
            "Correo", "Correo_Remitente", "Clave_Aplicacion",
            "FICHA", "NIVEL DE FORMACION", "NOMBRES Y APELLIDOS", "CERTIFICADO ETAPA PRODUCTIVA",
            "EVALUACION PARCIAL", "EVALUACION FINAL", "TYT (TECNOLOGOS)",
            "CERTIFICADO VIGENCIA", "CERTIFICADO AGENCIA DE EMPLEO SENA", "CARNET DESTRUIDO",
            "PAZ Y SALVO ACADEMICO ADMINISTRATIVO"
        ]

    df = pd.DataFrame(columns=columnas)
    df.to_excel(ruta_guardado, index=False)
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

def generar_mensaje(nombre, remitente, mensaje_base, datos_extra=None):
    saludo = obtener_saludo()

    if not isinstance(mensaje_base, str) or mensaje_base.strip() == "":
        mensaje_base = "📌 *Por favor, revisa esta información importante.*"

    mensaje = (f"👋 *Hola {nombre}*, {saludo}.\n\n"
               f"✍️ *Escribo* desde el *Centro Agropecuario La Granja.*\n\n"
               f"📢 {mensaje_base}\n")

    # Agregar datos extra si existen
    if datos_extra:
        for k, v in datos_extra.items():
            if pd.notna(v) and str(v).strip() != "":
                mensaje += f"\n📄 *{k.title()}*: {v}"

    mensaje += "\n\nGracias por tu atención. ✅"
    return mensaje


def mostrar_aviso():
    ventana_aviso = tk.Toplevel()
    ventana_aviso.iconbitmap("C:\\send-certi\\backend\\iconos\\send.ico")
    ventana_aviso.title("WhatsApp Web")
    ventana_aviso.geometry("350x100")
    
    ventana_aviso.update_idletasks()
    ancho_pantalla = ventana_aviso.winfo_screenwidth()
    alto_pantalla = ventana_aviso.winfo_screenheight()
    x_pos = ancho_pantalla - 360
    y_pos = alto_pantalla - 160
    ventana_aviso.geometry(f"350x100+{x_pos}+{y_pos}")

    tk.Label(ventana_aviso, text="📢 Inicie sesión en WhatsApp Web\n y haga clic en 'Listo' para continuar.", font=("Arial", 11), fg="black").pack(pady=10)

    listo_event = threading.Event()

    def on_listo():
        listo_event.set()
        ventana_aviso.destroy()

    tk.Button(ventana_aviso, text="Listo", command=on_listo, font=("Arial", 10, "bold"), bg="#2ecc71", fg="white").pack(pady=5)
    ventana_aviso.attributes("-topmost", True)
    ventana_aviso.transient()
    ventana_aviso.grab_set()
    ventana_aviso.wait_window()
    return listo_event


def enviar_mensajes_whatsapp(ruta_excel, ventana):
    """ Envía mensajes de WhatsApp y registra el historial de envíos. """
    try:
        datos = pd.read_excel(ruta_excel)
         # ✅ Limpiar los números de teléfono para eliminar espacios
        if "Numero_Telefono" in datos.columns:
            datos["Numero_Telefono"] = datos["Numero_Telefono"].astype(str).str.replace(r"\D", "", regex=True)

    except Exception as e:
        print(f"❌ Error al cargar el archivo Excel: {e}")
        return

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get("https://web.whatsapp.com/")

    # ✅ Mostrar el aviso una sola vez y esperar a que se cierre
    listo_event = mostrar_aviso()
    listo_event.wait()

    for _, fila in datos.iterrows():
        nombre = fila.get("Nombre")
        numero_telefono = fila.get("Numero_Telefono")
        if not numero_telefono.isdigit():  # ✅ Validar que solo tenga números
                print(f"❌ Número inválido: {numero_telefono}")
                registrar_historial(HISTORIAL_WHATSAPP, numero_telefono, "NÚMERO INVÁLIDO")
                continue
        remitente = fila.get("Remitente")
        mensaje_base = fila.get("Mensaje")

        if pd.isna(nombre) or pd.isna(numero_telefono) or pd.isna(remitente):
            registrar_historial(HISTORIAL_WHATSAPP, numero_telefono, "DATOS INCOMPLETOS")
            continue

        mensaje = generar_mensaje(nombre, remitente, mensaje_base)
        mensaje_codificado = quote(mensaje)
        url = f"https://web.whatsapp.com/send?phone=+57{int(numero_telefono)}&text={mensaje_codificado}"
        driver.get(url)

        try:
            # ✅ Espera a que aparezca el botón de enviar
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            ).click()
            time.sleep(2)

            print(f"✅ Mensaje enviado a {numero_telefono}")
            registrar_historial(HISTORIAL_WHATSAPP, numero_telefono, "ENVIADO")

        except Exception as e:
            print(f"❌ Error al enviar mensaje a {numero_telefono}: {e}")
            registrar_historial(HISTORIAL_WHATSAPP, numero_telefono, f"NO ENVIADO - {str(e)}")

    driver.quit()
    print("✅ Proceso de envío de WhatsApp finalizado.")



def validar_archivo_excel(ruta_excel, opcion):
    """ Verifica que el archivo tenga el formato correcto según la opción seleccionada. """
    try:
        df = pd.read_excel(ruta_excel)

        # Verificar que el archivo no esté vacío
        if df.empty:
            messagebox.showerror("Error", "⚠️ El archivo Excel está vacío. Carga un archivo válido.")
            return False


        # Definir las columnas esperadas según el tipo de mensaje
        columnas_esperadas = {
            1: {"Nombre", "Numero_Telefono", "Remitente", "Mensaje"},  # WhatsApp
            # opción 2 - correo 
            2: {
                "NOMBRES Y APELLIDOS", "Correo", "Correo_Remitente", "Clave_Aplicacion",
                "FICHA", "NIVEL DE FORMACION",
                "CERTIFICADO ETAPA PRODUCTIVA", "EVALUACION PARCIAL",
                "EVALUACION FINAL", "TYT (TECNOLOGOS)", "CERTIFICADO VIGENCIA",
                "CERTIFICADO AGENCIA DE EMPLEO SENA", "CARNET DESTRUIDO",
                "PAZ Y SALVO ACADEMICO ADMINISTRATIVO"
                }}.get(opcion, None)

        if columnas_esperadas is None:
            messagebox.showerror("Error", "⚠️ Opción de envío no válida.")
            return False

        # Obtener las columnas reales del archivo
        columnas_actuales = set(df.columns)

        # Verificar si faltan columnas
        columnas_faltantes = columnas_esperadas - columnas_actuales
        if columnas_faltantes:
            messagebox.showerror("Error", f"⚠️ Faltan columnas en el archivo: {', '.join(columnas_faltantes)}")
            return False

        # ⚠️ Solo realizar limpieza si es WhatsApp
        if opcion == 1 and "Numero_Telefono" in df.columns:
            df["Numero_Telefono"] = df["Numero_Telefono"].astype(str).str.replace(r'\D', '', regex=True)
            df = df[df["Numero_Telefono"].str.len() >= 10]  # Asegurar que tenga al menos 10 dígitos

        # ⚠️ Solo validar correos si es Gmail
        if opcion == 2 and "Correo" in df.columns:
            df = df[df["Correo"].str.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", na=False)]

        # Verificar si después de la limpieza queda algún dato válido
        if df.empty:
            messagebox.showerror("Error", "⚠️ Todos los registros del archivo son inválidos.")
            return False

        return df  # Retorna el DataFrame limpio si es válido

    except Exception as e:
        messagebox.showerror("Error", f"⚠️ No se pudo leer el archivo Excel. Verifica que sea un archivo válido.\n\nError: {e}")
        return False

def seleccionar_archivo(opcion, ventana):
    """ Permite seleccionar solo archivos Excel (.xlsx) y valida su contenido. """
    ruta_excel = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )

    if not ruta_excel:
        return  # Si el usuario cancela la selección, no hacer nada

    df_validado = validar_archivo_excel(ruta_excel, opcion)

    if df_validado is False:
        return  # Si hay un error, salir sin continuar

    # Mostrar la vista previa con el DataFrame validado
    mostrar_vista_previa(df_validado, ruta_excel, opcion, ventana)


PLANTILLA_MENSAJE_HTML = """\
<html>
<head>
  <style>
    body {{
      font-family: 'Segoe UI', Arial, sans-serif;
      font-size: 15px;
      line-height: 1.6;
      color: #2c3e50;
    }}
    h2 {{
      color: #e74c3c;
      font-weight: bold;
      font-size: 20px;
    }}
    .important {{
      background-color: #f9f1f1;
      border-left: 5px solid #e74c3c;
      padding: 10px 15px;
      margin: 20px 0;
    }}
    .alert {{
      color: red;
      font-weight: bold;
    }}
    ul {{
      padding-left: 20px;
    }}
    a {{
      color: #2980b9;
      text-decoration: none;
    }}
    a:hover {{
      text-decoration: underline;
    }}
  </style>
</head>
<body>

<h2>📌 Inconsistencias en Documentación - Proceso de Certificación</h2>

<p>Estimado/a <strong>{nombres_apellidos}</strong><br>
<strong>Nivel de Formación:</strong> {nivel_formacion}</p>

<div class="important">
  <p>
    Tras la revisión de su documentación para el proceso de certificación, se han identificado <strong>inconsistencias o documentos pendientes</strong> que deben ser corregidos o entregados con prontitud.
  </p>
</div>

<p><strong>📋 Estado de sus documentos:</strong></p>
<ul>
  <li>📄 Certificado Etapa Productiva: {certificado_etapa}</li>
  <li>📄 Evaluación Parcial: {evaluacion_parcial}</li>
  <li>📄 Evaluación Final: {evaluacion_final}</li>
  <li>📄 Prueba TYT (Tecnólogos): {tyt}</li>
  <li>📄 Certificado de Vigencia: {certificado_vigencia}</li>
  <li>📄 Certificado Agencia Pública de Empleo SENA: {certificado_agencia}</li>
  <li>📄 Evidencia Carnet Destruido o Constancia: {carnet_destruido}</li>
  <li>📄 Paz y Salvo Académico-Administrativo: {paz_salvo}</li>
</ul>

<div class="important">
  <p>📎 <strong>Formulario para cargar documentos:</strong><br>
  <a href="https://forms.gle/iukhL53YvGPDy1SU8">https://forms.gle/iukhL53YvGPDy1SU8</a></p>
</div>

<p><strong>📝 Instrucciones para el archivo PDF:</strong></p>
<ul>
  <li>Debe contener <strong>todos los documentos</strong> mencionados arriba, en el mismo orden.</li>
  <li>El archivo debe estar en formato PDF.</li>
  <li>Nombre del archivo: <em>[FICHA] [NOMBRES Y APELLIDOS EN MAYÚSCULA]</em><br>
  <small>Ejemplo: <strong>12345 JUANCHO DAVID VALDEZ CORTEZ</strong></small></li>
</ul>

<p class="important">
  ⚠️ Su pronta respuesta es fundamental para evitar retrasos en el proceso de certificación.
</p>

<p>Gracias por su atención.<br>
Atentamente,<br>
<strong>Yair Cárdenas</strong><br>
Equipo de Registro<br>
Centro Agropecuario “La Granja” – SENA, Espinal (Tolima)</p>

</body>
</html>
"""


def resaltar_mal(valor):
    if isinstance(valor, str) and valor.strip().upper() == "MAL":
        return '<span style="color:red;font-weight:bold;">MAL</span>'
    return valor if pd.notna(valor) else ""



def generar_mensaje_con_plantilla_html(fila):
    return PLANTILLA_MENSAJE_HTML.format(
        nombres_apellidos=fila.get("NOMBRES Y APELLIDOS", ""),
        nivel_formacion=resaltar_mal(fila.get("NIVEL DE FORMACION", "")),
        certificado_etapa=resaltar_mal(fila.get("CERTIFICADO ETAPA PRODUCTIVA", "")),
        evaluacion_parcial=resaltar_mal(fila.get("EVALUACION PARCIAL", "")),
        evaluacion_final=resaltar_mal(fila.get("EVALUACION FINAL", "")),
        tyt=resaltar_mal(fila.get("TYT (TECNOLOGOS)", "")),
        certificado_vigencia=resaltar_mal(fila.get("CERTIFICADO VIGENCIA", "")),
        certificado_agencia=resaltar_mal(fila.get("CERTIFICADO AGENCIA DE EMPLEO SENA", "")),
        carnet_destruido=resaltar_mal(fila.get("CARNET DESTRUIDO", "")),
        paz_salvo=resaltar_mal(fila.get("PAZ Y SALVO ACADEMICO ADMINISTRATIVO", ""))
    )





def procesar_correo(ruta_excel, ventana):
    mostrar_cargando(ventana)
    datos = cargar_datos(ruta_excel)
    if datos is None:
        ocultar_cargando()
        return

    columnas_requeridas = [
        "NOMBRES Y APELLIDOS", "Correo", "Correo_Remitente", "Clave_Aplicacion",
        "FICHA", "NIVEL DE FORMACION",
        "CERTIFICADO ETAPA PRODUCTIVA", "EVALUACION PARCIAL",
        "EVALUACION FINAL", "TYT (TECNOLOGOS)", "CERTIFICADO VIGENCIA",
        "CERTIFICADO AGENCIA DE EMPLEO SENA", "CARNET DESTRUIDO",
        "PAZ Y SALVO ACADEMICO ADMINISTRATIVO"
    ]

    for _, fila in datos.iterrows():
        # Validar datos básicos
        if any(pd.isna(fila.get(col)) or str(fila.get(col)).strip() == "" for col in columnas_requeridas[:4]):
            registrar_historial(HISTORIAL_CORREOS, fila.get("Correo", "DESCONOCIDO"), "DATOS INCOMPLETOS")
            continue

        mensaje = generar_mensaje_con_plantilla_html(fila)
        enviar_correo(
            correo=fila["Correo"],
            mensaje=mensaje,
            remitente=fila["Correo_Remitente"],
            clave_aplicacion=fila["Clave_Aplicacion"]
        )

    ocultar_cargando()
    messagebox.showinfo("Finalizado", "Correos enviados correctamente.")


def enviar_correo(correo, mensaje, remitente, clave_aplicacion):
    """ Envía un correo con credenciales dinámicas obtenidas del Excel. """
    try:
        if not remitente or not clave_aplicacion:
            print("⚠️ Error: Credenciales de correo no encontradas.")
            registrar_historial(HISTORIAL_CORREOS, correo, "ERROR: Credenciales no configuradas")
            return

        if not correo or "@" not in correo:
            print(f"⚠️ Error: Dirección de correo no válida ({correo})")
            registrar_historial(HISTORIAL_CORREOS, correo, "CORREO NO ENCONTRADO")
            return

        if not mensaje or mensaje.strip() == "":
            print(f"⚠️ Error: Mensaje vacío para {correo}. No se enviará el correo.")
            registrar_historial(HISTORIAL_CORREOS, correo, "NO ENVIADO - MENSAJE VACÍO")
            return

        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(remitente, clave_aplicacion)

        correo_msg = MIMEMultipart()
        correo_msg["From"] = remitente
        correo_msg["To"] = correo
        correo_msg["Subject"] = "Información Importante"
        correo_msg.attach(MIMEText(mensaje, "html"))
        servidor.sendmail(remitente, correo, correo_msg.as_string())
        servidor.quit()

        print(f"📧 Correo enviado correctamente a {correo} desde {remitente}")
        registrar_historial(HISTORIAL_CORREOS, correo, "ENVIADO")

    except Exception as e:
        print(f"❌ Error al enviar el correo a {correo}: {e}")
        registrar_historial(HISTORIAL_CORREOS, correo, f"NO ENVIADO - {str(e)}")




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
    
    ventana = tk.Tk()
    
    # Colores suaves para el fondo
    ventana.title("Envío de Mensajes Masivos Via WhatsApp y Gmail")
    ventana.geometry("600x460")
    ventana.resizable(False, False)
    ventana.config(bg='#2c3e50')  # Fondo oscuro

    # Establecer fuente personalizada
    font_button = tkFont.Font(family="Segoe UI", size=12, weight="bold")
    font_label = tkFont.Font(family="Segoe UI", size=16, weight="bold")
    font_footer = tkFont.Font(family="Arial", size=10, weight="bold")

    # Cambiar íconos
    ventana.iconbitmap("C:\\send-certi\\backend\\iconos\\send.ico") 

    # Cargar iconos y ajustar tamaño
    icono_whatsapp = PhotoImage(file="C:\\send-certi\\backend\\iconos\\wasap.png").subsample(8, 8)
    icono_gmail = PhotoImage(file="C:\\send-certi\\backend\\iconos\\gmaili.png").subsample(8, 8)
    icono_exit = PhotoImage(file="C:\\send-certi\\backend\\iconos\\exit.png").subsample(11, 11)

    # Frame para el contenido principal
    frame_principal = tk.Frame(ventana, bg='#34495e', padx=20, pady=20)
    frame_principal.pack(expand=True, fill='both')

    # Label con color de fondo claro
    Label(frame_principal, text="Seleccione una opción", font=font_label, bg='#34495e', fg='white').pack(pady=10)
    
    # Botones con colores suaves, bordes redondeados y sombras
    Button(frame_principal, text="Enviar mensajes a WhatsApp", command=lambda: seleccionar_archivo(1, ventana), 
           width=30, height=2, font=font_button, relief="flat", bg='#3498db', fg='white', 
           activebackground='#2980b9').pack(pady=5)
    
    Button(frame_principal, text="Enviar mensajes por Gmail", command=lambda: seleccionar_archivo(2, ventana), 
           width=30, height=2, font=font_button, relief="flat", bg='#e74c3c', fg='white', 
           activebackground='#c0392b').pack(pady=5)
 
    Button(frame_principal, text="   " "Descargar plantilla WhatsApp", image=icono_whatsapp, compound="left", 
           command=lambda: descargar_plantilla("whatsapp"), width=300, height=50, font=font_button, 
           relief="flat", bg='#008B8B', fg='white', activebackground='#27ae60').pack(pady=5)
    
    Button(frame_principal, text="   " "Descargar plantilla Gmail", image=icono_gmail, compound="left", 
           command=lambda: descargar_plantilla("Gmail"), width=300, height=50, font=font_button, 
           relief="flat", bg='#f39c12', fg='white', activebackground='#e67e22').pack(pady=5)
    
    Button(frame_principal, text="Salir", image=icono_exit, compound="left", 
        command=ventana.quit, width=285, height=50, font=font_button, 
        relief="flat", bg='#7f8c8d', fg='white', activebackground='#95a5a6',
        padx=10).pack(pady=2)

    # 📌 Pie de página interactivo 📌
    footer = Label(ventana, text="Desarrollado por Marlon Mosquera ADSO 2671143", font=font_footer, 
                   bg='#2c3e50', fg='white', cursor="hand2")
    footer.pack(side="bottom", pady=5)

    # 🎨 Función para cambiar color en hover
    def on_enter(e):
        footer.config(fg="#3498db")  # Azul

    def on_leave(e):
        footer.config(fg='white')  # Blanco

    # Asociar eventos hover
    footer.bind("<Enter>", on_enter)
    footer.bind("<Leave>", on_leave)

    ventana.mainloop()

if __name__ == "__main__":
    iniciar_interfaz()