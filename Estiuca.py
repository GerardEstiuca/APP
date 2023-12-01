import tkinter as tk
from tkinter import ttk
import pyautogui
from PIL import Image, ImageTk, ImageGrab
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests
from msal import PublicClientApplication
import threading
from msal import ConfidentialClientApplication
from io import BytesIO
import pandas as pd
import json
from datetime import datetime
import tkinter.messagebox as msgbox
import io
import base64
import os
import tempfile
import pandas as pd
from openpyxl import load_workbook
import openpyxl


def penjar_QR_github():
    global direccio, direction
    
    
    imagen = obtener_imagen_portapapeles2()
    if imagen is None:
        print("No hay imagen en el portapapeles.")
        return

    # Obtener la hora actual en formato adecuado
    now = datetime.now()
    formatted_time = now.strftime("%d%m%Y_%H%M")

    # Nombre del archivo
    file_name = f"{pais.get()}_CodiQR_{formatted_time}.png"
    print(file_name)
    # Guardar la imagen en un archivo temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_file:
        imagen.save(temp_file, format='png')
        file_path = temp_file.name

    # Subir a GitHub
    token = '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'  # Asegúrate de que este token sea seguro y no esté expuesto públicamente
    username = 'GerardEstiuca'
    repo = 'EstiuFLY'
    branch = 'main'
    path = f"Codis QR/{empresa.get()}/{dia}/{file_name}"  # Cambia esto por la ruta donde quieres subir el archivo en el repositorio

    url = f'https://api.github.com/repos/{username}/{repo}/contents/{path}?ref={branch}'
    headers = {'Authorization': f'token {token}'}
    with open(file_path, 'rb') as file:
        content = file.read()
    data = {
        'message': f'Adding QR Code {file_name}',
        'branch': branch,
        'content': base64.b64encode(content).decode('utf-8')
    }
    response = requests.put(url, headers=headers, json=data)
    if response.status_code == 201:
        print("Archivo subido con éxito.")
        direccio = response.json()['content']['download_url']
    else:
        print(f"Error al subir el archivo: {response.text}")
   
def obtener_imagen_portapapeles2():
    try:
        # Usar ImageGrab para obtener la imagen del portapapeles
        imagen_portapapeles = ImageGrab.grabclipboard()
        return imagen_portapapeles
    except Exception as e:
        print(f"Error al obtener la imagen del portapapeles: {e}")
        return None

def actualizar_imagen_frame():
    global foto
    imagen = obtener_imagen_portapapeles2()
    if imagen:
        # Redimensionar la imagen para que se ajuste al tamaño deseado (30x30)
        imagen = imagen.resize((200, 200), Image.Resampling.LANCZOS)
        foto = ImageTk.PhotoImage(imagen)
    

# Función para leer archivo excel de la SIM's
def leer_archivo_SIMS():
    global unique_values, df
    # Tus credenciales de Azure AD
    client_id = '7fc699bf-de32-4f0c-b75b-1bd6c5864c53'
    client_secret = 'Chq8Q~MIlZyug9JJ3N-g3urWIhSax1IY-BgWYdtI'
    tenant_id = '0450f773-f698-4e8f-ad64-a3477d6c579a'
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    scopes = ['https://graph.microsoft.com/.default']

    # Crear una instancia de la aplicación confidencial
    app2 = ConfidentialClientApplication(client_id, authority=authority_url, client_credential=client_secret)

    # Obtener el token de acceso
    result = app2.acquire_token_for_client(scopes=scopes)
    token = result['access_token']

    # Asegúrate de reemplazar con el ID de usuario de OneDrive o el ID del sitio de SharePoint
    user_id = 'gerard@estiuca.cat'

    # Formar la URL para acceder al archivo en OneDrive
    file_url = f'https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/Seguiment Propostes/RELACIÓ SIMs.xlsm:/content'

    # Headers para la solicitud
    headers = {'Authorization': f'Bearer {token}'}

    # Realizar la solicitud GET para obtener el archivo
    response = requests.get(file_url, headers=headers)

   
    # Procesar los datos
    data = BytesIO(response.content)
    df = pd.read_excel(data, engine='openpyxl', sheet_name='GENERAL')

    # Obtener valores únicos de la columna E
    unique_values = df['CLIENT'].dropna().unique().tolist()



# Función para leer archivo excel de EstiuFLY
def leer_archivo_Estiufly():
    global unique_values, resultado, df
    # Tus credenciales de Azure AD
    client_id = '7fc699bf-de32-4f0c-b75b-1bd6c5864c53'
    client_secret = 'Chq8Q~MIlZyug9JJ3N-g3urWIhSax1IY-BgWYdtI'
    tenant_id = '0450f773-f698-4e8f-ad64-a3477d6c579a'
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    scopes = ['https://graph.microsoft.com/.default']

    # Crear una instancia de la aplicación confidencial
    app2 = ConfidentialClientApplication(client_id, authority=authority_url, client_credential=client_secret)

    # Obtener el token de acceso
    result = app2.acquire_token_for_client(scopes=scopes)
    token = result['access_token']

    # Asegúrate de reemplazar con el ID de usuario de OneDrive o el ID del sitio de SharePoint
    user_id = 'gerard@estiuca.cat'

    # Formar la URL para acceder al archivo en OneDrive
    file_url = f'https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/ESTIUFLY/Control Comandes.xlsx:/content'

    # Headers para la solicitud
    headers = {'Authorization': f'Bearer {token}'}

    # Realizar la solicitud GET para obtener el archivo
    response = requests.get(file_url, headers=headers)

   
    # Procesar los datos
    data = BytesIO(response.content)
    df = pd.read_excel(data, engine='openpyxl', sheet_name='Llistat')

     # Encontrar la última celda con valor en la columna B
    last_cell_value = None
    for cell in df.iloc[:, 0]:
        if pd.notna(cell):
            last_cell_value = cell

    # Almacenar el último valor encontrado
    unique_values = last_cell_value

    # Ya no se suma 1, sino que se almacena el último valor encontrado
    resultado = unique_values + 1

# Función para escribir en archivo excel de EstiuFLY
def escribir_archivo_Estiufly():
    global unique_values, resultado, df
    # Tus credenciales de Azure AD
    client_id = '7fc699bf-de32-4f0c-b75b-1bd6c5864c53'
    client_secret = 'Chq8Q~MIlZyug9JJ3N-g3urWIhSax1IY-BgWYdtI'
    tenant_id = '0450f773-f698-4e8f-ad64-a3477d6c579a'
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    scopes = ['https://graph.microsoft.com/.default']

    # Crear una instancia de la aplicación confidencial
    app2 = ConfidentialClientApplication(client_id, authority=authority_url, client_credential=client_secret)

    # Obtener el token de acceso
    result = app2.acquire_token_for_client(scopes=scopes)
    token = result['access_token']

    # Asegúrate de reemplazar con el ID de usuario de OneDrive o el ID del sitio de SharePoint
    user_id = 'gerard@estiuca.cat'

    # Formar la URL para acceder al archivo en OneDrive
    file_url = f'https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/ESTIUFLY/Control Comandes.xlsx:/content'

    # Headers para la solicitud
    headers = {'Authorization': f'Bearer {token}'}

    # Realizar la solicitud GET para obtener el archivo
    response = requests.get(file_url, headers=headers)

    # Leer el archivo en un libro de openpyxl
    data = BytesIO(response.content)
    book = openpyxl.load_workbook(data)
    sheet = book['Llistat']  # Asegúrate de que el nombre de la hoja sea correcto


    # Encontrar la primera fila vacía en la hoja
    first_empty_row = 2  # Comenzar desde la fila 2
    while sheet[f'A{first_empty_row}'].value is not None:
        first_empty_row += 1
    
    print(first_empty_row)
    print(refcomanda.get())
    print(data1.get())

    # Asegúrate de que las variables sean Entry widgets de Tkinter y estén correctamente inicializadas
    sheet[f'A{first_empty_row}'] = float(refcomanda.get())
    sheet[f'B{first_empty_row}'] = data1.get()
    sheet[f'C{first_empty_row}'] = empresa.get()
    sheet[f'D{first_empty_row}'] = pais.get()
    sheet[f'E{first_empty_row}'] = float(vigencia.get())
    sheet[f'F{first_empty_row}'] = float(gigas.get())
    sheet[f'G{first_empty_row}'] = float(cost.get())
    sheet[f'H{first_empty_row}'] = float(cost.get())*3
    sheet[f'I{first_empty_row}'] = destinatario1.get()
    sheet[f'J{first_empty_row}'] = cobertura.get()
    sheet[f'K{first_empty_row}'] = recarga.get()
    sheet[f'L{first_empty_row}'] = apn.get()

    # Guardar los cambios en el objeto BytesIO
    data = BytesIO()
    book.save(data)
    data.seek(0)  # Rebobinar el objeto BytesIO al principio

    # Preparar los headers para la solicitud PUT
    headers_put = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}

    # Realizar la solicitud PUT para actualizar el archivo en OneDrive
    put_response = requests.put(file_url, headers=headers_put, data=data.getvalue())

    # Verificar si la subida fue exitosa
    if put_response.status_code == 200 or put_response.status_code == 201:
        print("Archivo actualizado con éxito en OneDrive")
    else:
        print("Error al actualizar el archivo:", put_response.text)

# Función para leer la plantilla HTML
def leer_plantilla_html(ruta_archivo):
    with open(ruta_archivo, 'r', encoding='utf-8') as archivo:
        contenido = archivo.read()
    return contenido

# Crear tooltip numero en ventana_portafixe
def mostrar_tooltip(event):
    global ventana_portaFixe, numero
    # Crea el tooltip como un Label en ventana_portaFixe
    tooltip = ttk.Label(ventana_portaFixe, text="Si hi ha més d'una numeració, separales amb un ;  ", background="white", relief="solid")

    # Obtén la posición del widget numero dentro de la ventana
    x = numero.winfo_x()
    y = numero.winfo_y()

    # Posiciona el tooltip debajo del widget numero
    tooltip.place(x=x, y=y + numero.winfo_height())

    # Programa el tooltip para que se destruya después de 3 segundos
    ventana_portaFixe.after(3000, tooltip.destroy)

# Crear tooltip hora inici en ventana_fibres
def mostrar_tooltip_fibra(event):
    global ventana_fibres, hora
    # Crea el tooltip como un Label en ventana_portaFixe
    tooltip = ttk.Label(ventana_fibres, text="Introduir hora inici instal·lació", background="white", relief="solid")

    # Obtén la posición del widget numero dentro de la ventana
    x = hora.winfo_x()
    y = hora.winfo_y()

    # Posiciona el tooltip debajo del widget numero
    tooltip.place(x=x, y=y + hora.winfo_height())

    # Programa el tooltip para que se destruya después de 3 segundos
    ventana_fibres.after(3000, tooltip.destroy)

# Crear tooltip hora final en ventana_fibres
def mostrar_tooltip_fibra2(event):
    global ventana_fibres, hora2
    # Crea el tooltip como un Label en ventana_portaFixe
    tooltip = ttk.Label(ventana_fibres, text="Introduir hora final instal·lació", background="white", relief="solid")

    # Obtén la posición del widget numero dentro de la ventana
    x = hora2.winfo_x()
    y = hora2.winfo_y()

    # Posiciona el tooltip debajo del widget numero
    tooltip.place(x=x, y=y + hora2.winfo_height())

    # Programa el tooltip para que se destruya después de 3 segundos
    ventana_fibres.after(3000, tooltip.destroy)



# Función para enviar el correo electrónico
def enviar_correo(destinatario1):
            global plantilla_html
            seleccion = combo_asunto.get()
            if seleccion == "welcome":
                print(f"Enviar email a {destinatario1} con asunto y plantilla de 'welcome'.")

                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"

                # Crear el objeto del mensaje
                msg = MIMEMultipart()
                msg['From'] = direccion_correo
                msg['To'] = destinatario1
                msg['Subject'] = "Benvingut a ESTIUCA!"
                msg['Bcc'] = "estiuca@estiuca.cat"

                # Cuerpo del mensaje con la plantilla HTML
                cuerpo_mensaje = leer_plantilla_html("C:/Users/camac/Desktop/Recusos/WELCOME.html")
                msg.attach(MIMEText(cuerpo_mensaje, 'html'))

                # Iniciar sesión en el servidor y enviar el correo
                try:
                    server = smtplib.SMTP('smtp.office365.com', 587)  # Ajusta según tu proveedor
                    server.starttls()
                    server.login(direccion_correo, password_correo)
                    server.send_message(msg)
                    server.quit()
                    print("Correu enviat satisfactòriament!!")
                    # Cerrar la ventana emergente
                    ventana_welcome.destroy()
                    # Borrar el valor seleccionado en el Combobox
                    combo_asunto.set("")

                except Exception as e:
                    print(f"Error al enviar el correu: {e}")

            elif seleccion == "Enviar documentació":
                print(f"Enviar email a {destinatario1} con asunto y plantilla 'documentació'.")
                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"

                # Crear el objeto del mensaje
                msg = MIMEMultipart()
                msg['From'] = direccion_correo
                msg['To'] = destinatario1
                msg['Subject'] = "Sol·licitud de documentació, ESTIUCA"
                msg['Bcc'] = "estiuca@estiuca.cat"

                # Cuerpo del mensaje con la plantilla HTML
                cuerpo_mensaje = leer_plantilla_html("C:/Users/camac/Desktop/Recusos/DOCUMENTACIÓ.html")
                msg.attach(MIMEText(cuerpo_mensaje, 'html'))

                # Iniciar sesión en el servidor y enviar el correo
                try:
                    server = smtplib.SMTP('smtp.office365.com', 587)  # Ajusta según tu proveedor
                    server.starttls()
                    server.login(direccion_correo, password_correo)
                    server.send_message(msg)
                    server.quit()
                    print("Correu enviat satisfactòriament!!")
                    # Cerrar la ventana emergente
                    ventana_documentació.destroy()
                    # Borrar el valor seleccionado en el Combobox
                    combo_asunto.set("")

                except Exception as e:
                    print(f"Error al enviar el correu: {e}")

            elif seleccion == "Instal·lació de Fibra":
                print(f"Enviar email a {destinatario1} con asunto y plantilla 'fibres'.")
                
                # Obtener los diferentes valores
                client_valor = client.get()
                dia_valor = dia.get()
                direcció_valor = direcció.get()
                hora_valor = hora.get()
                hora2_valor = str(hora2.get()).strip()
                print (hora2_valor)
                # Suponiendo que dia_valor es una cadena con la fecha
                dia_valor = dia_valor.strip()  # Eliminar espacios en blanco al principio y al final
                dia_valor = dia_valor.strip().replace(" ", "")
                hora_valor = str(hora_valor).strip()
                hora2_valor = str(hora2_valor).strip()
                print (hora_valor,hora2_valor)


                # Si hora_valor es solo una cifra, formatearlo como 'HH:00'
                if len(hora_valor) == 1:
                    hora_valor = f"0{hora_valor}:00"
                elif len(hora_valor) == 2 and hora_valor.isdigit():
                    hora_valor = f"{hora_valor}:00"
                print(hora_valor)

                 # Asegurarse de que hora_valor2 tenga el formato correcto
                if len(hora2_valor) == 1:
                    hora2_valor = f"0{hora2_valor}:00"  # Añadir cero y minutos
                elif len(hora2_valor) == 2 and hora2_valor.isdigit():
                    hora2_valor = f"{hora2_valor}:00"  # Añadir minutos

                # Convertir dia_valor y hora_valor en objetos datetime
                try:
                    fecha = datetime.strptime(dia_valor.strip(), '%d/%m/%Y')
                    hora_completa = datetime.strptime(hora_valor, '%H:%M')
                    hora_final=datetime.strptime(hora2_valor, '%H:%M')

                    # Combinar fecha y hora
                    fecha_hora_completa = datetime.combine(fecha.date(), hora_completa.time())
                    fecha_hora_final=datetime.combine(fecha.date(), hora_final.time())
                    fecha_hora_iso = fecha_hora_completa.strftime('%Y-%m-%dT%H:%M:%S')
                    fecha_hora2_iso= fecha_hora_final.strftime('%Y-%m-%dT%H:%M:%S')
                    print(fecha_hora_iso, fecha_hora2_iso)

                except ValueError as e:
                    print("Error en la conversión de fecha y hora:", e)
                    return


                
                
                # Combinar fecha y hora en formato ISO 8601
                inicio_iso = fecha_hora_iso
                final_iso = fecha_hora2_iso  

                
                # Credenciales y detalles de la aplicación
                client_id = '7fc699bf-de32-4f0c-b75b-1bd6c5864c53'
                client_secret = 'Chq8Q~MIlZyug9JJ3N-g3urWIhSax1IY-BgWYdtI'
                tenant_id = '0450f773-f698-4e8f-ad64-a3477d6c579a'
                authority_url = f'https://login.microsoftonline.com/{tenant_id}'
                scopes = ['https://graph.microsoft.com/.default']
                grant_type = 'client_credentials'

                # Asegúrate de reemplazar 'user_email_or_id' con el correo electrónico o ID del usuario de Outlook
                user_email_or_id = 'gerard@estiuca.cat' 


                # Obtener el token de acceso
                url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
                payload = {
                    'client_id': client_id,
                    'scope': scopes,
                    'client_secret': client_secret,
                    'grant_type': grant_type
                }
                response = requests.post(url, data=payload)
                access_token = response.json()['access_token']

                # Crear el evento
                event_payload = {
                    "subject": "Instal·lació Fibra " + client_valor,
                    "body": {
                        "contentType": "HTML",
                        "content": "porva"
                    },
                    "start": {
                        "dateTime": inicio_iso,
                        "timeZone": "UTC"
                    },
                    "end": {
                        "dateTime": final_iso,
                        "timeZone": "UTC"
                    },
                    "attendees": [
                        {
                            "emailAddress": {
                                "address": "estiuca@estiuca.cat",
                                "name": "estiuca"
                            },
                            "type": "required"
                        }
                    ]
                }

                # Enviar la solicitud para crear el evento
                headers = {
                    'Authorization': f'Bearer {access_token}',
                    'Content-Type': 'application/json'
                }
                response = requests.post(
                    f'https://graph.microsoft.com/v1.0/users/{user_email_or_id}/events',
                    headers=headers,
                    data=json.dumps(event_payload)
                )

                if response.status_code == 201:
                    print("Evento creado con éxito.")
                else:
                    print(f"Error al crear el evento: {response.status_code}")
                    print(response.text)

                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"
                
                # Crear el objeto del mensaje
                msg = MIMEMultipart()
                msg['From'] = direccion_correo
                msg['To'] = destinatario1
                msg['Subject'] = "Instal·lació de fibra"
                msg['Bcc'] = "estiuca@estiuca.cat"

                # Leer la plantilla HTML como una cadena
                with open("C:/Users/camac/Desktop/Recusos/INSTALLFIBRES.html", 'r', encoding='utf-8') as archivo:
                    plantilla_html = archivo.read()

                    # Reemplazar "TAULA1" con el valor de 'dia' en la plantilla
                    cuerpo_mensaje = plantilla_html.replace("TAULA1", dia_valor,)
                    
                    # Reemplazar "TAULA2" con el valor de 'direcció' en la plantilla
                    cuerpo_mensaje = cuerpo_mensaje.replace("TAULA2", direcció_valor,)

                    # Reemplazar "TAULA3" con el valor de 'hora' en la plantilla
                    hora_valor = hora.get() + "h "
                    cuerpo_mensaje = cuerpo_mensaje.replace("TAULA3", hora_valor,)

                    # Reemplazar "TAULA4" con el valor de 'hora2' en la plantilla
                    hora_valor2 = "les " + hora2.get()+ "h "
                    cuerpo_mensaje = cuerpo_mensaje.replace("TAULA4", hora_valor2,)

                    # Cuerpo del mensaje con la plantilla HTML
                    msg.attach(MIMEText(cuerpo_mensaje, 'html'))

                    # Iniciar sesión en el servidor y enviar el correo
                    try:
                        server = smtplib.SMTP('smtp.office365.com', 587) # Ajusta según tu proveedor
                        server.starttls()
                        server.login(direccion_correo, password_correo)
                        server.send_message(msg)
                        server.quit()
                        print("Correu enviat satisfactòriament!!")
                        # Cerrar la ventana emergente
                        ventana_fibres.destroy()
                        # Borrar el valor seleccionado en el Combobox
                        combo_asunto.set("")

                    except Exception as e:
                        print(f"Error al enviar el correu: {e}")
    
            elif seleccion == "Portabilitat Fixe":
                print(f"Enviar email a {destinatario1} con asunto y plantilla 'Portabilitat Fixe'.")
                
                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"

                # Obtener los diferentes valores
                client_valor = client.get()
                dia_valor = dia.get()
                numero_valor = numero.get()
                
                
                # Crear el objeto del mensaje
                msg = MIMEMultipart()
                msg['From'] = direccion_correo
                msg['To'] = destinatario1
                msg['Subject'] = "Portabilitat de línies fixes, ESTIUCA "
                msg['Bcc'] = "estiuca@estiuca.cat"

                # Leer la plantilla HTML como una cadena
                with open("C:/Users/camac/Desktop/Recusos/PORTAFIXE.html", 'r', encoding='utf-8') as archivo:
                    plantilla_html = archivo.read()

                    # Reemplazar "TAULA1" con el valor de 'dia' en la plantilla
                    cuerpo_mensaje = plantilla_html.replace("TAULA1", dia_valor,)
                    
                    # Reemplazar "TAULA2" con el valor de 'les numeracions' en la plantilla
                    cuerpo_mensaje = cuerpo_mensaje.replace("TAULA2", numero_valor,)


                    # Cuerpo del mensaje con la plantilla HTML
                    msg.attach(MIMEText(cuerpo_mensaje, 'html'))

                    # Iniciar sesión en el servidor y enviar el correo
                    try:
                        server = smtplib.SMTP('smtp.office365.com', 587) # Ajusta según tu proveedor
                        server.starttls()
                        server.login(direccion_correo, password_correo)
                        server.send_message(msg)
                        server.quit()
                        print("Correu enviat satisfactòriament!!")
                        # Cerrar la ventana emergente
                        ventana_portaFixe.destroy()
                        # Borrar el valor seleccionado en el Combobox
                        combo_asunto.set("")

                    except Exception as e:
                        print(f"Error al enviar el correu: {e}")

            elif seleccion == "Accés Àrea Client":
                print(f"Enviar email a {destinatario1} con asunto y plantilla 'Accés Àrea Client'.")
                
                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"

                # Obtener los diferentes valores
                usuari_valor = usuari.get()
                password_valor = password.get()
              
                # Crear el objeto del mensaje
                msg = MIMEMultipart()
                msg['From'] = direccion_correo
                msg['To'] = destinatario1
                msg['Subject'] = "Accés Àrea Client, ESTIUCA "
                msg['Bcc'] = "estiuca@estiuca.cat"

                # Leer la plantilla HTML como una cadena
                with open("C:/Users/camac/Desktop/Recusos/AREACLIENT.html", 'r', encoding='utf-8') as archivo:
                    plantilla_html = archivo.read()

                    # Reemplazar "TAULA1" con el valor de 'dia' en la plantilla
                    cuerpo_mensaje = plantilla_html.replace("TAULA1", usuari_valor,)
                    
                    # Reemplazar "TAULA2" con el valor de 'les numeracions' en la plantilla
                    cuerpo_mensaje = cuerpo_mensaje.replace("TAULA2", password_valor,)


                    # Cuerpo del mensaje con la plantilla HTML
                    msg.attach(MIMEText(cuerpo_mensaje, 'html'))

                    # Iniciar sesión en el servidor y enviar el correo
                    try:
                        server = smtplib.SMTP('smtp.office365.com', 587) # Ajusta según tu proveedor
                        server.starttls()
                        server.login(direccion_correo, password_correo)
                        server.send_message(msg)
                        server.quit()
                        print("Correu enviat satisfactòriament!!")
                        # Cerrar la ventana emergente
                        ventana_areaclient.destroy()
                        # Borrar el valor seleccionado en el Combobox
                        combo_asunto.set("")

                    except Exception as e:
                        print(f"Error al enviar el correu: {e}")   

            elif seleccion == "Portabilitat Línies Mòbils":
                print(f"Enviar email a {destinatario1} con asunto y plantilla 'portabilitat linies mòbils'.")
                
                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"

                # Obtener los diferentes valores
                dia_valor = dia.get()
                client_valor=client.get()

                # Leer la plantilla HTML como una cadena
                with open("C:/Users/camac/Desktop/Recusos/PORTASMOBILS.html", 'r', encoding='utf-8') as archivo:
                    plantilla_html = archivo.read()

                    # Reemplazar "TAULA1" con el valor de 'dia' en la plantilla
                    cuerpo_mensaje = plantilla_html.replace("TAULA1", dia_valor,)
                    
                    # Iniciar la tabla HTML con estilo para aumentar el tamaño de la fuente
                    tabla_html = "<table border='1' style='font-family: Calibri;font-size: 22px;color: #3395B3;'><tr><th>SIM</th><th>LÍNIA</th><th>PIN</th><th>PUK</th></tr>"

                    # Agregar cada fila de datos_copiados a la tabla
                    for fila in datos_copiados:
                        tabla_html += "<tr>"
                        for dato in fila:
                            tabla_html += f"<td>{dato}</td>"
                        tabla_html += "</tr>"

                    # Cerrar la tabla HTML
                    tabla_html += "</table>"

                    # Reemplazar en el cuerpo del mensaje
                    cuerpo_mensaje = cuerpo_mensaje.replace("TAULA2", tabla_html)

                # Convertir dia_valor y hora_valor en objetos datetime
                try:
                    fecha = datetime.strptime(dia_valor.strip(), '%d/%m/%Y')
                    hora_completa = datetime.strptime("08:00", '%H:%M')
                    hora_final=datetime.strptime("09:00", '%H:%M')

                    # Combinar fecha y hora
                    fecha_hora_completa = datetime.combine(fecha.date(), hora_completa.time())
                    fecha_hora_final=datetime.combine(fecha.date(), hora_final.time())
                    fecha_hora_iso = fecha_hora_completa.strftime('%Y-%m-%dT%H:%M:%S')
                    fecha_hora2_iso= fecha_hora_final.strftime('%Y-%m-%dT%H:%M:%S')
                    print(fecha_hora_iso, fecha_hora2_iso)

                except ValueError as e:
                    print("Error en la conversión de fecha y hora:", e)
                    return

                # Combinar fecha y hora en formato ISO 8601
                inicio_iso = fecha_hora_iso
                final_iso = fecha_hora2_iso  

                
                # Credenciales y detalles de la aplicación
                client_id = '7fc699bf-de32-4f0c-b75b-1bd6c5864c53'
                client_secret = 'Chq8Q~MIlZyug9JJ3N-g3urWIhSax1IY-BgWYdtI'
                tenant_id = '0450f773-f698-4e8f-ad64-a3477d6c579a'
                authority_url = f'https://login.microsoftonline.com/{tenant_id}'
                scopes = ['https://graph.microsoft.com/.default']
                grant_type = 'client_credentials'

                # Asegúrate de reemplazar 'user_email_or_id' con el correo electrónico o ID del usuario de Outlook
                user_email_or_id = 'gerard@estiuca.cat' 


                # Obtener el token de acceso
                url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
                payload = {
                    'client_id': client_id,
                    'scope': scopes,
                    'client_secret': client_secret,
                    'grant_type': grant_type
                }
                response = requests.post(url, data=payload)
                access_token = response.json()['access_token']

                # Crear el evento
                event_payload = {
                    "subject": "Portabilitats Mòbils " + client_valor,
                    "body": {
                        "contentType": "HTML",
                        "content": "Portabilitats de les següents línies mòbils: <br><br>" + tabla_html
                    },
                    "start": {
                        "dateTime": inicio_iso,
                        "timeZone": "UTC"
                    },
                    "end": {
                        "dateTime": final_iso,
                        "timeZone": "UTC"
                    },
                    "attendees": [
                        {
                            "emailAddress": {
                                "address": "carol@estiuca.cat",
                                "name": "Carol"
                            },
                            "type": "required"
                        }
                    ]
                }

                # Enviar la solicitud para crear el evento
                headers = {
                    'Authorization': f'Bearer {access_token}',
                    'Content-Type': 'application/json'
                }
                response = requests.post(
                    f'https://graph.microsoft.com/v1.0/users/{user_email_or_id}/events',
                    headers=headers,
                    data=json.dumps(event_payload)
                )

                if response.status_code == 201:
                    print("Evento creado con éxito.")
                else:
                    print(f"Error al crear el evento: {response.status_code}")
                    print(response.text)



                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"

                # Crear el objeto del mensaje
                msg = MIMEMultipart()
                msg['From'] = direccion_correo
                msg['To'] = destinatario1
                msg['Subject'] = "Portabilitat de línies mòbils, ESTIUCA "
                msg['Bcc'] = "estiuca@estiuca.cat"

            
                # Cuerpo del mensaje con la plantilla HTML
                msg.attach(MIMEText(cuerpo_mensaje, 'html'))

                # Iniciar sesión en el servidor y enviar el correo
                try:
                        server = smtplib.SMTP('smtp.office365.com', 587) # Ajusta según tu proveedor
                        server.starttls()
                        server.login(direccion_correo, password_correo)
                        server.send_message(msg)
                        server.quit()
                        print("Correu enviat satisfactòriament!!")
                        
                        # Cerrar la ventana emergente
                        ventana_enviarportasmobil.destroy()
                        # Borrar el valor seleccionado en el Combobox
                        combo_asunto.set("")

                except Exception as e:
                        print(f"Error al enviar el correu: {e}")   

            elif seleccion == "Alta Nova":
                print(f"Enviar email a {destinatario1} con asunto y plantilla 'alta nova'.")
                
                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"

                # Obtener los diferentes valores
                dia_valor = dia.get()
                client_valor=client.get()

                # Leer la plantilla HTML como una cadena
                with open("C:/Users/camac/Desktop/Recusos/ALTANOVA.html", 'r', encoding='utf-8') as archivo:
                    plantilla_html = archivo.read()

                    # Reemplazar "TAULA1" con el valor de 'dia' en la plantilla
                    cuerpo_mensaje = plantilla_html.replace("TAULA1", dia_valor,)
                    
                    # Iniciar la tabla HTML con estilo para aumentar el tamaño de la fuente
                    tabla_html = "<table border='1' style='font-family: Calibri;font-size: 22px;color: #3395B3;'><tr><th>SIM</th><th>LÍNIA</th><th>PIN</th><th>PUK</th></tr>"

                    # Agregar cada fila de datos_copiados a la tabla
                    for fila in datos_copiados:
                        tabla_html += "<tr>"
                        for dato in fila:
                            tabla_html += f"<td>{dato}</td>"
                        tabla_html += "</tr>"

                    # Cerrar la tabla HTML
                    tabla_html += "</table>"

                    # Reemplazar en el cuerpo del mensaje
                    cuerpo_mensaje = cuerpo_mensaje.replace("TAULA2", tabla_html)

                # Convertir dia_valor y hora_valor en objetos datetime
                try:
                    fecha = datetime.strptime(dia_valor.strip(), '%d/%m/%Y')
                    hora_completa = datetime.strptime("08:00", '%H:%M')
                    hora_final=datetime.strptime("09:00", '%H:%M')

                    # Combinar fecha y hora
                    fecha_hora_completa = datetime.combine(fecha.date(), hora_completa.time())
                    fecha_hora_final=datetime.combine(fecha.date(), hora_final.time())
                    fecha_hora_iso = fecha_hora_completa.strftime('%Y-%m-%dT%H:%M:%S')
                    fecha_hora2_iso= fecha_hora_final.strftime('%Y-%m-%dT%H:%M:%S')
                    print(fecha_hora_iso, fecha_hora2_iso)

                except ValueError as e:
                    print("Error en la conversión de fecha y hora:", e)
                    return

                # Combinar fecha y hora en formato ISO 8601
                inicio_iso = fecha_hora_iso
                final_iso = fecha_hora2_iso  

                # Credenciales y detalles de la aplicación
                client_id = '7fc699bf-de32-4f0c-b75b-1bd6c5864c53'
                client_secret = 'Chq8Q~MIlZyug9JJ3N-g3urWIhSax1IY-BgWYdtI'
                tenant_id = '0450f773-f698-4e8f-ad64-a3477d6c579a'
                authority_url = f'https://login.microsoftonline.com/{tenant_id}'
                scopes = ['https://graph.microsoft.com/.default']
                grant_type = 'client_credentials'

                # Asegúrate de reemplazar 'user_email_or_id' con el correo electrónico o ID del usuario de Outlook
                user_email_or_id = 'gerard@estiuca.cat' 


                # Obtener el token de acceso
                url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
                payload = {
                    'client_id': client_id,
                    'scope': scopes,
                    'client_secret': client_secret,
                    'grant_type': grant_type
                }
                response = requests.post(url, data=payload)
                access_token = response.json()['access_token']

                # Crear el evento
                event_payload = {
                    "subject": "Alta Mòbil " + client_valor,
                    "body": {
                        "contentType": "HTML",
                        "content": "S'han activat les línies mòbils: <br><br>" + tabla_html
                    },
                    "start": {
                        "dateTime": inicio_iso,
                        "timeZone": "UTC"
                    },
                    "end": {
                        "dateTime": final_iso,
                        "timeZone": "UTC"
                    },
                    "attendees": [
                        {
                            "emailAddress": {
                                "address": "carol@estiuca.cat",
                                "name": "Carol"
                            },
                            "type": "required"
                        }
                    ]
                }

                # Enviar la solicitud para crear el evento
                headers = {
                    'Authorization': f'Bearer {access_token}',
                    'Content-Type': 'application/json'
                }
                response = requests.post(
                    f'https://graph.microsoft.com/v1.0/users/{user_email_or_id}/events',
                    headers=headers,
                    data=json.dumps(event_payload)
                )

                if response.status_code == 201:
                    print("Evento creado con éxito.")
                else:
                    print(f"Error al crear el evento: {response.status_code}")
                    print(response.text)



                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"

                # Crear el objeto del mensaje
                msg = MIMEMultipart()
                msg['From'] = direccion_correo
                msg['To'] = destinatario1
                msg['Subject'] = "Alta de línia mòbil, ESTIUCA "
                msg['Bcc'] = "estiuca@estiuca.cat"

            
                # Cuerpo del mensaje con la plantilla HTML
                msg.attach(MIMEText(cuerpo_mensaje, 'html'))

                # Iniciar sesión en el servidor y enviar el correo
                try:
                        server = smtplib.SMTP('smtp.office365.com', 587) # Ajusta según tu proveedor
                        server.starttls()
                        server.login(direccion_correo, password_correo)
                        server.send_message(msg)
                        server.quit()
                        print("Correu enviat satisfactòriament!!")
                        
                        # Cerrar la ventana emergente
                        ventana_enviarportasmobil.destroy()
                        # Borrar el valor seleccionado en el Combobox
                        combo_asunto.set("")

                except Exception as e:
                        print(f"Error al enviar el correu: {e}")   

            elif seleccion == "EstiuFLY":
                global foto, imagen, country, days
                print(f"Enviar email a {destinatario1}  plantilla 'Estiufly'.")
                
                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiufly@estiufly.com"
                password_correo = "Exit2023*"

                if destinatario1 == "":  # Ajusta esta condición según cómo se manejen los valores no seleccionados en tu combobox
                    msgbox.showwarning("Advertencia", "Si us plau, introdueix un mail del client.")

                # Crear el objeto del mensaje
                msg = MIMEMultipart()
                msg['From'] = direccion_correo
                msg['To'] = destinatario1
                msg['Subject'] = "La teva eSIM d'EstiuFLY "
                msg['Bcc'] = "estiufly@estiufly.com"

                # Obtener los diferentes valores
                dia_i_mes=data1.get()
                country=pais.get()
                days=vigencia.get()
                capacitat=gigas.get()
                parametre=apn.get()
                cobert=cobertura.get()

                # Obtener la imagen del portapapeles
                #def imagen_a_base64(imagen_pil):
                    #buffer = BytesIO()
                    #imagen_pil.save(buffer, format="PNG")
                    #imagen_base64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
                    #return imagen_base64

                def obtener_imagen_portapapeles():
                    global imagen_portapapeles
                    try:
                        imagen_portapapeles = ImageGrab.grabclipboard()
                        return imagen_portapapeles
                    except Exception as e:
                        print(f"Error al obtener la imagen del portapapeles: {e}")
                        return None


                # Leer la plantilla HTML como una cadena
                with open("C:/Users/camac/Desktop/Recusos/ESTIUFLY.html", 'r', encoding='utf-8') as archivo:
                    plantilla_html = archivo.read()
                    
                    # Reemplazar "TAULA1" con el valor de 'dia' en la plantilla
                    cuerpo_mensaje = plantilla_html.replace("COMANDA1", refcomanda.get(),)
                    cuerpo_mensaje = cuerpo_mensaje.replace("DATA1", dia_i_mes,)
                    cuerpo_mensaje = cuerpo_mensaje.replace("PAIS1", country,)
                    cuerpo_mensaje = cuerpo_mensaje.replace("DIES1", days + " dies",)
                    cuerpo_mensaje = cuerpo_mensaje.replace("GBytes", capacitat + " GB",)
                    cuerpo_mensaje = cuerpo_mensaje.replace("APN1", parametre,)     
                    cuerpo_mensaje = cuerpo_mensaje.replace("COBERTURA1", cobert,)                  
                    imagen = obtener_imagen_portapapeles()

                    #if imagen:
                        #imagen = imagen.resize((200, 200), Image.Resampling.LANCZOS)
                        #foto_base64 = imagen_a_base64(imagen)
                        #data_uri = f'data:image/png;base64,{foto_base64}'
                        #imagen_html = f'<img src="{data_uri}" alt="Imagen">'
                        #cuerpo_mensaje = cuerpo_mensaje.replace("QR1", imagen_html)
                        #print(data_uri)

                    penjar_QR_github()

                    # Tamaño deseado para la imagen (en píxeles)
                    ancho_imagen = 250        
                    alto_imagen = 250

                    # Reemplazar "QR1" en el HTML con la etiqueta de la imagen, incluyendo el tamaño
                    cuerpo_mensaje = cuerpo_mensaje.replace("QR1",f'<p style="text-align: center;"> <br><br> <img src="{direccio}" width="{ancho_imagen}" height="{alto_imagen}" alt="Código QR"></p>')
                   


                    # Cuerpo del mensaje con la plantilla HTML
                    msg.attach(MIMEText(cuerpo_mensaje, 'html'))
                    

                    # Iniciar sesión en el servidor y enviar el correo
                    try:
                        server = smtplib.SMTP('smtp.office365.com', 587) # Ajusta según tu proveedor
                        server.starttls()
                        server.login(direccion_correo, password_correo)
                        server.send_message(msg)
                        server.quit()
                        print("Correu enviat satisfactòriament!!")
                        
                        # Borrar el valor seleccionado en el Combobox
                        combo_asunto.set("")

                    except Exception as e:
                        print(f"Error al enviar el correu: {e}")   

                    escribir_archivo_Estiufly()
                    # Cerrar la ventana emergente
                    EstiuFLY.destroy()
            elif seleccion == "Baixa Serveis Residuals":
                print(f"Enviar email a {destinatario1} con asunto y plantilla de 'baixa serveis'.")

                # Configura aquí tu dirección de correo y contraseña
                direccion_correo = "estiuca@estiuca.cat"
                password_correo = "Exit2022*"

                # Crear el objeto del mensaje
                msg = MIMEMultipart()
                msg['From'] = direccion_correo
                msg['To'] = destinatario1
                msg['Subject'] = "Baixes Residuals antic operador"
                #msg['Bcc'] = "estiuca@estiuca.cat"

                # Cuerpo del mensaje con la plantilla HTML
                cuerpo_mensaje = leer_plantilla_html("C:/Users/camac/Desktop/Recusos/BAIXARESIDUAL.html")
                msg.attach(MIMEText(cuerpo_mensaje, 'html'))

                # Iniciar sesión en el servidor y enviar el correo
                try:
                    server = smtplib.SMTP('smtp.office365.com', 587)  # Ajusta según tu proveedor
                    server.starttls()
                    server.login(direccion_correo, password_correo)
                    server.send_message(msg)
                    server.quit()
                    print("Correu enviat satisfactòriament!!")
                    # Cerrar la ventana emergente
                    ventana_baixaresidual.destroy()
                    # Borrar el valor seleccionado en el Combobox
                    combo_asunto.set("")

                except Exception as e:
                    print(f"Error al enviar el correu: {e}")
                
               
# --------------------------------------------------------------------------------------------Función para abrir la ventana emergente welcome
def abrir_ventana_welcome():
    global ventana_welcome,destinatario1
    ventana_welcome=tk.Toplevel(app)
    ventana_welcome.title("Missatge Welcome")

    # Establecer el icono de la ventana emergente
    ventana_welcome.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

    # Obtener las dimensiones de la pantalla
    ancho_pantalla = app.winfo_screenwidth()
    alto_pantalla = app.winfo_screenheight()

    # Tamaño de la ventana emergente (ajustado para ser más grande)
    ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
    alto_ventana_emergente = 425   # Cambiado a un alto mayor

    # Calcular la posición x e y para centrar la ventana
    x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
    y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

    # Configurar la posición y el tamaño de la ventana emergente
    ventana_welcome.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

     # Asegúrate de que la imagen se mantenga en una referencia persistente
    try:
        ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
        imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
        ventana_welcome.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
        label_imagen_ventana_emergente = tk.Label(ventana_welcome, image=ventana_welcome.imagen_ventana_emergente)
        label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,100), pady=10)
    except IOError:
        print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")

    label_mail_1 = tk.Label(ventana_welcome, text="Correu:", font=('Calibri', 16, 'bold'))
    label_mail_1.grid(row=5, column=0, padx=(150,0), pady=10)
    
    destinatario1=tk.Entry(ventana_welcome, font=('Calibri', 14),width=40)
    destinatario1.grid(row=5, column=1, padx=(5,300), pady=10)

   # Agregar espacios iniciales para simular margen
    destinatario1.insert(0, '  ')

    # Situarse en el textbox 'Correu'
    destinatario1.focus_set()

    # ... (resto de los widgets de la ventana emergente)

    boton_enviar = tk.Button(ventana_welcome, text="Enviar", command=lambda: enviar_correo(destinatario1.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
    boton_enviar.grid(row=7, column=0, columnspan=2, padx=(0,150), pady=(10,5))



# --------------------------------------------------------------------------------------------Función para abrir la ventana emergente documentació
def abrir_ventana_documentació():
    global ventana_documentació, destinatario1
    ventana_documentació=tk.Toplevel(app)
    ventana_documentació.title("Sol·licitud de Documentació")

    # Cambiar el icono de la ventana emergente
    ventana_documentació.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

    # Obtener las dimensiones de la pantalla
    ancho_pantalla = app.winfo_screenwidth()
    alto_pantalla = app.winfo_screenheight()

    # Tamaño de la ventana emergente (ajustado para ser más grande)
    ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
    alto_ventana_emergente = 425   # Cambiado a un alto mayor

    # Calcular la posición x e y para centrar la ventana
    x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
    y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

    # Configurar la posición y el tamaño de la ventana emergente
    ventana_documentació.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

    # Asegúrate de que la imagen se mantenga en una referencia persistente
    try:
        ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
        imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
        ventana_documentació.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
        label_imagen_ventana_emergente = tk.Label(ventana_documentació, image=ventana_documentació.imagen_ventana_emergente)
        label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,100), pady=10)
    except IOError:
        print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")

    label_mail_1 = tk.Label(ventana_documentació, text="Correu:", font=('Calibri', 16, 'bold'))
    label_mail_1.grid(row=5, column=0, padx=(150,0), pady=10)
    
    destinatario1 = tk.Entry(ventana_documentació, font=('Calibri', 14),width=40)
    destinatario1.grid(row=5, column=1, padx=(5,300), pady=10)

   # Agregar espacios iniciales para simular margen
    destinatario1.insert(0, '  ')

    # Situarse en el textbox 'Correu'
    destinatario1.focus_set()

    # ... (resto de los widgets de la ventana emergente)

    boton_enviar = tk.Button(ventana_documentació, text="Enviar", command=lambda: enviar_correo(destinatario1.get(),asunto='Benvinguts a ESTIUCA'), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
    boton_enviar.grid(row=7, column=0, columnspan=2, padx=(0,150), pady=(10,5))



# --------------------------------------------------------------------------------------------Función para abrir la ventana emergente Fibres
def abrir_ventana_fibres():
    global ventana_fibres, dia, direcció, hora, hora2, client, destinatario1
    ventana_fibres=tk.Toplevel(app)
    ventana_fibres.title("Instal·lació de Fibra")

    # Cambiar el icono de la ventana emergente
    ventana_fibres.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

    # Obtener las dimensiones de la pantalla
    ancho_pantalla = app.winfo_screenwidth()
    alto_pantalla = app.winfo_screenheight()

    # Tamaño de la ventana emergente (ajustado para ser más grande)
    ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
    alto_ventana_emergente = 600   # Cambiado a un alto mayor

    # Calcular la posición x e y para centrar la ventana
    x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
    y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

    # Configurar la posición y el tamaño de la ventana emergente
    ventana_fibres.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')
    
    # Asegúrate de que la imagen se mantenga en una referencia persistente
    try:
        ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
        imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
        ventana_fibres.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
        label_imagen_ventana_emergente = tk.Label(ventana_fibres, image=ventana_fibres.imagen_ventana_emergente)
        label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,200), pady=10)
    except IOError:
        print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")

    label_mail = tk.Label(ventana_fibres, text="Correu:", font=('Calibri', 16, 'bold'))
    label_mail.grid(row=5, column=0, padx=(150,0), pady=0)
    label_client = tk.Label(ventana_fibres, text="Client:", font=('Calibri', 16, 'bold'))
    label_client.grid(row=6, column=0, padx=(150,0), pady=0)
    label_direcció = tk.Label(ventana_fibres, text="Direcció:", font=('Calibri', 16, 'bold'))
    label_direcció.grid(row=7, column=0, padx=(150,0), pady=(0,0))
    label_dia = tk.Label(ventana_fibres, text="Dia:", font=('Calibri', 16, 'bold'))
    label_dia.grid(row=8, column=0, padx=(150,0), pady=0)
    label_hora = tk.Label(ventana_fibres, text="Hora:", font=('Calibri', 16, 'bold'))
    label_hora.grid(row=9, column=0, padx=(150,0), pady=0)
    label_horaguió = tk.Label(ventana_fibres, text="-", font=('Calibri', 16, 'bold'))
    label_horaguió.grid(row=9, column=1, padx=(0,600), pady=0)
    

    destinatario1 = tk.Entry(ventana_fibres, font=('Calibri', 14),width=40)
    destinatario1.grid(row=5, column=1, padx=(5,325), pady=10)
    client = tk.Entry(ventana_fibres, font=('Calibri', 14),width=40)
    client.grid(row=6, column=1, padx=(5,325), pady=0)
    direcció = tk.Entry(ventana_fibres, font=('Calibri', 14),width=40)
    direcció.grid(row=7, column=1, padx=(5,325), pady=10)
    dia = tk.Entry(ventana_fibres, font=('Calibri', 14),width=15)
    dia.grid(row=8, column=1, padx=(5,575), pady=10)
    hora = tk.Entry(ventana_fibres, font=('Calibri', 14),width=4)
    hora.grid(row=9, column=1, padx=(0,680), pady=10)
    hora2 = tk.Entry(ventana_fibres, font=('Calibri', 14),width=4)
    hora2.grid(row=9, column=1, padx=(5,525), pady=10)

   # Agregar espacios iniciales para simular margen
    destinatario1.insert(0, '  ')
    client.insert(0, '  ')
    direcció.insert(0, '  ')
    dia.insert(0, '  ')
    hora.insert(0, '  ')

    # Situarse en el textbox 'Correu'
    destinatario1.focus_set()

    # ... (resto de los widgets de la ventana emergente)
    
    boton_enviar = tk.Button(ventana_fibres, text="Enviar", command=lambda: enviar_correo(destinatario1.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
    boton_enviar.grid(row=10, column=0, columnspan=2, padx=(250,50), pady=(10,5))

    # Asociar la función de mostrar_tooltip al evento de pasar el mouse sobre la hora inicio y final
    hora.bind("<Enter>", mostrar_tooltip_fibra)
    hora2.bind("<Enter>", mostrar_tooltip_fibra2)



# --------------------------------------------------------------------------------------------Función para abrir la ventana emergente Portabilitat fixe
def abrir_ventana_portaFixe():
    global ventana_portaFixe, dia, numero, client, destinatario1
    ventana_portaFixe = tk.Toplevel(app)
    ventana_portaFixe.title("Portabilitat línia Fixa")

    # Cambiar el icono de la ventana emergente
    ventana_portaFixe.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

    # Obtener las dimensiones de la pantalla
    ancho_pantalla = app.winfo_screenwidth()
    alto_pantalla = app.winfo_screenheight()

    # Tamaño de la ventana emergente (ajustado para ser más grande)
    ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
    alto_ventana_emergente = 500  # Cambiado a un alto mayor

    # Calcular la posición x e y para centrar la ventana
    x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
    y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

    # Configurar la posición y el tamaño de la ventana emergente
    ventana_portaFixe.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')
    
    # Asegúrate de que la imagen se mantenga en una referencia persistente
    try:
        ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
        imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
        ventana_portaFixe.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
        label_imagen_ventana_emergente = tk.Label(ventana_portaFixe, image=ventana_portaFixe.imagen_ventana_emergente)
        label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,200), pady=10)
    except IOError:
        print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")

    label_mail = tk.Label(ventana_portaFixe, text="Correu:", font=('Calibri', 16, 'bold'))
    label_mail.grid(row=5, column=0, padx=(150,0), pady=0)
    label_client = tk.Label(ventana_portaFixe, text="Client:", font=('Calibri', 16, 'bold'))
    label_client.grid(row=6, column=0, padx=(150,0), pady=0)
    label_numero = tk.Label(ventana_portaFixe, text="Número:", font=('Calibri', 16, 'bold'))
    label_numero.grid(row=7, column=0, padx=(150,0), pady=(0,0))
    label_dia = tk.Label(ventana_portaFixe, text="Dia:", font=('Calibri', 16, 'bold'))
    label_dia.grid(row=8, column=0, padx=(150,0), pady=0)
    
    

    destinatario1 = tk.Entry(ventana_portaFixe, font=('Calibri', 14),width=40)
    destinatario1.grid(row=5, column=1, padx=(5,325), pady=10)
    client = tk.Entry(ventana_portaFixe, font=('Calibri', 14),width=40)
    client.grid(row=6, column=1, padx=(5,325), pady=0)
    numero = tk.Entry(ventana_portaFixe, font=('Calibri', 14),width=40)
    numero.grid(row=7, column=1, padx=(5,325), pady=10)
    dia = tk.Entry(ventana_portaFixe, font=('Calibri', 14),width=10)
    dia.grid(row=8, column=1, padx=(5,625), pady=0)
    
   
    
   # Agregar espacios iniciales para simular margen
    destinatario1.insert(0, '  ')
    client.insert(0, '  ')
    numero.insert(0, ' ')
    dia.insert(0, '  ')
    

    # Situarse en el textbox 'Correu'
    destinatario1.focus_set()

    # ... (resto de los widgets de la ventana emergente)
    
    boton_enviar = tk.Button(ventana_portaFixe, text="Enviar", command=lambda: enviar_correo(destinatario1.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
    boton_enviar.grid(row=8, column=0, columnspan=2, padx=(225,5), pady=(0,0))


    # Asociar la función de mostrar_tooltip al evento de pasar el mouse sobre el Entry
    numero.bind("<Enter>", mostrar_tooltip)

# --------------------------------------------------------------------------------------------Función para abrir la ventana emergente Area Client
def abrir_ventana_areaclient():
    global ventana_areaclient, usuari, password, destinatario1
    ventana_areaclient = tk.Toplevel(app)
    ventana_areaclient.title("Accés Àrea Client")

    # Cambiar el icono de la ventana emergente
    ventana_areaclient.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

    # Obtener las dimensiones de la pantalla
    ancho_pantalla = app.winfo_screenwidth()
    alto_pantalla = app.winfo_screenheight()

    # Tamaño de la ventana emergente (ajustado para ser más grande)
    ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
    alto_ventana_emergente = 500  # Cambiado a un alto mayor

    # Calcular la posición x e y para centrar la ventana
    x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
    y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

    # Configurar la posición y el tamaño de la ventana emergente
    ventana_areaclient.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')
    
    # Asegúrate de que la imagen se mantenga en una referencia persistente
    try:
        ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
        imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
        ventana_areaclient.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
        label_imagen_ventana_emergente = tk.Label(ventana_areaclient, image=ventana_areaclient.imagen_ventana_emergente)
        label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,200), pady=10)
    except IOError:
        print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")

    label_mail = tk.Label(ventana_areaclient, text="Correu:", font=('Calibri', 16, 'bold'))
    label_mail.grid(row=5, column=0, padx=(150,0), pady=0)
    label_usuari = tk.Label(ventana_areaclient, text="Usuari:", font=('Calibri', 16, 'bold'))
    label_usuari.grid(row=6, column=0, padx=(150,0), pady=0)
    label_password = tk.Label(ventana_areaclient, text="Password:", font=('Calibri', 16, 'bold'))
    label_password.grid(row=7, column=0, padx=(150,0), pady=(0,0))   

    destinatario1 = tk.Entry(ventana_areaclient, font=('Calibri', 14),width=40)
    destinatario1.grid(row=5, column=1, padx=(5,325), pady=10)
    usuari = tk.Entry(ventana_areaclient, font=('Calibri', 14),width=40)
    usuari.grid(row=6, column=1, padx=(5,325), pady=0)
    password = tk.Entry(ventana_areaclient, font=('Calibri', 14),width=40)
    password.grid(row=7, column=1, padx=(5,325), pady=10)
   
   # Agregar espacios iniciales para simular margen
    destinatario1.insert(0, '  ')
    usuari.insert(0, '  ')
    password.insert(0, ' ')   

    # Situarse en el textbox 'Correu'
    destinatario1.focus_set()

    # ... (resto de los widgets de la ventana emergente)
    
    boton_enviar = tk.Button(ventana_areaclient, text="Enviar", command=lambda: enviar_correo(destinatario1.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
    boton_enviar.grid(row=8, column=0, columnspan=2, padx=(225,5), pady=(0,0))

# --------------------------------------------------------------------------------------------Función para abrir la ventana emergente Portabilitat línies mòbils
def abrir_ventana_portamobil():
    global ventana_portamobil, destinatario1, combo
    
    # Crear una ventana Tkinter
    ventana_portamobil = tk.Toplevel(app)
    ventana_portamobil.withdraw()  # Ocultar la ventana mientras se configura
    ventana_portamobil.title("Portabilitat de linies mòbils")

    # Cambiar el icono de la ventana emergente
    ventana_portamobil.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

    # Obtener las dimensiones de la pantalla
    ancho_pantalla = app.winfo_screenwidth()
    alto_pantalla = app.winfo_screenheight()

    # Tamaño de la ventana emergente (ajustado para ser más grande)
    ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
    alto_ventana_emergente = 500  # Cambiado a un alto mayor

    # Calcular la posición x e y para centrar la ventana
    x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
    y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

    # Configurar la posición y el tamaño de la ventana emergente
    ventana_portamobil.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

    # Asegúrate de que la imagen se mantenga en una referencia persistente
    try:
            ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
            imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
            ventana_portamobil.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
            label_imagen_ventana_emergente = tk.Label(ventana_portamobil, image=ventana_portamobil.imagen_ventana_emergente)
            label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,200), pady=10)
    except IOError:
            print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")


    # Crear un Combobox y su label
    label_combo = tk.Label(ventana_portamobil, text="Nom client:", font=('Calibri', 16, 'bold'), width=15)
    label_combo.grid(row=7, column=0, padx=(50,600), pady=0)

    leer_archivo_SIMS()

    combo = ttk.Combobox(ventana_portamobil, values=unique_values, font=('Calibri', 12),width=40)
    combo.grid(row=7, column=0, padx=(250,350), pady=0)   

    boton_visualizar = tk.Button(ventana_portamobil, text="Visualitzar", command=lambda: visualitzardades(combo.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
    boton_visualizar.grid(row=11, column=0, columnspan=1, padx=(0,175), pady=(25,0))

    # Finalmente, mostrar la ventana después de toda la configuración
    ventana_portamobil.deiconify()

    # Función para manejar la selección
    def visualitzardades(combo):
        if combo == "" or combo is None:  # Ajusta esta condición según cómo se manejen los valores no seleccionados en tu combobox
            msgbox.showwarning("Advertencia", "Si us plau, selecciona una opció.")
            ventana_portamobil.lift()
            ventana_portamobil.focus_force()
            return  # Salimos de la función para no continuar
    
        print("Seleccionado:", combo)

        # Filtrar los datos
        filtro = df[(df['CLIENT'] == combo)]

        # Crear una ventana Tkinter
        ventanilla_portas = tk.Tk()
        ventanilla_portas.title("Línies registrades de: " + combo)
        ventanilla_portas.geometry('600x600')  # Ajustar al tamaño deseado

        # Establecer el icono de la ventana emergente
        ventanilla_portas.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

        # Obtener las dimensiones de la pantalla
        ancho_pantalla = app.winfo_screenwidth()
        alto_pantalla = app.winfo_screenheight()

        # Tamaño de la ventana emergente (ajustado para ser más grande)
        ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
        alto_ventana_emergente = 500  # Cambiado a un alto mayor

        # Configurar la posición y el tamaño de la ventana emergente
        ventanilla_portas.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

        # Crear la vista de árbol (Treeview)
        tree = ttk.Treeview(ventanilla_portas)

        # Configurar el estilo para los encabezados de las columnas
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=('Calibri', 14, 'bold'), foreground='#505050', background='#A9A9A9', anchor=tk.CENTER)
        style.configure("Treeview", font=('Calibri', 12))  # Aplicar estilo a las filas del Treeview


        # Crear la vista de árbol
        tree = ttk.Treeview(ventanilla_portas)

        # Definir las columnas
        tree['columns'] = ('SIM', 'LÍNIA', 'PIN', 'PUK')

        # Formatear las columnas
        tree.column('#0', width=10, stretch=tk.NO)  # Columna fantasma
        tree.column('SIM', anchor=tk.CENTER, width=200)
        tree.column('LÍNIA', anchor=tk.CENTER, width=100)
        tree.column('PIN', anchor=tk.CENTER, width=100)
        tree.column('PUK', anchor=tk.CENTER, width=100)

        # Crear los encabezados
        tree.heading('#0', text='', anchor=tk.CENTER)
        tree.heading('SIM', text='SIM', anchor=tk.CENTER)
        tree.heading('LÍNIA', text='LÍNIA', anchor=tk.CENTER)
        tree.heading('PIN', text='PIN', anchor=tk.CENTER)
        tree.heading('PUK', text='PUK', anchor=tk.CENTER)
        

        # Agregar los datos
        for index, row in filtro.iterrows():
            # Extraemos el valor de 'PIN' y 'PUK'
            pin = row['PIN']
            puk =row['PUK']
            linia= row['LINEA']

            # Verificamos los dígitos de 'PIN' y agregamos tantos '0' como sea necesario al principio 
            if len(str(pin)) == 0:
                pin = "0000" + str(pin)
            if len(str(pin)) == 1:
                pin = "000" + str(pin)
            if len(str(pin)) == 2:
                pin = "00" + str(pin)
            if len(str(pin)) == 3:
                pin = "0" + str(pin)
            # Verificamos los dígitos de 'PUK' y agregamos tantos '0' como sea necesario al principio 
            if len(str(puk)) == 0:
                puk = "00000000" + str(puk)
            if len(str(puk)) == 1:
                puk = "0000000" + str(puk)
            if len(str(puk)) == 2:
                puk = "000000" + str(puk)
            if len(str(puk)) == 3:
                puk = "00000" + str(puk)
            if len(str(puk)) == 4:
                puk = "0000" + str(puk)
            if len(str(puk)) == 5:
                puk = "000" + str(puk)
            if len(str(puk)) == 6:
                puk = "00" + str(puk)
            if len(str(puk)) == 7:
                puk = "0" + str(puk)
            # Verificamos que el campo ' LINIA' no este vacío, si lo esta agregamos la palabra 'alta nova' 
            if linia == "nan":
                linia = "alta nova"
            
            # Insertamos los valores en el árbol
            tree.insert('', tk.END, values=(f"{row['Nº SIM_1']}{row['Nº SIM_2']}", linia, pin, puk))

            # Empaquetar la vista de árbol en la ventana
            tree.pack(expand=True, fill='both')
      
        def copiar_seleccion():
            global datos_copiados, combo
            seleccionadas = tree.selection()
            datos_copiados = [tree.item(i, 'values') for i in seleccionadas]
            ventanilla_portas.destroy()
            abrir_ventana_enviarportasmobil()
            
            # Aquí puedes copiar los datos a donde necesites
            print(datos_copiados)  # Imprime los datos seleccionados para prueba

        # Crear botón para copiar selección
        boton_copiar = tk.Button(ventanilla_portas, text='Seleccionar', command=copiar_seleccion,font=('Calibri', 20, 'bold'), bg='#3395B3', fg='white')
        boton_copiar.pack(side='bottom', pady=10)  # Centrado en la parte inferior

        # Función para manejar la selección/deselección
        def alternar_seleccion():
            # Comprobar si actualmente hay elementos seleccionados
            if tree.selection():
                # Si hay elementos seleccionados, deseleccionar todos
                tree.selection_remove(tree.get_children())
            else:
                # Si no hay elementos seleccionados, seleccionar todos
                for item in tree.get_children():
                    tree.selection_add(item)

        # Crear un botón para alternar la selección
        boton_seleccionar_todo = tk.Button(ventanilla_portas, text='Seleccionar / Deseleccionar-ho tot', command=alternar_seleccion, font=('Calibri', 12))
        boton_seleccionar_todo.pack(side='bottom', pady=10, padx=15)




    # Crear ventana para enviar los datos
    def abrir_ventana_enviarportasmobil():
        global ventana_enviarportasmobil, destinatario1, dia, client
    
        # Crear una ventana Tkinter
        ventana_enviarportasmobil = tk.Toplevel(app)
        ventana_enviarportasmobil.title("Portabilitat de linies mòbils")

        # Cambiar el icono de la ventana emergente
        ventana_enviarportasmobil.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

        # Obtener las dimensiones de la pantalla
        ancho_pantalla = app.winfo_screenwidth()
        alto_pantalla = app.winfo_screenheight()

        # Tamaño de la ventana emergente (ajustado para ser más grande)
        ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
        alto_ventana_emergente = 600  # Cambiado a un alto mayor

        # Calcular la posición x e y para centrar la ventana
        x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
        y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

        # Configurar la posición y el tamaño de la ventana emergente
        ventana_enviarportasmobil.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

        # Asegúrate de que la imagen se mantenga en una referencia persistente
        try:
                ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
                imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
                ventana_enviarportasmobil.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
                label_imagen_ventana_emergente = tk.Label(ventana_enviarportasmobil, image=ventana_enviarportasmobil.imagen_ventana_emergente)
                label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,200), pady=10)
        except IOError:
                print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")

        label_mail = tk.Label(ventana_enviarportasmobil, text="Correu:", font=('Calibri', 16, 'bold'))
        label_mail.grid(row=5, column=0, padx=(150,0), pady=0)
        label_client = tk.Label(ventana_enviarportasmobil, text="Client:", font=('Calibri', 16, 'bold'))
        label_client.grid(row=6, column=0, padx=(150,0), pady=0)
        label_dia = tk.Label(ventana_enviarportasmobil, text="Dia:", font=('Calibri', 16, 'bold'))
        label_dia.grid(row=7, column=0, padx=(150,0), pady=(10,0))

        destinatario1 = tk.Entry(ventana_enviarportasmobil, font=('Calibri', 14),width=40)
        destinatario1.grid(row=5, column=1, padx=(5,325), pady=10)
        client = tk.Entry(ventana_enviarportasmobil, font=('Calibri', 14, 'italic'), fg='gray',width=40,)
        client.grid(row=6, column=1, padx=(0,320), pady=0)
        client.insert(0,combo.get())
        client.config(state='readonly')
        dia = tk.Entry(ventana_enviarportasmobil, font=('Calibri', 14),width=10)
        dia.grid(row=7, column=1, padx=(5,625), pady=(10,0))

        # Situarse en el textbox 'Correu'
        destinatario1.focus_set()
        
        boton_enviar = tk.Button(ventana_enviarportasmobil, text="Enviar", command=lambda: enviar_correo(destinatario1.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
        boton_enviar.grid(row=9, column=0, columnspan=2, padx=(225,5), pady=(0,0))
        ventana_portamobil.destroy()

    # Iniciar el loop de la GUI
    ventana_enviarportasmobil.mainloop()

def abrir_ventana_altanova():
    global ventana_altanova, destinatario1, combo
    
    # Crear una ventana Tkinter
    ventana_altanova = tk.Toplevel(app)
    ventana_altanova.title("Alta Nova")

    # Cambiar el icono de la ventana emergente
    ventana_altanova.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

    # Obtener las dimensiones de la pantalla
    ancho_pantalla = app.winfo_screenwidth()
    alto_pantalla = app.winfo_screenheight()

    # Tamaño de la ventana emergente (ajustado para ser más grande)
    ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
    alto_ventana_emergente = 500  # Cambiado a un alto mayor

    # Calcular la posición x e y para centrar la ventana
    x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
    y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

    # Configurar la posición y el tamaño de la ventana emergente
    ventana_altanova.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

    # Asegúrate de que la imagen se mantenga en una referencia persistente
    try:
            ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
            imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
            ventana_altanova.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
            label_imagen_ventana_emergente = tk.Label(ventana_altanova, image=ventana_altanova.imagen_ventana_emergente)
            label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,200), pady=10)
    except IOError:
            print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")


    # Tus credenciales de Azure AD
    client_id = '7fc699bf-de32-4f0c-b75b-1bd6c5864c53'
    client_secret = 'Chq8Q~MIlZyug9JJ3N-g3urWIhSax1IY-BgWYdtI'
    tenant_id = '0450f773-f698-4e8f-ad64-a3477d6c579a'
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    scopes = ['https://graph.microsoft.com/.default']

    # Crear una instancia de la aplicación confidencial
    app2 = ConfidentialClientApplication(client_id, authority=authority_url, client_credential=client_secret)

    # Obtener el token de acceso
    result = app2.acquire_token_for_client(scopes=scopes)
    token = result['access_token']

    # Asegúrate de reemplazar con el ID de usuario de OneDrive o el ID del sitio de SharePoint
    user_id = 'gerard@estiuca.cat'

    # Formar la URL para acceder al archivo en OneDrive
    file_url = f'https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/Seguiment Propostes/RELACIÓ SIMs.xlsm:/content'

    # Headers para la solicitud
    headers = {'Authorization': f'Bearer {token}'}

    # Realizar la solicitud GET para obtener el archivo
    response = requests.get(file_url, headers=headers)

   
    # Procesar los datos
    data = BytesIO(response.content)
    df = pd.read_excel(data, engine='openpyxl', sheet_name='GENERAL')

    # Obtener valores únicos de la columna E
    unique_values = df['CLIENT'].dropna().unique().tolist()

    # Crear un Combobox y su label
    label_combo = tk.Label(ventana_altanova, text="Nom client:", font=('Calibri', 16, 'bold'), width=15)
    label_combo.grid(row=7, column=0, padx=(50,600), pady=0)
    
    combo = ttk.Combobox(ventana_altanova, values=unique_values, font=('Calibri', 12),width=40)
    combo.grid(row=7, column=0, padx=(250,350), pady=0)

    boton_visualizar = tk.Button(ventana_altanova, text="Visualitzar", command=lambda: visualitzarAltaMobil(combo.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
    boton_visualizar.grid(row=11, column=0, columnspan=1, padx=(0,175), pady=(25,0))

    # Función para manejar la selección
    def visualitzarAltaMobil(combo):
        if combo == "" or combo is None:  # Ajusta esta condición según cómo se manejen los valores no seleccionados en tu combobox
            msgbox.showwarning("Advertencia", "Si us plau, selecciona una opció.")
            ventana_portamobil.lift()
            ventana_portamobil.focus_force()
            return  # Salimos de la función para no continuar
    
        print("Seleccionado:", combo)

        # Filtrar los datos
        filtro = df[(df['CLIENT'] == combo) & (df['ORIGEN'] == "ALTA")]

        # Crear una ventana Tkinter
        ventanilla_altes = tk.Tk()
        ventanilla_altes.title("Línies registrades de: " + combo)
        ventanilla_altes.geometry('600x600')  # Ajustar al tamaño deseado

        # Establecer el icono de la ventana emergente
        ventanilla_altes.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

        # Obtener las dimensiones de la pantalla
        ancho_pantalla = app.winfo_screenwidth()
        alto_pantalla = app.winfo_screenheight()

        # Tamaño de la ventana emergente (ajustado para ser más grande)
        ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
        alto_ventana_emergente = 500  # Cambiado a un alto mayor

        # Configurar la posición y el tamaño de la ventana emergente
        ventanilla_altes.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

        # Crear la vista de árbol (Treeview)
        tree = ttk.Treeview(ventanilla_altes)

        # Configurar el estilo para los encabezados de las columnas
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", font=('Calibri', 14, 'bold'), foreground='#505050', background='#A9A9A9', anchor=tk.CENTER)
        style.configure("Treeview", font=('Calibri', 12))  # Aplicar estilo a las filas del Treeview


        # Crear la vista de árbol
        tree = ttk.Treeview(ventanilla_altes)

        # Definir las columnas
        tree['columns'] = ('SIM', 'LÍNIA', 'PIN', 'PUK')

        # Formatear las columnas
        tree.column('#0', width=10, stretch=tk.NO)  # Columna fantasma
        tree.column('SIM', anchor=tk.CENTER, width=200)
        tree.column('LÍNIA', anchor=tk.CENTER, width=100)
        tree.column('PIN', anchor=tk.CENTER, width=100)
        tree.column('PUK', anchor=tk.CENTER, width=100)

        # Crear los encabezados
        tree.heading('#0', text='', anchor=tk.CENTER)
        tree.heading('SIM', text='SIM', anchor=tk.CENTER)
        tree.heading('LÍNIA', text='LÍNIA', anchor=tk.CENTER)
        tree.heading('PIN', text='PIN', anchor=tk.CENTER)
        tree.heading('PUK', text='PUK', anchor=tk.CENTER)
        

        # Agregar los datos
        for index, row in filtro.iterrows():
            # Extraemos el valor de 'PIN' y 'PUK'
            pin = row['PIN']
            puk =row['PUK']
            linia= row['LINEA']

            # Verificamos los dígitos de 'PIN' y agregamos tantos '0' como sea necesario al principio 
            if len(str(pin)) == 0:
                pin = "0000" + str(pin)
            if len(str(pin)) == 1:
                pin = "000" + str(pin)
            if len(str(pin)) == 2:
                pin = "00" + str(pin)
            if len(str(pin)) == 3:
                pin = "0" + str(pin)
            # Verificamos los dígitos de 'PUK' y agregamos tantos '0' como sea necesario al principio 
            if len(str(puk)) == 0:
                puk = "00000000" + str(puk)
            if len(str(puk)) == 1:
                puk = "0000000" + str(puk)
            if len(str(puk)) == 2:
                puk = "000000" + str(puk)
            if len(str(puk)) == 3:
                puk = "00000" + str(puk)
            if len(str(puk)) == 4:
                puk = "0000" + str(puk)
            if len(str(puk)) == 5:
                puk = "000" + str(puk)
            if len(str(puk)) == 6:
                puk = "00" + str(puk)
            if len(str(puk)) == 7:
                puk = "0" + str(puk)
            # Verificamos que el campo ' LINIA' no este vacío, si lo esta agregamos la palabra 'alta nova' 
            if linia == "nan":
                linia = "alta nova"
            
            # Insertamos los valores en el árbol
            tree.insert('', tk.END, values=(f"{row['Nº SIM_1']}{row['Nº SIM_2']}", linia, pin, puk))

            # Empaquetar la vista de árbol en la ventana
            tree.pack(expand=True, fill='both')
      
        def copiar_seleccion():
            global datos_copiados, combo
            seleccionadas = tree.selection()
            datos_copiados = [tree.item(i, 'values') for i in seleccionadas]
            ventana_altanova.destroy()
            abrir_ventana_enviaraltamobil

        # Crear botón para copiar selección
        boton_copiar = tk.Button(ventana_altanova, text='Seleccionar', command=copiar_seleccion,font=('Calibri', 20, 'bold'), bg='#3395B3', fg='white')
        boton_copiar.pack(side='bottom', pady=10)  # Centrado en la parte inferior

        # Función para manejar la selección/deselección
        def alternar_seleccion():
            # Comprobar si actualmente hay elementos seleccionados
            if tree.selection():
                # Si hay elementos seleccionados, deseleccionar todos
                tree.selection_remove(tree.get_children())
            else:
                # Si no hay elementos seleccionados, seleccionar todos
                for item in tree.get_children():
                    tree.selection_add(item)

        # Crear un botón para alternar la selección
        boton_seleccionar_todo = tk.Button(ventanilla_altes, text='Seleccionar / Deseleccionar-ho tot', command=alternar_seleccion, font=('Calibri', 12))
        boton_seleccionar_todo.pack(side='bottom', pady=10, padx=15)

    # Crear ventana para enviar los datos
    def abrir_ventana_enviaraltamobil():
        global ventana_enviaraltamobil, destinatario1, dia, client
    
        # Crear una ventana Tkinter
        ventana_enviaraltamobil = tk.Toplevel(app)
        ventana_enviaraltamobil.title("Portabilitat de linies mòbils")

        # Cambiar el icono de la ventana emergente
        ventana_enviaraltamobil.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

        # Obtener las dimensiones de la pantalla
        ancho_pantalla = app.winfo_screenwidth()
        alto_pantalla = app.winfo_screenheight()

        # Tamaño de la ventana emergente (ajustado para ser más grande)
        ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
        alto_ventana_emergente = 600  # Cambiado a un alto mayor

        # Calcular la posición x e y para centrar la ventana
        x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
        y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

        # Configurar la posición y el tamaño de la ventana emergente
        ventana_enviaraltamobil.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

        # Asegúrate de que la imagen se mantenga en una referencia persistente
        try:
                ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
                imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
                ventana_enviaraltamobil.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
                label_imagen_ventana_emergente = tk.Label(ventana_enviaraltamobil, image=ventana_enviaraltamobil.imagen_ventana_emergente)
                label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,200), pady=10)
        except IOError:
                print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")

        label_mail = tk.Label(ventana_enviaraltamobil, text="Correu:", font=('Calibri', 16, 'bold'))
        label_mail.grid(row=5, column=0, padx=(150,0), pady=0)
        label_client = tk.Label(ventana_enviaraltamobil, text="Client:", font=('Calibri', 16, 'bold'))
        label_client.grid(row=6, column=0, padx=(150,0), pady=0)
        label_dia = tk.Label(ventana_enviaraltamobil, text="Dia:", font=('Calibri', 16, 'bold'))
        label_dia.grid(row=7, column=0, padx=(150,0), pady=(10,0))

        destinatario1 = tk.Entry(ventana_enviaraltamobil, font=('Calibri', 14),width=40)
        destinatario1.grid(row=5, column=1, padx=(5,325), pady=10)
        client = tk.Entry(ventana_enviaraltamobil, font=('Calibri', 14, 'italic'), fg='gray',width=40,)
        client.grid(row=6, column=1, padx=(0,320), pady=0)
        client.insert(0,combo.get())
        client.config(state='readonly')
        dia = tk.Entry(ventana_enviaraltamobil, font=('Calibri', 14),width=10)
        dia.grid(row=7, column=1, padx=(5,625), pady=(10,0))

        # Situarse en el textbox 'Correu'
        destinatario1.focus_set()
        
        boton_enviar = tk.Button(ventana_enviaraltamobil, text="Enviar", command=lambda: enviar_correo(destinatario1.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
        boton_enviar.grid(row=9, column=0, columnspan=2, padx=(225,5), pady=(0,0))
        ventana_portamobil.destroy()

    # Iniciar el loop de la GUI
    ventana_enviaraltamobil.mainloop()

# --------------------------------------------------------------------------------------------Función para abrir la ventana emergente ESTIUFLY
def abrir_EstiuFLY():
    global EstiuFLY,destinatario1, foto, apn, gigas, vigencia, recarga, refcomanda, pais, empresa, dia, cost, empresa, cobertura, fecha, data1
    EstiuFLY=tk.Toplevel(app)
    EstiuFLY.withdraw()  # Ocultar la ventana mientras se configura
    EstiuFLY.title("Benvingut a EstiuFLY!!")

    # Establecer el icono de la ventana emergente
    EstiuFLY.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

    # Obtener las dimensiones de la pantalla
    ancho_pantalla = app.winfo_screenwidth()
    alto_pantalla = app.winfo_screenheight()

    # Tamaño de la ventana emergente (ajustado para ser más grande)
    ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
    alto_ventana_emergente = 600   # Cambiado a un alto mayor

    # Ajustar el tamaño de la ventana al tamaño de la pantalla (por ejemplo, 80% del tamaño de la pantalla)
    ancho_ventana_emergente = int(ancho_pantalla * 0.8)
    alto_ventana_emergente = int(alto_pantalla * 0.8)

    # Calcular la posición x e y para centrar la ventana
    x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
    y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

    # Configurar la posición y el tamaño de la ventana emergente
    EstiuFLY.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

     # Asegúrate de que la imagen se mantenga en una referencia persistente
    try:
        ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiuFlyLogo.jpg'
        imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)

        # Redimensionar la imagen para ajustarse al ancho de la ventana
        ancho_imagen = ancho_ventana_emergente
        factor_redimension = ancho_imagen / imagen_original_ventana_emergente.width
        alto_imagen = int(imagen_original_ventana_emergente.height * factor_redimension)
        imagen_redimensionada = imagen_original_ventana_emergente.resize((ancho_imagen, alto_imagen), Image.Resampling.LANCZOS)

        EstiuFLY.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_redimensionada)
        label_imagen_ventana_emergente = tk.Label(EstiuFLY, image=EstiuFLY.imagen_ventana_emergente)
        label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=3)  # Centrar la imagen
    except IOError:
        print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")

    obtener_imagen_portapapeles2()
       
    actualizar_imagen_frame()
        
    leer_archivo_Estiufly()

    def actualizar_preu(event=None):
        try:
            # Obtener el valor de 'cost'
            valor_cost = float(cost.get())

            # Calcular el nuevo valor de 'preu'
            nuevo_valor_preu = valor_cost * 3

            # Limpiar el campo 'preu' y insertar el nuevo valor
            preu.delete(0, tk.END)
            preu.insert(0, nuevo_valor_preu)
        except ValueError:
            # Si 'cost' no es un número, limpia 'preu'
            preu.delete(0, tk.END)

    dia=datetime.now()
    dia=dia.strftime("%d/%m/%Y")
    

     # Crear tres frames para las columnas
    frame_dades = tk.Frame(EstiuFLY)
    frame_dades2 = tk.Frame(EstiuFLY)
    frame_dades3 = tk.Frame(EstiuFLY)

    # Organizar los frames en la ventana principal
    frame_dades.grid(row=1, column=0, sticky='nsew', padx=(50, 0), pady=15)
    frame_dades2.grid(row=1, column=1, sticky='nsew', padx=(0, 0), pady=15)
    frame_dades3.grid(row=1, column=2, sticky='nsew', padx=(5, 0), pady=15)

    # Configurar la distribución de columnas en la ventana principal
    EstiuFLY.grid_columnconfigure(0, weight=1)
    EstiuFLY.grid_columnconfigure(1, weight=1)
    EstiuFLY.grid_columnconfigure(2, weight=1)

    # Agregar widgets al Frame Dades

    label_refcomanda = tk.Label(frame_dades, text="Comanda", font=('Calibri', 14, 'bold'))
    label_refcomanda.grid(row=0, column=0, sticky='w', padx=5, pady=10)
    label_data = tk.Label(frame_dades, text="Data", font=('Calibri', 14, 'bold'))
    label_data.grid(row=1, column=0, sticky='w', padx=5, pady=10)
    label_empresa = tk.Label(frame_dades, text="Client", font=('Calibri', 14, 'bold'))
    label_empresa.grid(row=2, column=0, sticky='w', padx=5, pady=10)
    label_destinatario1 = tk.Label(frame_dades, text="Mail", font=('Calibri', 14, 'bold'))
    label_destinatario1.grid(row=3, column=0, sticky='w', padx=5, pady=10)
    label_cost= tk.Label(frame_dades, text="Cost", font=('Calibri', 14, 'bold'))
    label_cost.grid(row=4, column=0, sticky='w', padx=5, pady=10)
    label_preu= tk.Label(frame_dades, text="PVP", font=('Calibri', 14,), fg="gray")
    label_preu.grid(row=5, column=0, sticky='w', padx=5, pady=10)

    # Agregar campos de entrada al Frame Dades

    refcomanda = tk.Entry(frame_dades, font=('Calibri', 14))
    refcomanda.grid(row=0, column=1, padx=(0,135), pady=10)
    refcomanda.insert(0, resultado)
    data1 = tk.Entry(frame_dades, font=('Calibri', 14))
    data1.grid(row=1, column=1, padx=(0,135), pady=10)
    data1.insert(0, dia)
    empresa = tk.Entry(frame_dades, font=('Calibri', 14),width=30)
    empresa.grid(row=2, column=1, padx=(0,10), pady=10)
    destinatario1 = tk.Entry(frame_dades, font=('Calibri', 14), width=30)
    destinatario1.grid(row=3, column=1, padx=(0,10), pady=10)
    cost = tk.Entry(frame_dades, font=('Calibri', 14))
    cost.grid(row=4, column=1, padx=(0,135), pady=10)
    cost.bind('<KeyRelease>', actualizar_preu)  # Enlazar el evento de liberación de tecla

    preu = tk.Entry(frame_dades, font=('Calibri', 14), fg="gray")
    preu.grid(row=5, column=1, padx=(0,135), pady=10)
    # Agregar widgets al Frame Dades2

    label_pais = tk.Label(frame_dades2, text="Païs", font=('Calibri', 14, 'bold'), width=5)
    label_pais.grid(row=0, column=0, sticky='w', padx=5, pady=10)
    label_gigas = tk.Label(frame_dades2, text="GB", font=('Calibri', 14, 'bold'), width=5)
    label_gigas.grid(row=1, column=0, sticky='w', padx=5, pady=10)
    label_vigencia = tk.Label(frame_dades2, text="Vigència", font=('Calibri', 14, 'bold'), width=8)
    label_vigencia.grid(row=2, column=0, sticky='w', padx=5, pady=10)
    label_recarga = tk.Label(frame_dades2, text="Recarga", font=('Calibri', 14, 'bold'), width=8)
    label_recarga.grid(row=3, column=0, sticky='w', padx=5, pady=10)
    label_apn = tk.Label(frame_dades2, text="APN", font=('Calibri', 14, 'bold'))
    label_apn.grid(row=4, column=0, sticky='w', padx=5, pady=10)
    label_cobertura = tk.Label(frame_dades2, text="Cobertura", font=('Calibri', 14, 'bold'))
    label_cobertura.grid(row=5, column=0, sticky='w', padx=5, pady=10)
    
    

    pais=tk.Entry(frame_dades2, font=('Calibri', 14), width=25)
    pais.grid(row=0, column=1, padx=(0,60), pady=10)
    gigas=tk.Entry(frame_dades2, font=('Calibri', 14))
    gigas.grid(row=1, column=1, padx=(0,135), pady=10)
    vigencia=tk.Entry(frame_dades2, font=('Calibri', 14))
    vigencia.grid(row=2, column=1, padx=(0,135), pady=10)
    recarga=tk.Entry(frame_dades2, font=('Calibri', 14))
    recarga.grid(row=3, column=1, padx=(0,135), pady=10)
    apn=tk.Entry(frame_dades2, font=('Calibri', 14))
    apn.grid(row=4, column=1, padx=(0,135), pady=10)
    cobertura=tk.Entry(frame_dades2, font=('Calibri', 14))
    cobertura.grid(row=5, column=1, padx=(0,135), pady=10)
    
    fecha=data1.get()


    # Agregar un aviso para seleccionar y copiar el QR
    msgbox.showwarning("Advertencia", "Si us plau, copia el codi QR de la web, i clica acceptar.")

    # Crear un Label en frame_dades3 para mostrar la imagen
    actualizar_imagen_frame()
    label_imagen_portapapeles = tk.Label(frame_dades3, image=foto)
    label_imagen_portapapeles.image = foto  # Guardar una referencia
    label_imagen_portapapeles.grid(row=2, column=0, padx=5, pady=10)  # Ajustar según sea necesario    
    
    boton_enviar = tk.Button(frame_dades3, text="Enviar", command=lambda: enviar_correo(destinatario1.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
    boton_enviar.grid(row=4, column=0, padx=(0,0), pady=(10,0))

    # Finalmente, mostrar la ventana después de toda la configuración
    EstiuFLY.deiconify()
    EstiuFLY.mainloop()

# --------------------------------------------------------------------------------------------Función para abrir la ventana emergente Baixa Residual
def abrir_ventana_baixaresidual():
    global ventana_baixaresidual,destinatario1
    ventana_baixaresidual=tk.Toplevel(app)
    ventana_baixaresidual.title("Donar de Baixa serveis residuals")

    # Establecer el icono de la ventana emergente
    ventana_baixaresidual.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')

    # Obtener las dimensiones de la pantalla
    ancho_pantalla = app.winfo_screenwidth()
    alto_pantalla = app.winfo_screenheight()

    # Tamaño de la ventana emergente (ajustado para ser más grande)
    ancho_ventana_emergente = 800  # Cambiado a un ancho mayor
    alto_ventana_emergente = 425   # Cambiado a un alto mayor

    # Calcular la posición x e y para centrar la ventana
    x = (ancho_pantalla // 2) - (ancho_ventana_emergente // 2)
    y = (alto_pantalla // 2) - (alto_ventana_emergente // 2)

    # Configurar la posición y el tamaño de la ventana emergente
    ventana_baixaresidual.geometry(f'{ancho_ventana_emergente}x{alto_ventana_emergente}+{x}+{y}')

     # Asegúrate de que la imagen se mantenga en una referencia persistente
    try:
        ruta_imagen_ventana_emergente = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Cambia a la ruta correcta de tu imagen
        imagen_original_ventana_emergente = Image.open(ruta_imagen_ventana_emergente)
        ventana_baixaresidual.imagen_ventana_emergente = ImageTk.PhotoImage(imagen_original_ventana_emergente,)
        label_imagen_ventana_emergente = tk.Label(ventana_baixaresidual, image=ventana_baixaresidual.imagen_ventana_emergente)
        label_imagen_ventana_emergente.grid(row=0, column=0, columnspan=2, padx=(0,100), pady=10)
    except IOError:
        print(f"No se pudo cargar la imagen desde {ruta_imagen_ventana_emergente}")

    label_mail_1 = tk.Label(ventana_baixaresidual, text="Correu:", font=('Calibri', 16, 'bold'))
    label_mail_1.grid(row=5, column=0, padx=(150,0), pady=10)
    
    destinatario1=tk.Entry(ventana_baixaresidual, font=('Calibri', 14),width=40)
    destinatario1.grid(row=5, column=1, padx=(5,300), pady=10)

   # Agregar espacios iniciales para simular margen
    destinatario1.insert(0, '  ')

    # Situarse en el textbox 'Correu'
    destinatario1.focus_set()

    # ... (resto de los widgets de la ventana emergente)

    boton_enviar = tk.Button(ventana_baixaresidual, text="Enviar", command=lambda: enviar_correo(destinatario1.get()), font=('Calibri', 18, 'bold'), bg='#3395B3', fg='white',width=10, height=1)
    boton_enviar.grid(row=7, column=0, columnspan=2, padx=(0,150), pady=(10,5))



# Configuración de la ventana principal
def on_combobox_select(event):
    # Coordenadas del Combobox en la pantalla
    x0 = app.winfo_rootx() + combo_asunto.winfo_x()
    y0 = app.winfo_rooty() + combo_asunto.winfo_y()

    # Dimensiones del Combobox
    width = combo_asunto.winfo_width()
    height = combo_asunto.winfo_height()

    # Coordenadas para hacer clic en el centro del campo de texto del Combobox
    x = x0 + width // 2
    y = y0 + height // 2

    # Hacer clic en el campo de texto del Combobox
    pyautogui.click(x, y)

    # Esperar un breve momento antes de cambiar el foco
    app.after(100, lambda: button_enviar.focus_set())


# Configurar el boton enviar
def on_button_enviar():
    seleccion = combo_asunto.get()
  
    if seleccion == "welcome":
        abrir_ventana_welcome()
    elif seleccion == "Enviar documentació":
        abrir_ventana_documentació()
    elif seleccion == "Instal·lació de Fibra":
        abrir_ventana_fibres()
    elif seleccion == "Portabilitat Fixe":
        abrir_ventana_portaFixe()
    elif seleccion == "Accés Àrea Client":
        abrir_ventana_areaclient()
    elif seleccion == "Portabilitat Línies Mòbils":
        abrir_ventana_portamobil()
    elif seleccion == "Alta Nova":
        abrir_ventana_altanova()
    elif seleccion=="EstiuFLY":
        abrir_EstiuFLY()
    elif seleccion=="Baixa Serveis Residuals":
        abrir_ventana_baixaresidual()

app = tk.Tk()
app.title("Enviar Missatges a Clients")

# Establecer el icono de la ventana
app.iconbitmap('C:/Users/camac/Desktop/Recusos/Favicon.ico')  # Reemplaza con la ruta a tu archivo .ico
# Obtener las dimensiones de la pantalla
ancho_pantalla = app.winfo_screenwidth()
alto_pantalla = app.winfo_screenheight()

# Tamaño de la ventana
ancho_ventana = 800  # Ajusta esto a las dimensiones deseadas
alto_ventana = 450   # Ajusta esto a las dimensiones deseadas

# Calcular la posición x e y para centrar la ventana
x = (ancho_pantalla // 2) - (ancho_ventana // 2)
y = (alto_pantalla // 2) - (alto_ventana // 2)

# Establecer el tamaño y la posición de la ventana
app.geometry(f'{ancho_ventana}x{alto_ventana}+{x}+{y}')

style = ttk.Style(app)
style.theme_use('clam')
style.configure("TCombobox", font=('Calibri', 16))

# Cargar y mostrar la imagen
ruta_imagen = 'C:/Users/camac/Desktop/Recusos/EstiucaCap.png'  # Asegúrate de que esta ruta sea correcta
try:
    imagen_original = Image.open(ruta_imagen)
    imagen = ImageTk.PhotoImage(imagen_original)
    label_imagen = tk.Label(app, image=imagen)
    label_imagen.pack(pady=(0, 20))
except IOError:
    print(f"No se pudo cargar la imagen desde {ruta_imagen}")


# Menú desplegable para el asunto
opciones_asunto = ["welcome", "Enviar documentació", "Instal·lació de Fibra", "Portabilitat Fixe", "Accés Àrea Client" , "Portabilitat Línies Mòbils" , "Alta Nova", "Baixa Serveis Residuals", "EstiuFLY"]
combo_asunto = ttk.Combobox(app, values=opciones_asunto, style="TCombobox", font=('Calibri', 16, 'bold'))
combo_asunto.pack(pady=(0, 20))

# Enlace del evento de selección en el Combobox
combo_asunto.bind("<<ComboboxSelected>>", on_combobox_select)

# Crear un Frame como contenedor con un tamaño específico
frame_contenedor = tk.Frame(app, width=150, height=100)
frame_contenedor.pack_propagate(False)  # Evita que el contenedor se ajuste al tamaño del botón
frame_contenedor.pack(pady=10)

# Hace que el botón llene el Frame
button_enviar = tk.Button(frame_contenedor, text="Enviar", command=on_button_enviar,font=('Arial', 20, 'bold'), bg='#3395B3', fg='white')
button_enviar.place(relwidth=1, height=60)  # Aquí puedes ajustar 'height' como desees



app.mainloop()
