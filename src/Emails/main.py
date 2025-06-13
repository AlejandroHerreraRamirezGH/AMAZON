import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import win32com.client as win32
from src.Fuji.get_data import GetData
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication



class Envios_de_correo:

    def Extraer_informacion_get_data_credenciales(self):
        getdata = GetData()
        datos = getdata.get_datos_id(1)

        if datos:
            server_smtp = datos.get("server_smtp")
            port_smtp = datos.get("port_smtp")
            user_smtp = datos.get("user_smtp")
            pass_smtp = datos.get("pass_smtp")

            return server_smtp, port_smtp, user_smtp, pass_smtp
        else:
            print("No se encontraron datos para el ID proporcionado.")
            return None, None, None, None

    def Enviar_correo(self,server_smtp, port_smtp, user_smtp, pass_smtp, destinatario, asunto, cuerpo, path_root):
        try:
            ruta_adjunto = os.path.join(path_root, 'Productos_por_paginas_divididos.xlsx')
            mensaje = MIMEMultipart()
            mensaje['From'] = user_smtp
            mensaje['To'] = destinatario
            mensaje['Subject'] = asunto
            mensaje.attach(MIMEText(cuerpo, 'plain'))

            with open(ruta_adjunto, 'rb') as adjunto:
                parte = MIMEApplication(adjunto.read(), Name=os.path.basename(ruta_adjunto))
                parte['Content-Disposition'] = f'attachment; filename="{os.path.basename(ruta_adjunto)}"'
                mensaje.attach(parte)

            servidor = smtplib.SMTP(server_smtp, port_smtp)
            servidor.starttls() 
            servidor.login(user_smtp, pass_smtp)  
            servidor.send_message(mensaje)
            servidor.quit()

            print("Correo enviado exitosamente.")

        except Exception as e:
            print("Error al enviar el correo:", e)