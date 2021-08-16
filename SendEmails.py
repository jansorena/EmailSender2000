# librerias
import email, smtplib, ssl, csv, getpass, os
import pandas as pd

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# contraseña
password = getpass.getpass('Ingrese la contraseña: ')
print(' ')
filename2=Path('boleta.pdf').stem

# arreglos con los datos
nombres=[]
apellidopaterno=[]
apellidomaterno=[]
correos=[]
estadodepago=[]

nombres_pagados=[]
apellidopaterno_pagados=[]
apellidomaterno_pagados=[]
correos_pagados=[]
estadodepago_pagados=[]

nombres_nopagados=[]
apellidopaterno_nopagados=[]
apellidomaterno_nopagados=[]
correos_nopagados=[]
estadodepago_nopagados=[]

#leer e indexar los datos
df = pd.read_csv("~/Documents/SendEmails/BoletasPrueba.csv")

nombres = df['Nombre'].tolist()
apellidopaterno=df['Apellido Paterno'].tolist()
apellidomaterno=df['Apellido Materno'].tolist()
correos=df['Correo'].tolist()
estadodepago=df['Mensualidad_1'].tolist()

for (a,b,c,d,e) in zip(apellidopaterno,apellidomaterno,nombres,correos,estadodepago):
    filename = (filename2+' '+c+' '+a+' '+b+'.pdf')
    try:
        open(filename, "rb")
        nombres_pagados.append(c)
        apellidopaterno_pagados.append(a)
        apellidomaterno_pagados.append(b)
        correos_pagados.append(d)
        estadodepago_pagados.append(e)
    except FileNotFoundError:
        nombres_nopagados.append(c)
        apellidopaterno_nopagados.append(a)
        apellidomaterno_nopagados.append(b)
        correos_nopagados.append(d)
        estadodepago_nopagados.append(e)

print('-- NO PAGADOS -- La lista de correos es la siguiente :', end='\n')
print(' ')
for (a,b,c,d,e) in zip(apellidopaterno_nopagados,apellidomaterno_nopagados,nombres_nopagados,correos_nopagados,estadodepago_nopagados):
    print(c+" "+a+" "+b+"----"+d, end='\n')
print(' ')
print('-- PAGADOS -- La lista de correos es la siguiente :', end='\n')
print(' ')
for (a,b,c,d,e) in zip(apellidopaterno_pagados,apellidomaterno_pagados,nombres_pagados,correos_pagados,estadodepago_pagados):
    filename = (filename2+' '+c+' '+a+' '+b+'.pdf')
    print(c+" "+a+" "+b+"----"+d+"----"+"Archivo: "+filename, end='\n')
    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
    "Content-Disposition",
    f"attachment; filename= {filename}",
)
print(' ')

entrada = input('Desea enviar correo a:\n [1] PAGADOS\n [2] NO PAGADOS\n [3] SALIR\n')
if entrada == '1':
    print(' ')
    while True:
        subject = input('Ingrese el titulo: ')
        print(' ')
        body = input('Ingrese el mensaje: ')
        print(' ')
        print('Vista previa del mensaje: ')
        print(' ')
        print(subject,body,sep='\n')
        print(' ')
        entrada2 = input("Desea reescribir el mensaje?: (si/no) ")
        print(' ')
        if entrada2 == "no":
            break

    entrada = input("Desea enviar los correos? (si/no): ")
    print(' ')
    if entrada == 'si':
        for (a,b,c,d,e) in zip(apellidopaterno_pagados,apellidomaterno_pagados,nombres_pagados,correos_pagados,estadodepago_pagados):
            filename = (filename2+' '+c+' '+a+' '+b+'.pdf')
            subject
            body
            sender_email = "mensualidades@colegiomanuelrodriguez.cl"
            receiver_email = d

    # Create a multipart message and set headers
            message = MIMEMultipart()
            message["From"] = 'Mensualidades Colegio'
            message["To"] = d
            message["Subject"] = subject

    # Add body to email
            message.attach(MIMEText(body, "plain"))

            with open(filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename}",
    )

            message.attach(part)
            text = message.as_string()

    # Log in to server using secure context and send email
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(sender_email, d, text)
                print("Mensaje se ha enviado a: "+c+" "+a+" "+b+"----"+d)

    else:
        exit()

elif entrada == '2':
    print(' ')
    while True:
        subject = input('Ingrese el titulo: ')
        print(' ')
        body = input('Ingrese el mensaje: ')
        print(' ')
        print('Vista previa del mensaje: ')
        print(' ')
        print(subject,body,sep='\n')
        print(' ')
        entrada2 = input("Desea reescribir el mensaje?: (si/no) ")
        print(' ')
        if entrada2 == "no":
            break
    entrada=input("Desea enviar los correos? (si/no): ")
    print(' ')
    if entrada == 'si':
        for (a,b,c,d,e) in zip(apellidopaterno_nopagados,apellidomaterno_nopagados,nombres_nopagados,correos_nopagados,estadodepago_nopagados):
            filename = (filename2+' '+c+' '+a+'.pdf')
            subject
            body
            sender_email = "mensualidades@colegiomanuelrodriguez.cl"
            receiver_email = d

    # Create a multipart message and set headers
            message = MIMEMultipart()
            message["From"] = 'Mensualidades Colegio'
            message["To"] = d
            message["Subject"] = subject

    # Add body to email
            message.attach(MIMEText(body, "plain"))
            text = message.as_string()

    # Log in to server using secure context and send email
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(sender_email, d, text)
                print("Mensaje se ha enviado a: "+c+" "+a+" "+b+"----"+d)
    else:
        exit()
else:
    exit()
