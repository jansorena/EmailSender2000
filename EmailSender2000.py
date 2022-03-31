import yagmail, os, csv
import pandas as pd
import PySimpleGUI as sg
from pathlib import Path
import datetime
import numpy as np

#inicializacion yagmail
yag = yagmail.SMTP('user','pass')

#inicializacion tiempo
tiempo = datetime.datetime.now()

#tema
sg.theme('DarkPurple2')
#Archivos de texto
with open(r'C:\Users\Jaime\Documents\GitHub\EmailSender2000\titulo_pagados.txt',encoding="utf8") as tituloOpen:
    subject_pagados = tituloOpen.read()
with open(r'C:\Users\Jaime\Documents\GitHub\EmailSender2000\mensaje_pagados.txt',encoding="utf8") as mensajeOpen:
    contents_pagados = mensajeOpen.read()
with open(r'C:\Users\Jaime\Documents\GitHub\EmailSender2000\titulo_pagados.txt',encoding="utf8") as tituloOpen:
    subject_nopagados = tituloOpen.read()
with open(r'C:\Users\Jaime\Documents\GitHub\EmailSender2000\mensaje_pagados.txt',encoding="utf8") as mensajeOpen:
    contents_nopagados = mensajeOpen.read()

#declaracion de arreglos
nombres_pagados, nombres_nopagados = [], []
apellidopaterno_pagados, apellidopaterno_nopagados = [], []
apellidomaterno_pagados, apellidomaterno_nopagados = [], []
email_pagados, email_nopagados = [], []
array_pagados, array_pagados_tras = [], []
array_nopagados, array_nopagados_tras = [], []

#abrir csv
data = pd.read_csv(r'C:\Users\Jaime\Documents\GitHub\EmailSender2000\PLANTILLA BOLETAS EMITIDAS 2022.csv', sep=',')

#append a los arreglos respectivos
for (a,b,c,d) in zip(data['APELLIDO PATERNO'],data['APELLIDO MATERNO'],data['NOMBRES'],data['EMAIL']):
    filename = (r'C:\Users\Jaime\Documents\GitHub\EmailSender2000\boleta '+c+' '+a+' '+b+'.pdf')
    try:
        open(filename, "rb")
        nombres_pagados.append(c)
        apellidopaterno_pagados.append(a)
        apellidomaterno_pagados.append(b)
        email_pagados.append(d)
    except FileNotFoundError:
        nombres_nopagados.append(c)
        apellidopaterno_nopagados.append(a)
        apellidomaterno_nopagados.append(b)
        email_nopagados.append(d)

array_pagados = [nombres_pagados,apellidopaterno_pagados,apellidomaterno_pagados,email_pagados]
array_pagados_tras = list(map(list, zip(*array_pagados)))
array_nopagados = [nombres_nopagados,apellidopaterno_nopagados,apellidomaterno_nopagados,email_nopagados]
array_nopagados_tras = list(map(list, zip(*array_nopagados))) 

#funcion enviar correos 
def enviarEmail(receiver,subject,contents,filename):
    yag.send(receiver,subject,contents,filename)
    sg.Print("Mensaje enviado correctamente a: "+receiver)

def cambiarDirectorio(filename):
    filename2 = Path(filename).stem
    os.rename(filename,r'C:\Users\Jaime\Documents\GitHub\EmailSender2000\PAGADOS\\'+filename2+'_'+tiempo.strftime('%d-%b-%Y')+'.pdf')
    
def interfaz():
    headings = ['Nombre','Apellido Paterno','Apellido Materno','Email']
    layoutl = [
        [sg.Table(values=array_pagados_tras,
            headings=headings,
            max_col_width=35,
            auto_size_columns=True,
            justification='right',
            num_rows=20,
            key='-Table1-',
            row_height=30)],
        [sg.Button('Enviar email pagados')],  
    ]
    layoutr = [
        [sg.Table(values=array_nopagados_tras,
            headings=headings,
            max_col_width=35,
            auto_size_columns=True,
            justification='right',
            num_rows=20,
            key='-Table2-',
            row_height=30)],
        [sg.Button('Enviar email no pagados')]
    ]
    layout = [
        [sg.T('Lista pagados                                           Lista no pagados', font='_ 18', justification='c', expand_x=True)],
        [sg.Col(layoutl), sg.Col(layoutr)],
    ]
    
    # Create the window
    window = sg.Window("EmailSender2000",layout)
    # Create an event loop
    while True:
        event, values = window.read()
        # End program if user closes window or presses the OK button
        if event == sg.WIN_CLOSED:
            break
        if event == "Enviar email pagados":
            try:
                for (a,b,c,d) in zip(apellidopaterno_pagados,apellidomaterno_pagados,nombres_pagados,email_pagados):
                    filename = (r'C:\Users\Jaime\Documents\GitHub\EmailSender2000\boleta '+c+' '+a+' '+b+'.pdf')
                    enviarEmail(d,subject_pagados,contents_pagados,filename)
                    cambiarDirectorio(filename)
            except BaseException:
                sg.Print("Hubo un error :(")
            sg.Print("Todos los mensajes se enviaron correctamente!")
            apellidopaterno_pagados.clear()
            apellidomaterno_pagados.clear()
            nombres_pagados.clear()
            email_pagados.clear()
            array_pagados.clear()
            array_pagados_tras.clear()
            window.Element('-Table1-').Update(values=array_pagados_tras)
        if event == "Enviar email no pagados":
            try: 
                for (a,b,c,d) in zip(apellidopaterno_nopagados,apellidomaterno_nopagados,nombres_nopagados,email_nopagados):                
                    filename = None
                    enviarEmail(d,subject_nopagados,contents_nopagados,filename)    
            except BaseException:
                sg.Print("Hubo un error :(")
            sg.Print("Todos los mensajes se enviaron correctamente!")
            apellidopaterno_nopagados.clear()
            apellidomaterno_nopagados.clear()
            nombres_nopagados.clear()
            email_nopagados.clear()
            array_nopagados.clear()
            array_nopagados_tras.clear()
            window.Element('-Table2-').Update(values=array_nopagados_tras)
    window.close()

def main():
    interfaz()

if __name__ == "__main__":
    main()
