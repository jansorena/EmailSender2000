import yagmail, os, csv
import pandas as pd
from pathlib import Path
import datetime

#inicializacion yagmail
yag = yagmail.SMTP('user','pass')

#inicializacion tiempo
tiempo = datetime.datetime.now()
#Archivos de texto
with open('titulo_pagados.txt',encoding="utf8") as tituloOpen:
    subject_pagados = tituloOpen.read()
with open('mensaje_pagados.txt',encoding="utf8") as mensajeOpen:
    contents_pagados = mensajeOpen.read()
with open('titulo_pagados.txt',encoding="utf8") as tituloOpen:
    subject_nopagados = tituloOpen.read()
with open('mensaje_pagados.txt',encoding="utf8") as mensajeOpen:
    contents_nopagados = mensajeOpen.read()

#funcion enviar correos
def enviarEmail(receiver,subject,contents,filename):
    yag.send(receiver,subject,contents,filename)
    print("Mensaje enviado correctamente a: "+receiver)

def cambiarDirectorio(filename):
    filename2 = Path(filename).stem
    os.rename(filename,'PAGADOS/'+filename2+'_'+tiempo.strftime('%d-%b-%Y')+'.pdf')

#declarion de arreglos
nombres_pagados, nombres_nopagados = [], []
apellidopaterno_pagados, apellidopaterno_nopagados = [], []
apellidomaterno_pagados, apellidomaterno_nopagados = [], []
email_pagados, email_nopagados = [], []

#abrir csv
data = pd.read_csv('PLANTILLA BOLETAS EMITIDAS 2022.csv', sep=',')

#append a los arreglos respectivos
for (a,b,c,d) in zip(data['APELLIDO PATERNO'],data['APELLIDO MATERNO'],data['NOMBRES'],data['EMAIL']):
    filename = ('boleta '+c+' '+a+' '+b+'.pdf')
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

def main():
    print('-- NO PAGADOS -- La lista de correos es la siguiente :', end='\n')
    print(' ')
    for (a,b,c,d) in zip(apellidopaterno_nopagados,apellidomaterno_nopagados,nombres_nopagados,email_nopagados):
        print(c+" "+a+" "+b+"----"+d, end='\n')
    print(' ')
    print('-- PAGADOS -- La lista de correos es la siguiente :', end='\n')
    print(' ')
    for (a,b,c,d) in zip(apellidopaterno_pagados,apellidomaterno_pagados,nombres_pagados,email_pagados):
        filename = ('boleta '+c+' '+a+' '+b+'.pdf')
        print(c+" "+a+" "+b+"----"+d+"----"+"Archivo: "+filename, end='\n')
    print(' ')

    entrada = input('Desea enviar correo a:\n [1] PAGADOS\n [2] NO PAGADOS\n [3] SALIR\n')
    if entrada == '1':
        print(' ')
        while True:
            print('Vista previa del mensaje: ')
            print(' ')
            print(subject_pagados,contents_pagados,sep='\n')
            print(' ')
            entrada2 = input("Desea reescribir el mensaje?: (si/no) ")
            print(' ')
            if entrada2 == "si":
                print("Edite los archivos de texto", end='\n')
                exit()
            if entrada2 == "no":
                break
        entrada = input("Desea enviar los correos? (si/no): ")
        print(' ')
        if entrada == 'si':
            for (a,b,c,d) in zip(apellidopaterno_pagados,apellidomaterno_pagados,nombres_pagados,email_pagados):
                filename = ('boleta '+c+' '+a+' '+b+'.pdf')
                enviarEmail(d,subject_pagados,contents_pagados,filename)
                cambiarDirectorio(filename)
        else:
            exit()

    elif entrada == '2':
        print(' ')
        while True:
            print('Vista previa del mensaje: ')
            print(' ')
            print(subject_nopagados,contents_nopagados,sep='\n')
            print(' ')
            entrada2 = input("Desea reescribir el mensaje?: (si/no) ")
            print(' ')
            if entrada2 == "si":
                print("Edite los archivos de texto", end='\n')
                exit()
            if entrada2 == "no":
                break
        entrada = input("Desea enviar los correos? (si/no): ")
        print(' ')
        if entrada == 'si':
            for (a,b,c,d) in zip(apellidopaterno_nopagados,apellidomaterno_nopagados,nombres_nopagados,email_nopagados):
                filename = None
                enviarEmail(d,subject_nopagados,contents_nopagados,filename)
        else:
            exit()
    else:
        exit()

if __name__ == "__main__":
    main()
