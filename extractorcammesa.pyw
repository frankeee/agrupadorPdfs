# -*- coding: utf-8 -*-
"""
Created on Tue Jan 26 16:01:28 2021

@author: Franco
"""

def returnEntry(arg=None):
    resultLabel.config(text="La ruta no es valida, ingrese otra ruta o presione quit para cerrar el programa")
    result = myEntry.get()
    os.chdir(result.strip())
    #resultLabel.config(text="La ruta es valida,presione quit para comenzar o entre una nueva ruta")
    root.destroy()

def BuscaVariable(indice,texto,string,mode):
    temp = ""
    cant_numeros = 0
    estoyennum = False
    if indice+len(string) < len(texto) and text[indice:indice+len(string)] == string:
        h = indice+len(string)
        while cant_numeros/2!= mode:
            if not texto[h].isspace():
                if not estoyennum:
                    cant_numeros+=1
                    estoyennum=True
            if texto[h].isspace():
                if estoyennum:
                    estoyennum=False
            h+=1
        h-=1
        r=h
        while not texto[r].isspace():
            r+=1
        temp = texto[h:r]
    return temp

def agregaSimple(valor,sheet,letra,string):
    sheet[letra+str(1)] = string
    sheet[letra+str(2)] = valor

from tkinter import *                
import openpyxl
from zipfile import ZipFile
from io import BytesIO
import os
import requests
from bs4 import BeautifulSoup
import datetime 
import pdfplumber
import sys

pathIsCorrect = False
path =os.getcwd()

root = Tk()

myEntry = Entry(root, width=20)
myEntry.focus()
myEntry.bind("<Return>",returnEntry)
myEntry.pack()
resultLabel = Label(root, text = "")
resultLabel.pack(fill=X) 
# Create the Enter button
enterEntry = Button(root, text= "Enter", command=returnEntry)
enterEntry.pack(fill=X)
 
# Create and emplty Label to put the result in

 
 
root.geometry("+750+400")


buton =Button(root, text="Quit", command=root.destroy)
buton.pack()

root.mainloop()


if(os.getcwd()==path):
    sys.exit(-1)
path = os.getcwd()
#os.chdir(path)
if not os.path.isdir('pds'):
    os.mkdir('pds')
#strings = ["DEMANDA  TOTAL MMm³","Demanda Prioritaria (R+SGP)","Ajuste de Demanda (conforme a proyecciones del Sist. de Transporte)","Gas Combustible","GNC","Industria (P3+GU+Grandes C.)","Usinas dentro del Sistema de Transporte","GNL BB (llenado)","Exportaciones en el Sistema de Transporte TGN","Exportaciones en el Sistema de Transporte TGS","Demanda Tierra del Fuego","Usinas en Boca Pozo","Gas Pacífico","Mega","Refinor","Gas Atacama","INYECCIONES (TGN + TGS) MMm³","PRODUCTORES - INYECCION NACIONAL","Sur","Neuba I","Neuba II","Patagónico","Norte (incluye Bolivia)","Neuquén","BOLIVIA","Inyección desde Chile","PEAK SHAVING","INYECCIÓN BUQUE ESCOBAR","INYECCIÓN PROPANO-AIRE (PIPA)"]#,"STOCK (TGN+TGS)","DELTA MMm³"]
#valores=  [None]*len(strings)
letras = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA"]

for file in os.listdir():
    if file[len(file)-3:len(file)] == 'pdf':
        date= file[9:11]+' '+file[7:9]+' ' +file[3:7]
        dateforbiformat = int(file[3:7]+file[7:9])
        day_name= ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday','Sunday']
        day = datetime.datetime.strptime(date, '%d %m %Y').weekday()
        day2= datetime.datetime.strptime(date, '%d %m %Y')-datetime.timedelta(days=1)
        times = 2
        if (day_name[day]) == 'Monday':
            times = 4
            day2 -= datetime.timedelta(days=2)
        #daytemp = str(day2)
        #dateforbiformat = int(daytemp[0:4]+daytemp[5:7])
        mode = 1
        while mode < times:
            
            strings = ["Fecha","Demanda Prioritaria (R+SGP)","Ajuste de Demanda (conforme a proyecciones del Sist. de Transporte)","Gas Combustible","GNC","Industria (P3+GU+Grandes C.)","Usinas dentro del Sistema de Transporte","GNL BB (llenado)","Exportaciones en el Sistema de Transporte TGN","Exportaciones en el Sistema de Transporte TGS","Demanda Tierra del Fuego","Usinas en Boca Pozo","Gas Pacífico","Mega","Refinor","Gas Atacama","Sur","Neuba I","Neuba II","Patagónico","Norte (incluye Bolivia)","Neuquén","BOLIVIA","Inyección desde Chile","PEAK SHAVING","INYECCIÓN BUQUE ESCOBAR","INYECCIÓN PROPANO-AIRE (PIPA)"]#,"STOCK (TGN+TGS)","DELTA MMm³"]
            valores=  [None]*len(strings)
            with pdfplumber.open(file) as pdf:
                daytemp = str(day2)
                dateforbiformat = (daytemp[5:7]+"/"+daytemp[8:10]+"/"+daytemp[0:4])
                #print(day2)
                #print(dateforbiformat)
                for page in pdf.pages:
                    text = page.extract_text()
                    i = 0
                    while i < len(text):
                        n=0
                        while n < len(strings):
                            if BuscaVariable(i,text,strings[n],mode)!= "" and valores[n]==None:
                               valores[n] = BuscaVariable(i,text,strings[n],mode)
                            n+=1
                        i+=1
                    valores[0] = dateforbiformat
                    wb = openpyxl.Workbook() 
                    sheet = wb.active  
                    h = 0
                    while h < len(strings):
                        agregaSimple(valores[h],sheet,letras[h],strings[h])
                        h+=1
                    os.chdir(path+"\\pds")
                    wb.save(file[0:-4]+str(mode)+'.xlsx')
                    os.chdir(path)
                    
                    
            mode+=1
            day2+= datetime.timedelta(days=1)

#Arranca Cammesa

req = requests.get("https://cammesaweb.cammesa.com/sintesis-mensual/")
soup = BeautifulSoup(req.text, "lxml")
juan = soup.find_all('a')
mylink=""
for text in juan:
    if text.get('data-downloadurl')!= None:
        mylink = text.get('data-downloadurl')
        break
res = requests.get(mylink)
res.raise_for_status()
#playFile = open("oferta.xlsx","wb")
z = ZipFile(BytesIO(res.content))
z.extractall()
lista = z.namelist()
for files in lista:
    if files[0:11] == "BASE_OFERTA":
        if os.path.exists("BASE_OFERTA.xlsx"):
            os.remove("BASE_OFERTA.xlsx")
        os.rename(os.getcwd()+"\\"+files,os.getcwd()+"\\BASE_OFERTA.xlsx")
    if files[0:12] == "BASE_DEMANDA":
        if os.path.exists("BASE_DEMANDA.xlsx"):
            os.remove("BASE_DEMANDA.xlsx")
        os.rename(os.getcwd()+"\\"+files,os.getcwd()+"\\BASE_DEMANDA.xlsx")
    if files[0:16] == "BASE_ADICIONALES":
        if os.path.exists("BASE_ADICIONALES.xlsx"):
            os.remove("BASE_ADICIONALES.xlsx")
        os.rename(os.getcwd()+"\\"+files,os.getcwd()+"\\BASE_ADICIONALES.xlsx")
    
        
