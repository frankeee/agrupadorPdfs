# -*- coding: utf-8 -*-
"""
Created on Fri Mar 26 14:44:01 2021

@author: Franco
"""

import pdfplumber
import os
import openpyxl
from openpyxl import Workbook
from tkinter import *

def agregaInicial(listadeletras,listadestrings,sheet):
    i = 0
    while i < len(listadeletras):
        sheet[listadeletras[i]+'1'] = listadestrings[i]
        i+=1
        
def agregaComplejo(lista,sheet,letra,libre):
     
     i = libre
     j= 0
     while j<len(lista):
         sheet[letra+str(i)]= lista[j]
         i+=1
         j+=1

def buscaParrafo(text,i):
    h = i
    while text[h:h+2]!= ".\n":
        h+=1
    j = i
    while text[j-2:j] != ".\n":
        j-=1
    return text[j:h+1]

def buscaLinea(text,i):
    h=i
    while not text[h]=="\n":
        h+=1
    return (text[i:h],h+1)

def foundWord(text,i,listaClaves):
    for item in listaClaves:
        if text[i:i+len(item)].lower() == item:
            return (True,item)
    return (False,"")

def pasarLista(listaCompleta):
    listaRetorno = []
    for item in listaCompleta:
        i = 0
        contadorEspacios = 0
        while i < len(item) and contadorEspacios<2:
            if item[i] == " ":
                contadorEspacios +=1
            i+=1
        listaRetorno.append(item[0:i-1])
    return listaRetorno

def buscaNombre(text,i,listaCorta,listaCompleta):
    h = i
    encontro = False
    while h>= 0 and not encontro:
        n = 0
        while n < len(listaCorta):
            if listaCorta[n] == text[h-len(listaCorta[n]):h]:
                return listaCompleta[n]
            n+=1
        h-=1
    return ""

def returnEntry(arg=None):
    resultLabel.config(text="La ruta no es valida, ingrese otra ruta o presione quit para cerrar el programa")
    result = myEntry.get()
    os.chdir(result.strip())
    #resultLabel.config(text="La ruta es valida,presione quit para comenzar o entre una nueva ruta")
    root.destroy()

path =os.getcwd()
 
root = Tk()

myEntry = Entry(root, width=20)
myEntry.focus()
myEntry.bind("<Return>",returnEntry)
myEntry.pack()
resultLabel = Label(root, text = "")
resultLabel.pack(fill=X) 
enterEntry = Button(root, text= "Enter", command=returnEntry)
enterEntry.pack(fill=X)
root.geometry("+750+400")
buton =Button(root, text="Quit", command=root.destroy)
buton.pack()

root.mainloop()


if(os.getcwd()==path):
    sys.exit(-1)
        
wb = Workbook()
sheet = wb.active

listaDeClaves = ["latin america","south america","vaca muerta","argentina","chile","bolivia","ypf","pae"]
text = ""
listadeletras=["A","B","C","D"]
listadestrings = ["Palabra","Parrafo","Quien lo dijo","Compañia"]
listaPalabrasEncontradas = []
listaDeNombresEncontrados = []
listaParrafosEncontrados = []
listaCompañias = []
agregaInicial(listadeletras,listadestrings,sheet)
libre = 2
#Busca todos los archivos del directorio seleccionado
for file in os.listdir():
    #selecciona solo los pdf
    if file[len(file)-3:len(file)].lower() == 'pdf': 
        with pdfplumber.open(file) as pdf:
            listaCompleta = []
            listaCorta = []
            text =pdf.pages[1].extract_text()
            #print(text)
            i = 0
            while i < len(text):
                if i+ 21 < len(text) and text[i:i+22] ==  "CORPORATE PARTICIPANTS":
                    i=i+23
                    while text[i:i+12] != "PRESENTATION":#text[i:i+28] != "CONFERENCE CALL PARTICIPANTS":
                        tupla = buscaLinea(text,i)
                        listaCompleta.append(tupla[0])
                        i = tupla[1]
                i+=1
            listaCorta = pasarLista(listaCompleta)
            text = ""
           
            i = 0
            while i < len(pdf.pages):
                text += pdf.pages[i].extract_text()
                i+=1
            i = 0
            while i < len(text):
                tupla = foundWord(text,i,listaDeClaves)
                if tupla[0]:
                    listaPalabrasEncontradas.append(tupla[1])
                    listaDeNombresEncontrados.append(buscaNombre(text,i,listaCorta,listaCompleta))
                    listaParrafosEncontrados.append(buscaParrafo(text,i))
                i+=1
            while len(listaDeNombresEncontrados)>len(listaCompañias):
                listaCompañias.append(file[0:len(file)-4])

agregaComplejo(listaPalabrasEncontradas,sheet,"A",libre) 
agregaComplejo(listaParrafosEncontrados,sheet,"B",libre) 
agregaComplejo(listaDeNombresEncontrados,sheet,"C",libre)
agregaComplejo(listaCompañias,sheet,"D",libre)      
wb.save('transcripts.xlsx') 
    