# -*- coding: utf-8 -*-
"""
Created on Tue Mar  9 12:45:36 2021

@author: Franco
"""
def agregaInicial(sheet):
    sheet["A1"] = "Fecha"
    sheet["B1"] = "Planta"
    sheet["C1"]= "Material"
    sheet["D1"]= "Currency"
    sheet["E1"]= "Standard Price"
    sheet["F1"]= "Current Negotiation"
    sheet["G1"]= "Delta (a-b)/a"
    sheet["H1"]= "Mid-Term Reposition"
    sheet["I1"]= "Delta (b-c))b"
    sheet["J1"]= "Current Indicator"
    sheet["K1"]= "Mid-Termi Indicator"
    sheet["L1"]= "Description"


def agregaColumnadeElems(sheet,lista,letra,ultimoIngresado):
    i = ultimoIngresado
    for items in lista:
        sheet[letra+str(i)] = items
        i+=1
    return i
def trimeaData(listaacortar,cortaPorabajo,cortaPorArriba,planta):
    indiceSuperior = None
    indiceInferior = None
    i = 0
    while i < len(listaacortar):
        if i + len(planta) < len(listaacortar) and listaacortar[i:i+len(planta)] == planta:
            h = i
            encontroSuperior = False
            while not encontroSuperior:
                if listaacortar[h-len(cortaPorArriba):h] == cortaPorArriba:
                    indiceSuperior = h-len(cortaPorArriba)
                    encontroSuperior = True
                h-=1
            h = i
            encontroInferior = False
            while not encontroInferior:
                if listaacortar[h:h+len(cortaPorabajo)] == cortaPorabajo:
                    indiceInferior = h+len(cortaPorabajo)
                    encontroInferior = True
                h+=1
            break
        i+=1
    return listaacortar[indiceSuperior:indiceInferior]

def uneFilas(lista,planta):
    lista = separaItems(lista)
    copia2 = []
    for item in lista:
        if len(item)!=0 and item!= planta:
           if item[0].isdigit():
               copia2[len(copia2)-1]+= " "+item
           else:
               copia2.append(item)
    return copia2

def extraeFecha(text):
    i = 0
    while i < len(text):
        if text[i:i+24] == "Raw Material Thermometer" and text[i:i+35] != "Raw Material Thermometer DDP Prices": 
            return (text[i+24:i+32])
            
        i+=1


def buscaSiguienteNumero(text,h):
    j=h
    while h<len(text) and (text[h].isdigit() or text[h]=="," or text[h]=="." or text[h]=="-" or text[h]=="%"):
        h+=1
    tupla = (text[j:h],h)
    return tupla

def equipararListas(variables):
    maxLen = 0
    for items in variables:
        if len(items) > maxLen:
            maxLen = len(items)
    i = 0
    while i < len(variables):
        if len(variables[i]) < maxLen:
            variables[i].append(" ")
        i+=1

def separaItems(lista):
    listaFinal = []
    items = ""
    i = 0
    while i < len(lista):
        if lista[i] == "\n":
            listaFinal.append(items)
            items = ""
        else:
            items+= lista[i]
        i+=1
    listaFinal.append(items)
    return listaFinal
#variables = [plant,material,currency,standardPrice,CurrentNegotiation,delta(a-b),midtermreposition,delta(b-c),currentIndicator,midTermIndicator,description]

def separaEnColumnas(item,variables):
    materialEncontrado = False
    numsencontrados = 4
    i = 0
    while i < len(item):
        if item[i].isdigit():
            tupla = buscaSiguienteNumero(item,i)
            variables[numsencontrados].append(tupla[0])
            numsencontrados+=1
            i = tupla[1]
        elif item[i].isalpha():
            if not materialEncontrado:
                while item[i:i+3]!= "EUR" and item[i:i+3]!= "USD":
                    i+=1
                materialEncontrado = True
                variables[2].append(item[0:i])
                h = i
                while item[i]!= " ":
                    i+=1
                variables[3].append(item[h:i])
            else:
                variables[11].append(item[i:len(item)])
                i = len(item)
        i+=1
    equipararListas(variables)

def returnEntry(arg=None):
    resultLabel.config(text="La ruta no es valida, ingrese otra ruta o presione quit para cerrar el programa")
    result = myEntry.get()
    os.chdir(result.strip())
    #resultLabel.config(text="La ruta es valida,presione quit para comenzar o entre una nueva ruta")
    root.destroy()

import os.path
from os import path
import os
from openpyxl import Workbook
import pdfplumber
from tkinter import *
import sys
import openpyxl

pat =os.getcwd()
 
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

wb = None
if(os.getcwd()==pat):
    sys.exit(-1)
if path.exists("rawMaterialAgrupado.xlsx"):
    wb =  openpyxl.load_workbook("rawMaterialAgrupado.xlsx")
else:    
    wb = Workbook()
sheet = wb.active
agregaInicial(sheet)
ultimoIngresado=2
letras=["A","B","C","D","E","F","G","H","I","J","K","L"]
for file in os.listdir():
    #selecciona solo los pdf
    if file[len(file)-3:len(file)].lower() == 'pdf': 
        with pdfplumber.open(file) as pdf:
            page = pdf.pages[1]
            lista = page.extract_tables()
            text = page.extract_text()
            fecha = extraeFecha(text)
            listaacortar  = (lista[0][0][0])
            variables = [[] for _ in range(12)]
            variables[0].append(fecha)
            cortador = separaItems(lista[2][0][1])
            cortador = cortador[0]
            listaDalmine = trimeaData(listaacortar,cortador,"a b c","Dalmine")
            listaDalmine = uneFilas(listaDalmine,"Dalmine")
            listaDalmine = listaDalmine[1:len(listaDalmine)-1]    
            variables[1].append("Dalmine")
            for item in listaDalmine:
                separaEnColumnas(item,variables)
            r = 1
            while r < len(variables[1]):
                variables[1][r] = variables[1][0]
                r+=1
            i = 2
            while i < len(lista):
                corteSuperior = separaItems(lista[i][0][1])
                corteSuperior = corteSuperior[0]
                corteInferior = separaItems(lista[i][0][10])
                corteInferior = corteInferior[len(corteInferior)-1]
                planta = lista[i][0][0]
                variables[1].append(planta)
                listor = (trimeaData(listaacortar,corteInferior,corteSuperior,planta))
                listor = uneFilas(listor,planta)
                
                for item in listor:
                   separaEnColumnas(item,variables)
                ii = 0
                itemnoNulo = None
                while ii < len(variables[1]):
                    if variables[1][ii] == " ":
                        variables[1][ii] = itemnoNulo
                    else:
                        itemnoNulo = variables[1][ii]
                    ii+=1
                
                
                i+=1
            t = 1
            while t < len(variables[0]):
                variables[0][t] = variables[0][0]
                t+=1
            n =0
            while n < len(variables):
                agregaColumnadeElems(sheet,variables[n],letras[n],ultimoIngresado)
                n+=1
            ultimoIngresado+= len(variables[1])+1
        wb.save("rawMaterialAgrupado.xlsx")