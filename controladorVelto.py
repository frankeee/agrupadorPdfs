# -*- coding: utf-8 -*-
"""
Created on Thu Mar 25 16:49:15 2021

@author: Franco
"""
import pyodbc
#import mysql.connector.connection
server = 'tcp:voltosqlserver.database.windows.net,1433'
database = 'voltosqlDB'
username = 'voltouser'
password = 'Volto1998'   
driver= '{ODBC Driver 17 for SQL Server}'


conn=  pyodbc.connect('DRIVER='+driver+';SERVER='+server+';PORT=3306;DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = conn.cursor() 
cursor.execute("DELETE FROM Pedido")
cursor.commit()
cursor.execute("DELETE FROM PedidoEntregado")
cursor.commit()
cursor.execute("DELETE FROM Persona")
cursor.commit()
cursor.close()
