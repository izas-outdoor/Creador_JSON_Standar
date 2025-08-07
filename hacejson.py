import os
import numpy as np
import pandas as pd
import openpyxl
import csv
import xlrd
import json
excels=os.listdir("Excels")
print(excels)
i=1

#Crear los csv
for e in excels:
     ex="Excels/"+e
     #Pa los excel
     if e.endswith(".xlsx"):   
          wb = openpyxl.load_workbook(ex, data_only=True)
          sh = wb.active
          with open('resultados/'+e+'.csv',"w", encoding='utf-8' ) as f:  
               c = csv.writer(f)
               
               for r in sh.rows:               
                    linea=[]
                    
                    for a in r:
                         
                         try:
                              if isinstance(a.value, float):
                                   ro=round(a.value,2)
                                   linea.append(ro)
                              elif a.value != None:
                                   linea.append(a.value)
                         except Exception:
                              pass

                    if len(linea)>1:
                         print(linea)
                         c.writerow(linea)
                    


     #Pa los archivos xls                    
     if e.endswith(".xls"):   
          wb = xlrd.open_workbook(ex)
          sh = wb.sheet_by_index(0)
          with open('resultados/'+e+'.csv',"w") as f:  
               c = csv.writer(f)
               for n in range(sh.nrows):
                    r=sh.row(n)               
                    linea=[]
                    for a in r:
                         try:
                              if isinstance(a.value, float):
                                   ro=round(a.value,2)
                                   linea.append(ro)
                              elif a.value != None:
                                   linea.append(a.value)
                         except Exception:
                              pass
                    if len(linea)>1:
                         c.writerow(linea)

#de csv a json

ceseuves=os.listdir("resultados")

#lee carpeta de los csv
for ce in ceseuves:
     
     #filtra para solo leer los csv
     if ce.endswith(".csv"):

          #escribir listas para crear los objetos
          with open('resultados/'+ce,"r") as f:
               reader = csv.reader(f)
               data=[]
               #crear lista de tallas
               tallas=[]
               CINTURA=[]
               CADERAS=[]
               MUSLOS=[]
               TOTALES=[]
               HOMBROS=[]
               PECHOS=[]
               MANGAS=[]
               for row in reader:
                    i=0
                    if len(tallas)>1:
                         cont=len(row) - len(tallas)
                         cont=cont
                         if len(row)>3:
                              #Añadir cinturas
                              if((("CINTURA")in row[1] or ("CINTURA")in row[2])and(i==0)):
                                   i+=1
                                   while cont < len(row):
                                        CINTURA.append(row[cont])
                                        cont+=1
                              
                              #Añadir caderas
                              elif(("CADERA") in row[1] or ("CADERA")in row[2]):
                                   while cont < len(row):
                                        CADERAS.append(row[cont])
                                        cont+=1

                              #Añadir muslos
                              elif(("MUSLO") in row[1] or ("MUSLO") in row[2]):
                                   while cont < len(row):
                                        MUSLOS.append(row[cont])
                                        cont+=1
                              
                              #Añadir total
                              elif(("TOTAL") in row[1] or ("TOTAL") in row[2]):
                                   while cont < len(row):
                                        TOTALES.append(row[cont])
                                        cont+=1
                              
                              #Añadir hombros
                              elif(("HOMBROS") in row[1] or ("HOMBROS") in row[2]):
                                   while cont < len(row):
                                        HOMBROS.append(row[cont])
                                        cont+=1
                              
                              #Añadir pecho
                              elif(("PECHO") in row[1] or ("PECHO") in row[2]):
                                   while cont < len(row):
                                        PECHOS.append(row[cont])
                                        cont+=1

                              #Añadir mangas
                              elif(("MANGA") in row[1]or ("MANGA") in row[2]):
                                   while cont < len(row):
                                        MANGAS.append(row[cont])
                                        cont+=1
                    elif "M" in row:
                          for num in range(len(row)):
                                if not row[num].endswith("H"):
                                      if not row[num].endswith("+"):
                                            if not row[num].endswith("-"):
                                                  tallas.append(row[num])
                                

               if len(MUSLOS)>0:
                    for i in range(12):
                         try:
                              data.append({"Tallas":tallas[i], "CINTURA":float(CINTURA[i]), "CADERA":float(CADERAS[i]), "MUSLO":float(MUSLOS[i]), "TOTAL": float(TOTALES[i])})
                         except:
                              pass
                    
               elif  len(HOMBROS)>0:
                    for i in range(12):
                         try:
                              data.append({"Tallas":tallas[i], "HOMBROS":float(HOMBROS[i]), "PECHO":float(PECHOS[i]), "CINTURA":float(CINTURA[i]), "MANGA":float(MANGAS[i]), "TOTAL":float(TOTALES[i])})
                         except:
                              pass
       
                    
               with open("resultados/"+ce+".json","w")as file:
                                   json.dump(data, file)
