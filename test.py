import json
import pandas as pd
import time

# Abrimos el Json.
with open("C:\\Users\\everis\\Documents\\Python\\Json\\Tutorial\\data\\data.json") as data:

    # Cargamos la Data del Json.
    data = json.loads(data.read())

    # Obtenemos el Diccionario del Json.
    data = data['AM2_MCC1_7944'][0]

    # Del Diccionario obtenemos los Valores y los transformamos a una Lista.
    listValues = list(data.values())

    # Lista Final que usaremos.
    listFinal_Values = []

    for eachValue in listValues:

        # Convertimos cada Valor a Byte ya que Trabajaremos con el .Dat
        byteData = bytes(eachValue,'utf-8')

        # Lo Guardamos en la Lista que Usaremos.
        listFinal_Values.append(byteData)

    #print("Lista Final: ", listFinal_Values)

    # Creamos el DataFrame.
    df = pd.DataFrame()

with open('C:\\Users\\everis\\Documents\\Python\\Json\\Tutorial\\AM2_MCC1_7944.dat', 'rb') as myFile:
    
    # Lectura del .DAT
    myFile = myFile.readlines()

    # Archivo para Guardar la Informacion Correcta.
    txtFile = open("C:\\Users\\everis\\Documents\\Python\\Json\\Tutorial\\Save_Data.txt", 'w')

    # Recorremos la Lista con las Palabras Clave.
    for i in range(len(listFinal_Values)):

        # Para Cada Linea del Archivo
        for eachLine in myFile:       

            # Validacion si la Palabra Clave aparece en el Archivo.
            if listFinal_Values[i] in eachLine:

                newString = eachLine.decode(encoding = "utf-8")
                
                # Eliminamos el Tabulador y el Salto de Linea.
                newString = newString.strip("\r\n")

                # Buscamos los Elementos que estan Despues de los (:)
                newString = newString.split(":")

                nameColumn = listFinal_Values[i]

                df = df.append( {nameColumn : newString[1]},  ignore_index=True)
    
    df_3 = pd.DataFrame() # Conservar

    for i in range(1, len(listFinal_Values)):

        # Nuevos DataFrames para Limpiar
        df_2 = pd.DataFrame() # Limpiar

        df_2 = df[listFinal_Values[i]]

        df_2 = df_2.dropna()

        df_2 = df_2.reset_index(drop=True)

        df_3[listFinal_Values[i]] = df_2

    print(df_3)

    txtFile.write(df_3.to_string())

    txtFile.close()