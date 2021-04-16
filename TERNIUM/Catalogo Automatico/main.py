import time
import xlrd
import json
import pandas as pd

def create_dfGroups():

    print("Creacion del DF por Grupos: ")

    # Abrimos el Json.
    with open("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\data\\data.json") as data:

        # Cargamos la Data del Json.
        data_Groups = json.loads(data.read())

        # Obtenemos el Diccionario del Json.
        data_Groups = data_Groups['dfGroups'][0]

        # Del Diccionario obtenemos los Valores y los transformamos a una Lista.
        listValues = list(data_Groups.values())

        # Lista Final que usaremos para almacenar las Columnas del DF.
        listFinal_Values = []

        for eachValue in listValues:

            # Convertimos cada Valor a Byte ya que Trabajaremos con el .Dat
            byteData = bytes(eachValue,'utf-8')

            # Lo Guardamos en la Lista que Usaremos.
            listFinal_Values.append(byteData)

        print("Lista Final de Valores: ", listFinal_Values)

    # Abrimos el .DAT
    with open('C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\data\\AM2_MCC1_7944.dat', 'rb') as myDAT_File:

        # Lectura del .DAT
        myDAT_File = myDAT_File.readlines()

        # Archivo para Guardar la Informacion Correcta.
        txtFile = open("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\results\\Data_Groups.txt", 'w+')

        # Creamos un DataFrame.
        df = pd.DataFrame()

        # Recorremos la Lista con las Palabras Clave que usaremos para las Columnas del DF.
        for i in range(len(listFinal_Values)):

            # Para Cada Linea del Archivo
            for eachLine in myDAT_File:
                
                # Validacion si la Palabra Clave aparece en el Archivo.
                if (listFinal_Values[i] in eachLine):
                    
                    # Transformamos a UTF-8.
                    newString = eachLine.decode(encoding = "utf-8")
                    
                    # Eliminamos el Tabulador y el Salto de Linea.
                    newString = newString.strip("\r\n")

                    # Buscamos los Elementos que estan Despues de los (:)
                    newString = newString.split(":")

                    nameColumn = listFinal_Values[i]

                    df = df.append( {nameColumn : newString[1]},  ignore_index=True)
                elif eachLine == b'endheader:\r\n':
                    break
        
        # DF Final que Usaremos
        dfGroups = pd.DataFrame() 

        for i in range(len(listFinal_Values)):

            # Nuevos DataFrames para Limpiar
            copy_df = pd.DataFrame() 

            copy_df = df[listFinal_Values[i]]

            copy_df = copy_df.dropna()

            copy_df = copy_df.reset_index(drop=True)

            dfGroups[listFinal_Values[i]] = copy_df

        txtFile.write(dfGroups.to_string())

        print("FIN")

def create_dfSignals():
    
    print("Creacion del DF por Signal: ")

    # Abrimos el Json.
    with open("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\data\\data.json") as data:

        # Cargamos la Data del Json.
        data_Signals = json.loads(data.read())

        # Obtenemos el Diccionario del Json.
        data_Signals = data_Signals['dfSignals'][0]

        # Del Diccionario obtenemos los Valores y los transformamos a una Lista.
        listValues = list(data_Signals.values())

        # Lista Final que usaremos para almacenar las Columnas del DF.
        listFinal_Values = []

        for eachValue in listValues:

            # Convertimos cada Valor a Byte ya que Trabajaremos con el .Dat
            byteData = bytes(eachValue,'utf-8')

            # Lo Guardamos en la Lista que Usaremos.
            listFinal_Values.append(byteData)

        print("Lista Final de Valores: ", listFinal_Values)

    # Abrimos el .DAT
    with open('C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\data\\AM2_MCC1_7944.dat', 'rb') as myDAT_File:

        # Lectura del .DAT
        myDAT_File = myDAT_File.readlines()

        # Archivo para Guardar la Informacion Correcta.
        txtFile = open("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\results\\Data_Signals.txt", 'w+')

        # Creamos un DataFrame.
        df = pd.DataFrame()

        blnFlag = False

        # Recorremos la Lista con las Palabras Clave que usaremos para las Columnas del DF.
        for i in range(len(listFinal_Values)):

            # Para Cada Linea del Archivo
            for eachLine in myDAT_File:
                                 
                if (eachLine == b'endheader:\r\n'):
                    blnFlag = True

                # Si se Activo la Bandera
                if (blnFlag):

                    # Validacion si la Palabra Clave aparece en el Archivo.
                    if (listFinal_Values[i] in eachLine):

                        # Transformamos a UTF-8.
                        newString = eachLine.decode(encoding = "utf-8")
                        
                        # Eliminamos el Tabulador y el Salto de Linea.
                        newString = newString.strip("\r\n")

                        # Buscamos los Elementos que estan Despues de los (:)
                        newString = newString.split(":")

                        nameColumn = listFinal_Values[i]

                        df = df.append( {nameColumn : newString[1]},  ignore_index=True)
                    elif eachLine == b'endASCII:\r\n':
                        break

            blnFlag = False

        # DF Final que Usaremos
        dfSignals = pd.DataFrame() 

        for i in range(len(listFinal_Values)):

            # Nuevos DataFrames para Limpiar
            copy_df = pd.DataFrame() 

            copy_df = df[listFinal_Values[i]]

            copy_df = copy_df.dropna()

            copy_df = copy_df.reset_index(drop=True)

            dfSignals[listFinal_Values[i]] = copy_df

        txtFile.write(dfSignals.to_string())

        print("FIN")

if __name__ == '__main__':

    create_dfGroups()

    create_dfSignals()