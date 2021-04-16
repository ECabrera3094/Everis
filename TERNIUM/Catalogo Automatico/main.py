import time
import xlrd
import json
import openpyxl
import pandas as pd

def create_dfGroups():

    print("\nCreacion del DF por Grupos: ")

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

        return dfGroups

def create_dfSignals():
    
    print("\nCreacion del DF por Signal: ")

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

        return dfSignals

def create_Catalogue(dfGroups, dfSignals):
    
    print("\nInicia creacion de Catalogo.")

    # Call a Workbook() function of openpyxl to create a new blank Workbook object.
    wb = openpyxl.Workbook()

    # Get workbook active sheet from the active attribute.
    sheet = wb.active

    sheet.title = "Catalogo"

    # -- Creamos todos los Headers del Catalogo.
    columns_Catalogue = ['Index', 'Grupo', 'Nombre_Grupo', 'Nombre_Senial', 'Channel_Number']
    
    for i in range(1, len(columns_Catalogue)):
        ## NOTA: El Excel empieza con la Celda (1,1)
        name_Columns_Catalogue = sheet.cell(row = 1, column = i)
        name_Columns_Catalogue.value = columns_Catalogue[i]

    # -- Llenamos la Columna del Nombre_Grupo.
    # Lista donde alamcenaremos CADA UNO de los Elementos de (Name * SignalCount) en dfGroups.
    list_Nombre_Grupo = [] 

    for index, row in dfGroups.iterrows():
        for i in range( int(row[b'signalCount:']) ):
            list_Nombre_Grupo.append( row[b'name:'] )
            
    # Bandera que Inicializa en 2 porque en 1 esta el Header de la Columna.
    flag_row = 2

    # Iteramos CADA UNO de los Nombres del Grupo y los Insertamos en las Celdas HACIA ABAJO.
    for i in range(len(list_Nombre_Grupo)):
        value_Column_Nombre_Grupo = sheet.cell(row = flag_row, column = 2)
        value_Column_Nombre_Grupo.value = list_Nombre_Grupo[i]
        flag_row += 1

    # -- Llenamos la Columna del Nombre_Senial
    # Lista donde alamcenaremos CADA UNO de los Elementos de (Name) en dfSignals.
    list_Nombre_Senial = []

    for index, row in dfSignals.iterrows():
        list_Nombre_Senial.append(row[b'name:'])

    # Bandera que Inicializa en 2 porque en 1 esta el Header de la Columna.
    flag_row = 2   

    # Iteramos CADA UNO de los Nombres del Grupo y los Insertamos en las Celdas HACIA ABAJO.
    for i in range(len(list_Nombre_Senial)):
        value_Column_Nombre_Senial = sheet.cell(row = flag_row, column = 3)
        value_Column_Nombre_Senial.value = list_Nombre_Senial[i]
        flag_row += 1

    # -- Llenamos la Columna del Channel_Number.
    # Lista donde alamcenaremos CADA UNO de los Elementos de (beginchannel) en dfSignals.
    list_Channel_Number = []

    for index, row in dfSignals.iterrows():
        list_Channel_Number.append(row[b'beginchannel:'])

    # Bandera que Inicializa en 2 porque en 1 esta el Header de la Columna.
    flag_row = 2   

    for i in range(len(list_Channel_Number)):
        value_Column_Channel_Number = sheet.cell(row = flag_row, column = 4)
        value_Column_Channel_Number.value = list_Channel_Number[i]
        flag_row += 1

    # Guardamos el Archivo.
    wb.save("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\results\\Catalogue.xlsx")

    print("FIN")


if __name__ == '__main__':

    dfGroups = create_dfGroups()

    dfSignals = create_dfSignals()

    create_Catalogue(dfGroups, dfSignals)