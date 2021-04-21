import sys
import time
import xlrd
import json
import openpyxl
from openpyxl import load_workbook
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
                        
                        # Eliminamos posibles Espacios Vacios 
                        newString = newString.strip()

                        # Eliminamos el Tabulador y el Salto de Linea.
                        newString = newString.strip("\r\n")

                        # Eliminamos las Comillas en caso de que Existan al Principio y Fin.
                        newString = newString.rstrip('\"')

                        newString = newString.lstrip('\"')

                        newString = newString.strip('\"')

                        # Buscamos los Elementos que estan Despues de los (:)
                        # El Split (char, 1) SOLO dividira hasta la PRIMERA APARICION del Char.
                        newString = newString.split(":", 1)

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

    # ----- Creamos todos los Headers del Catalogo.
    columns_Catalogue = ['Index', 'Grupo', 'Nombre_Grupo', 'Nombre_Senial', 'Channel_Number', 'Unit']
    
    for i in range(1, len(columns_Catalogue)):
        ## NOTA: El Excel empieza con la Celda (1,1)
        name_Columns_Catalogue = sheet.cell(row = 1, column = i)
        name_Columns_Catalogue.value = columns_Catalogue[i]

    # ----- Llenamos la Columna del Nombre_Grupo. -----
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

    # ----- Llenamos la Columna del Nombre_Senial
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

    # ----- Llenamos la Columna del Channel_Number.
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

    # ----- Llenamos la Columna del Unit.
    # Lista donde alamcenaremos CADA UNO de los Elementos de (unit) en dfSignals.
    list_Unit = []

    for index, row in dfSignals.iterrows():
        list_Unit.append(row[b'unit:'])

    # Bandera que Inicializa en 2 porque en 1 esta el Header de la Columna.
    flag_row = 2   
    
    for i in range(len(list_Unit)):
        value_Column_Unit = sheet.cell(row = flag_row, column = 5)
        value_Column_Unit.value = list_Unit[i]
        flag_row += 1

    # ----------------------------------------------#

    print("\nInicia creacion del Match.")

    # Abrimos el Json.
    with open("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\data\\data.json") as data:

        # Cargamos la Data del Json.
        INFO = json.loads(data.read())

        # Obtenemos el Diccionario del Json.
        INFO = INFO['INFO'][0]
        
        # Del Diccionario obtenemos los Valores y los transformamos a una Lista.
        INFO_Path = list(INFO.values())

        # Cargamos el INFO Excel para comenzar el Match
        INFO_wb = load_workbook(INFO_Path[0])
        
        # Get workbook active sheet from the active attribute.
        INFO_sheet = INFO_wb.active

        # Contamos el Total de Filas de CADA Excel.
        INFO_Row_Count = INFO_sheet.max_row

        Catalogue_Row_Count = sheet.max_row

        # Validamos si Existe un Desfase en la Cantidad de Informacion.
        if (INFO_Row_Count == Catalogue_Row_Count):
            print("> Existe la Misma Cantidad de Informacion.")
        else:
            print("> Existe un Desfase en la Informacion.")
            sys.exit()

        # Catalogo -> INFO -> Match
        for i in range(2, Catalogue_Row_Count):

            for j in range(2,  INFO_Row_Count):

                # Validamos si la Columna 3 esta de INFO Vacia.
                # Si esta VACIA, se valida la Senial.
                # Sino se pasa a la Siguiente Linea.
                ######## NOTA: El Excel empieza con la Celda (1, 1)
                myEmpty_cell = INFO_sheet.cell(row = j, column = 3)

                # Si la Celda SI ESTA VACIA.
                if myEmpty_cell.value is None:

                    # ----- Comparamos la Senial de INFO contra Nombre_Senial del Catalogo.
                    if (INFO_sheet.cell(row = j, column = 2).value == sheet.cell(row = i, column = 3).value): 

                        # 1.- Colocar una Bandera a INFO para que ya no cuente ese Match.
                        myEmpty_cell.value = 1

                        # 2.- Debo Obtener el Valor del Grupo en INFO.
                        match_Group = INFO_sheet.cell(row = j, column = 1)
                        match_Group = match_Group.value

                        # 3.- Debo Guardar el Match en el Catlogo
                        exact_Match = sheet.cell(row = i, column = 1)
                        exact_Match.value = match_Group

                        # 4.- Guardamos los cambios en INFO.
                        INFO_wb.save(INFO_Path[0])

                        break

                else:
                    # Si la Celda NO ESTA VACIA.
                    pass

        # -- Analizamos las Posiciones Vacias.
        print("\nAnalizamos las Posiciones Vacias.")
        # Catalogo -> INFO -> Match
        for i in range(2, Catalogue_Row_Count):

            for j in range(2,  INFO_Row_Count):

                # Validamos si la Columna 3 esta de INFO Vacia.
                # Si esta VACIA, se valida la Senial.
                # Sino se pasa a la Siguiente Linea.
                ######## NOTA: El Excel empieza con la Celda (1, 1)
                myEmpty_cell = INFO_sheet.cell(row = j, column = 3)

                if (myEmpty_cell.value is None) and (sheet.cell(row = i, column = 3).value in [None,'None' ,'']):
                    
                    print(i, j, ": ", myEmpty_cell.value, " == ", sheet.cell(row = i, column = 3).value)

                    # 1.- Colocar una Bandera a INFO para que ya no cuente ese Match.
                    myEmpty_cell.value = 1

                    # 2.- Debo Obtener el Valor del Grupo en INFO.
                    match_Group = INFO_sheet.cell(row = j, column = 1)
                    match_Group = match_Group.value

                    # 3.- Debo Guardar el Match en el Catlogo
                    exact_Match = sheet.cell(row = i, column = 1)
                    exact_Match.value = match_Group

                    # 4.- Guardamos los cambios en INFO.
                    INFO_wb.save(INFO_Path[0])

                    break
                else:
                    pass

    # ----------------------------------------------#

    # Guardamos el Archivo.
    wb.save("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\results\\Catalogue.xlsx")

    wb.close()

    INFO_wb.close()

    print("FIN")

    return wb

if __name__ == '__main__':

    dfGroups = create_dfGroups()

    dfSignals = create_dfSignals()

    catalogue_WB = create_Catalogue(dfGroups, dfSignals)