import time
import xlrd
import json
import pandas as pd


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
    
    final_df = pd.DataFrame() # Conservar

    for i in range(1, len(listFinal_Values)):

        # Nuevos DataFrames para Limpiar
        copy_df = pd.DataFrame() # Limpiar

        copy_df = df[listFinal_Values[i]]

        copy_df = copy_df.dropna()

        copy_df = copy_df.reset_index(drop=True)

        final_df[listFinal_Values[i]] = copy_df

# Lectura del Excel.
# Abrimos el Json.
with open("C:\\Users\\everis\\Documents\\Python\\Json\\Tutorial\\data\\data.json") as data:

    # Localidad del Archivo.
    # Cargamos la Data del Json.
    excel_data = json.loads(data.read())

    # Hacemos Lectura de la Seccion del Excel_Path
    myExcel = excel_data['Excel_Path'][0]

    # Lista de los Valores en el Diccionario 
    myExcel = list(myExcel.values())

    # Abrimos el WorkBook
    wb = xlrd.open_workbook(myExcel[0])

    # Abrimos el Sheet INFO_DAT
    sheet = wb.sheet_by_index(0)

    excel_Rows = sheet.nrows

    # Obtengo en Numero de filas del DataFrame (con Cabezera)
    df_Rows = len(final_df)

    # Creamos la Nueva Columna 'Grupo' llena de Ceros
    final_df[b'Grupo'] = 0

    print("INICIO")

    time.sleep(5)

    # Recorremos TODAS las Filas del Excel SIN Cabezeras
    for i in range(1, excel_Rows):

        # Recorremos TODAS las Filas del DF SIN Cabezeras
        for j in range(1, df_Rows):

            # Comparamos el Excel contra el DF 
            if ( sheet.cell_value(i, 1) == final_df.iloc[j][b'name:'] ):
                # Si existe el Match, se hace la Relacion entre el Grupo y el Nombre
                # Agregamos un Valor a una Fila/Columna especifica. 
                final_df.loc[j, b'Grupo'] = sheet.cell_value(i, 0)
                # Si no Posee Grupo alguno, El Nombre se queda por Default.   

    print("FIN :)")

    txtFile.write(final_df.to_string())

    txtFile.close()

    # Guardamos el DF en Excel. 
    #final_df.to_excel(r'C:\\Users\\everis\\Documents\\Python\\Json\\Tutorial\\data\\Results.xlsx')

    print(final_df)