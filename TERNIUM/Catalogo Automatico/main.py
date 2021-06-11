import sys
import time
import json
import logging
import threading
import pandas as pd

logging.basicConfig(level=logging.DEBUG)

def generate_File_Path(myKey):
    
    # Abrimos el Json para la Lectura de las Columnas del Grupo.
    with open("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\data\\data.json") as data:

        # Cargamos la Data del Json.
        data_JSON = json.loads(data.read())

        # Obtenemos el Diccionario del Json para la Lectura de la Ruta del .DAT.
        myFile = data_JSON[myKey][0]

        # Del Diccionario obtenemos los Valores y los transformamos a una Lista.
        path_File = list(myFile.values())

        # La Elemento 0 sera la Ruta del .DAT
        return path_File[0] 

def create_dfGroups():

    logging.info("Creacion del DF por Grupos: ")

    # Abrimos el Json para la Lectura de las Columnas del Grupo.
    with open("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\data\\data.json") as data:

        # Cargamos la Data del Json.
        data_JSON = json.loads(data.read())

        # Obtenemos el Diccionario del Json.
        data_Groups = data_JSON['dfGroups'][0]

        # Del Diccionario obtenemos los Valores y los transformamos a una Lista.
        listValues = list(data_Groups.values())

        # Lista Final que usaremos para almacenar las Columnas del DF.
        listFinal_Values = [bytes(eachValue,'utf-8') for eachValue in listValues]

        # Obtenemos el Diccionario del Json para la Lectura de la Ruta del .DAT.
        path_DAT_File = generate_File_Path('DAT')    

    # Abrimos el .DAT
    with open(path_DAT_File, 'rb') as myDAT_File:

        # Lectura del .DAT
        myDAT_File = myDAT_File.readlines()

        # Creamos un DataFrame.
        df = pd.DataFrame()

        # Recorremos la Lista con las Palabras Clave que usaremos para las Columnas del DF.
        for i in range(len(listFinal_Values)):

            # Para Cada Linea del Archivo
            for eachLine in myDAT_File:
                
                # Validacion si la Palabra Clave aparece en el Archivo.
                if (listFinal_Values[i] in eachLine):
                    
                    try:
                        # Transformamos a UTF-8.
                        newString = eachLine.decode(encoding = "utf-8")
                    except:
                        # Transformamos a ISO-8859-1
                        newString = eachLine.decode(encoding = "ISO-8859-1")
                        
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

        # Obtenemos el Diccionario del Json para la Lectura de la Ruta del PKL del Grupo.
        pkl_Path_File = generate_File_Path('txt_Data_Groups')

        dfGroups.to_pickle(pkl_Path_File)
        
        logging.info("\tFIN")

def create_dfSignals():
    
    logging.info("Creacion del DF por Signal: ")

    # Abrimos el Json.
    with open("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\data\\data.json") as data:

        # Cargamos la Data del Json.
        data_JSON = json.loads(data.read())

        # Obtenemos el Diccionario del Json.
        data_Signals = data_JSON['dfSignals'][0]

        # Del Diccionario obtenemos los Valores y los transformamos a una Lista.
        listValues = list(data_Signals.values())

        # Lista Final que usaremos para almacenar las Columnas del DataFrame.
        # Convertimos cada Valor a Byte ya que Trabajaremos con el .DAT
        listFinal_Values = [bytes(eachValue,'utf-8') for eachValue in listValues]

        # Obtenemos el Diccionario del Json.
        path_DAT_File = generate_File_Path('DAT') 

    # Abrimos el .DAT
    with open(path_DAT_File, 'rb') as myDAT_File:

        # Lectura del .DAT
        myDAT_File = myDAT_File.readlines()

        # Creamos un DataFrame.
        df = pd.DataFrame()

        # Bandera
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
                        
                        try:
                            # Transformamos a UTF-8.
                            newString = eachLine.decode(encoding = "utf-8")
                        except:
                            # Transformamos a ISO-8859-1
                            newString = eachLine.decode(encoding = "ISO-8859-1")
                        
                        # Eliminamos PRIMERO posibles Espacios Vacios 
                        newString = newString.strip()

                        # Eliminamos el Tabulador y el Salto de Linea.
                        newString = newString.strip("\r\n")

                        # Eliminamos las Comillas en caso de que Existan al Principio y Fin.
                        newString = newString.lstrip('"')
                        
                        newString = newString.rstrip('"')

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

            copy_df = copy_df.reset_index(drop = True)

            dfSignals[listFinal_Values[i]] = copy_df

        # Obtenemos el Diccionario del Json para la Lectura de la Ruta del PKL del Grupo.
        pkl_Path_File = generate_File_Path('txt_Data_Signals')

        dfSignals.to_pickle(pkl_Path_File)
        
        logging.info("\tFIN")

def create_Catalogue():
    
    logging.info("Inicia creacion de Catalogo.")

        # ---------- Abrimos los 2 DF: dfGroups y dfSignals (que se encuentran en Formatos PKL)
    Groups_Path_File = generate_File_Path('txt_Data_Groups')
    dfGroups = pd.read_pickle(Groups_Path_File)

    Signals_Path_File = generate_File_Path('txt_Data_Signals')
    dfSignals = pd.read_pickle(Signals_Path_File)

        # ---------- Abrimos el JSON y Obtenemos la Ruta del CSV que Servira como Nuestro Catalogo. 
    CSV_Path_File = generate_File_Path('CSV Catalogue Path')

        # Creamos un DataFrame para Convertirlo a CSV.
    df = pd.DataFrame()

        # Convertimos a CSV
    df_Data_Signals = df.to_csv(CSV_Path_File)


        # ---------- Creamos todos los Headers del Catalogo.
    columns_Catalogue = ['Grupo', 'Nombre_Grupo', 'Nombre_Senial', 'Channel_Number', 'Unit', 'Nombre_IBA']

        # Leemos y Agregamos las Columnas al CSV.
    df_Data_Signals = pd.read_csv(CSV_Path_File, names = columns_Catalogue)


        # ---------- Llenamos la Columna del Nombre_Grupo. -----
        # Lista donde alamcenaremos CADA UNO de los Elementos de (Name * SignalCount) en dfGroups.
    list_Nombre_Grupo = []
    for index, row in dfGroups.iterrows():
        for i in range( int(row[b'signalCount:']) ):
            list_Nombre_Grupo.append( row[b'name:'] )


        # Iteramos CADA UNO de los Nombre_Grupo y los Insertamos en las Celdas HACIA ABAJO.
        # Bandera que Inicializa en 2 porque en 1 esta el Header de la Columna.
    flag_row = 0
    for i in list_Nombre_Grupo:
        df_Data_Signals.loc[flag_row, 'Nombre_Grupo'] = i
        flag_row += 1


        # ---------- Llenamos la Columna del Nombre_Senial
    # Lista donde alamcenaremos CADA UNO de los Elementos de (Name) en dfSignals.
    list_Nombre_Senial = [row[b'name:'] for index, row in dfSignals.iterrows()]

        # Iteramos CADA UNO de los Nombre_Senial y los Insertamos en las Celdas HACIA ABAJO.
        # Bandera que Inicializa en 2 porque en 1 esta el Header de la Columna.
    flag_row = 0
    for i in list_Nombre_Senial:
        df_Data_Signals.loc[flag_row, 'Nombre_Senial'] = i
        flag_row += 1


        # ---------- Llenamos la Columna del Channel_Number.
    # Lista donde alamcenaremos CADA UNO de los Elementos de (beginchannel) en dfSignals.
    list_Channel_Number = [row[b'beginchannel:'] for index, row in dfSignals.iterrows()]

        # Iteramos CADA UNO de los Channel_Number y los Insertamos en las Celdas HACIA ABAJO.
        # Bandera que Inicializa en 2 porque en 1 esta el Header de la Columna.
    flag_row = 0
    for i in list_Channel_Number:
        if int(i) <= 999:
            df_Data_Signals.loc[flag_row, 'Channel_Number'] = '%04d' % int(i)
        else:
            df_Data_Signals.loc[flag_row, 'Channel_Number'] = i
        flag_row += 1


        # ---------- Llenamos la Columna del Unit.
    # Lista donde alamcenaremos CADA UNO de los Elementos de (unit) en dfSignals.
    list_Unit = [row[b'unit:'] for index, row in dfSignals.iterrows()]

    # Iteramos CADA UNO de los Unit y los Insertamos en las Celdas HACIA ABAJO.
    # Bandera que Inicializa en 2 porque en 1 esta el Header de la Columna.
    flag_row = 0
    for i in list_Unit:
        df_Data_Signals.loc[flag_row, 'Unit'] = i
        flag_row += 1


        # ----- Llenamos la Columna del Nombre_IBA
    # Es la Concatenacion del Nombre_Senial en Lowercase y el Channel_Number con 000
    flag_row = 0
    for i in range(len(list_Nombre_Senial)):
            # Eliminamos los Espacios y Sustituimos por (_). Convertimos a Minusculas.
            # Si el Nombre tiene DOBLE ESPACIO, se Sustituye por UNO.
            # Si la Lista tiene elemento, se trabaja aqui, sino se Continua.
        if list_Nombre_Senial[i]:
                # Nombre IBA
            name_IBA = list_Nombre_Senial[i].translate( {ord(c): "_" for c in "__ !  @#$%^&*()[]{}"";:,./<>¿?\|`~-=_+'"} )

        if (int(list_Channel_Number[i]) <= 999):
                # Si es Menor a 999, se le asignan 0000.
            number_IBA = '%04d' % int(list_Channel_Number[i])
        else:
            number_IBA = int(list_Channel_Number[i])

            #Concatemoa el Nombre Final.
        complete_name_IBA = name_IBA + "_C" + str(number_IBA)
        df_Data_Signals.loc[flag_row, 'Nombre_IBA'] = complete_name_IBA
        flag_row += 1


    # ---------- Abrimos el JSON y Obtenemos la Ruta de INFO para Convertirlo en CSV.
    df_INFO = pd.DataFrame()

    INFO_Excel_Path = generate_File_Path('INFO')

    # Leemos la INFO del Excel en Formato DataFrame
    df_INFO = pd.read_excel(INFO_Excel_Path)
        
    # Convertimos a CSV la Informacion del Excel
    INFO_CSV_Path = generate_File_Path('CSV INFO Path')

    df_INFO.to_csv(INFO_CSV_Path)

    # Leemos el CSV
    df_INFO = pd.read_csv(INFO_CSV_Path, dtype = str)

    #df_INFO['Index'] = [x for x in range(0, len(df_INFO.values))]

    # Agregamos la Columna 'flag', llena de NaN, que Servira para Marcar las Celdas que YA Transcurrimos.
    #df_INFO['flag'] = np.nan

    # ---------- Iniciamos la Creacion del Match.
    logging.info("Inicia creacion del Match.")

    df_INFO_Copy = df_INFO.copy(deep=True)

    for i in range(len(df_Data_Signals.index)):

        for j in range(len(df_INFO.index)):

            try:
                          
                # Volvemos a tratar el Problema de las "", tratando de Eliminarlos
                #text_Data = str(list_Nombre_Senial[i]).replace('"', '').strip()
                text_Data_Signals = str(df_Data_Signals.iloc[i]['Nombre_Senial']).replace('"', '').strip()
            
                text_INFO = str(df_INFO_Copy.iloc[j]['Señal']).replace('"', '').strip()

                if text_Data_Signals == text_INFO:

                    logging.info(">:" + str(i) + " + " + df_INFO_Copy.iloc[j]['Grupo'] + " + " + str(len(df_INFO_Copy)))

                    # 2.- Debo Obtener el Valor del Grupo en INFO.
                    var = df_INFO_Copy.iloc[j]['Grupo']

                    # 3.- Debo Guardar el Match en el Catlogo
                    # Eliminamos Posibles Textos que NO deban Ir.
                    var = var.replace("_text", "")

                    df_Data_Signals.loc[i, 'Grupo'] = var

                    # 4.- Modificamos el Nombre para AGREGARLE el PRIMER Numero del Grupo al Principio de este
                    # Obtengo el Indice que me Indica hasta que Cantidad de Caracteres termina el Primer Numero.
                    index_Grupo = var.find(":")

                    if not index_Grupo >= 0:
                        # Si no encuentra : que busque el .
                        index_Grupo = var.find(".")

                    # Delimito el Numero desde el Primer Caracter despues del '[' hasta mi Indice.
                    number_Grupo = var[1:index_Grupo]

                    # 5.- Obtengo el Nombre_Grupo
                    name = df_Data_Signals.iloc[i]['Nombre_Grupo']

                    # 6.- Concateno el Numero del Grupo con el Nombre.
                    final_name = str(number_Grupo) + ". " + str(name)

                    # 7.- Guardo de nueva cuenta el Nombre del Grupo en su respectiva Celda.
                    df_Data_Signals.loc[i, 'Nombre_Grupo'] = final_name

                    # 8.- Eliminamos la Fila en INFO.
                    df_INFO_Copy.drop([j], axis = 0, inplace = True)

                    # Reestructuramos los Indices
                    df_INFO_Copy = df_INFO_Copy.reset_index(drop=True) 

                    break
                
                else:

                    pass
                
                text_INFO_Copy = text_INFO.replace("[","").replace("]","").replace(".",":").split(":")

                text_INFO_Copy = [int(a) for a in text_INFO_Copy if a.isdigit()] 

                if (text_Data_Signals == '' and (len(text_INFO_Copy) <= 2 and  len(text_INFO_Copy) >= 1)) or (text_Data_Signals == '' and text_INFO == ''):

                    logging.info("VV: " + str(i) + " + " + str(text_INFO_Copy) + " + " + str(len(df_INFO_Copy)))

                    # 2.- Debo Obtener el Valor del Grupo en INFO.
                    var = df_INFO_Copy.iloc[j]['Grupo']

                    # 3.- Debo Guardar el Match en el Catalogo.
                    # Eliminamos Posibles Textos que NO deban Ir.
                    var = var.replace("_text", "")

                    df_Data_Signals.loc[i, 'Grupo'] = var
                    
                    # 4.- Al NO tener Nombre, el Grupo pasa a ser el Nombre.
                    # Nombre_Senial Debe ir sin Corchetes
                    var = var.replace("[", "").replace("]", "").replace("_text", "")
                    
                    # Lo Escribimos en la Columna de 'Nombre_Senial'.
                    df_Data_Signals.loc[i, 'Nombre_Senial'] = var

                    # 5.- Al YA tener un Nombre, procedemos a crear el Nombre_IBA 
                    name_IBA = var

                    var_2 = df_Data_Signals.iloc[i]['Channel_Number']
                    
                    if (int(var_2) <= 999 ):

                        # Si es Menor a 999, se le asignan 0000.
                        number_IBA = '%04d' % int(var_2)
                        
                        # Ya que andamos por Aqui, modificamos este mismo Numero a la Columna 'Channel_Number'
                        df_Data_Signals.loc[i, 'Channel_Number'] = number_IBA

                    else:
                        number_IBA = int(var_2)

                    #Concatemoa el Nombre Final.
                    complete_name_IBA = name_IBA + "_C" + str(number_IBA)
                    
                    # Lo Escribimos en la Columna de 'Nombre_IBA'.
                    df_Data_Signals.loc[i, 'Nombre_IBA'] = complete_name_IBA

                    # 8.- Eliminamos la Fila en INFO.
                    df_INFO_Copy.drop([j], axis = 0, inplace = True)

                    # Reestructuramos los Indices
                    df_INFO_Copy = df_INFO_Copy.reset_index(drop=True)

                    break

                else:

                    pass
                
            except IndexError as e:

                logging.info("\tError: " + str(i) + ":" + str(j) + " + " + str(text_INFO_Copy)) 
                
                break
                
    #df_Data_Signals.to_excel("C:\\Users\\everis\\Documents\\TERNIUM\\Catalogo Automatico\\results\\Catalogue_2.xlsx")
     
    parket_Path_File = generate_File_Path('Parquet Path')
    
    df_Data_Signals.to_parquet(parket_Path_File)

if __name__ == '__main__':
    
    start_time = time.time()

    dfGroups = threading.Thread(target=create_dfGroups)

    dfSignals = threading.Thread(target=create_dfSignals)

    dfGroups.start()
    dfSignals.start()

    dfGroups.join()
    dfSignals.join()
    
    create_Catalogue()
    
    print("--- %s seconds ---" % (time.time() - start_time))