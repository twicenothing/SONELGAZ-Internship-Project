import pandas as pd
import numpy as np
from region_funcs_gaz import *
from funcs import *
from data import *


def load_process_xl(file='Clientele DD TO 08-2024.xlsx'):
    xl = pd.ExcelFile(file)
    df = xl.parse(0)
    df = df.dropna(how='all')

    cpt = 0
    for index, row in df.iterrows():
        cpt += 1
    
    df.index = np.arange(0, cpt)

    def detect_tables(df, keywords):
        titles = []
        for index, row in df.iterrows():
            if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                titles.append(index)
        return titles
    
    keywords = ['Nombre d\'abonnés', 'résiliations', 'réabonnés']
    table_titles = detect_tables(df, keywords)
    
    # SETTING UP NOMBRE D'ABONNEES DATA FRAME
    start_row = table_titles[0] + 1
    end_row = table_titles[0] + 16
    nombre_abonne = df.iloc[start_row + 1:end_row].reset_index(drop=True)
    nombre_abonne.columns = nombre_abonne.iloc[0]
    nombre_abonne = nombre_abonne[1:]
    nombre_abonne = nombre_abonne.iloc[:, :11]
    nombre_abonne.columns.values[0] = 'placeholder'
    nombre_abonne.index = nombre_abonne['placeholder']
    del nombre_abonne['placeholder']
    nombre_abonne.index.name = None
    nombre_abonne.index.values[0] = 'déc-23'
    nombre_abonne.columns = pd.MultiIndex.from_arrays([
        ['électricité'] * 5 + ['gaz'] * 5, nombre_abonne.columns
    ])
    nombre_abonne.fillna(0, inplace=True)
    
    # Convert to int safely
    nombre_abonne.replace('-', 0, inplace=True)
    nombre_abonne = nombre_abonne.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

    index = find_repeated_index(nombre_abonne)
    
    if index is not None:
        if nombre_abonne.iloc[index].isnull().all() or (nombre_abonne.iloc[index] == 0).all():
            nombre_abonne.iloc[index:] = np.nan
        else:
            nombre_abonne.iloc[index + 1:] = np.nan
    
    # SETTING UP ACCROISSEMENT
    accroissement = nombre_abonne.diff().fillna(0)[1:]
    accroissement.loc[len(accroissement)] = accroissement.sum()
    accroissement.index.values[-1] = 'Total'
    accroissement.replace('-', 0, inplace=True)
    
    # SETTING UP APPORT
    start_row = table_titles[1] + 1
    end_row = table_titles[1] + 16
    resiliation = df.iloc[start_row + 1:end_row].reset_index(drop=True)
    resiliation.columns = resiliation.iloc[0]
    resiliation = resiliation[1:]
    resiliation.columns.values[0] = 'placeholder'
    resiliation.index = resiliation['placeholder']
    del resiliation['placeholder']
    resiliation.index.name = None
    resiliation = resiliation.iloc[:, :10]
    resiliation.columns = pd.MultiIndex.from_arrays([
        ['électricité'] * 5 + ['gaz'] * 5, resiliation.columns
    ])
    resiliation.fillna(0, inplace=True)
    resiliation = resiliation.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

    apport = accroissement + resiliation
    
    # SETTING UP APPORT NOUVEAU
    start_row = table_titles[3] + 1
    end_row = table_titles[3] + 16
    reabonne = df.iloc[start_row + 1:end_row].reset_index(drop=True)
    reabonne.columns = reabonne.iloc[0]
    reabonne = reabonne[1:]
    reabonne.columns.values[0] = 'placeholder'
    reabonne.index = reabonne['placeholder']
    del reabonne['placeholder']
    reabonne.index.name = None
    reabonne = reabonne.iloc[:, :8]
    reabonne.columns = pd.MultiIndex.from_arrays([
        ['électricité'] * 4 + ['gaz'] * 4, reabonne.columns
    ])
    reabonne.fillna(0, inplace=True)
    reabonne = reabonne.apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)

    apport_nouv = apport - reabonne

    return nombre_abonne, accroissement, apport, apport_nouv



def load_old_data_gaz(file='T.B.xlsx'):
    df = pd.read_excel(file,sheet_name='gaz')
    df = df.dropna(how='all')
    first_row = df.columns.to_list()
    df.loc[-1] = first_row  # Temporarily add the row with index -1
    df.index = df.index + 1  # Shift all indices by 1
    df = df.sort_index()
    arr= np.arange(0,59)
    df.columns = arr

    def detect_tables(df, keywords):
        titles = []
        for index, row in df.iterrows():
            
            # Check if any cell in the row contains one of the keywords
            if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                
                titles.append(index)
        
        # Return the list of detected titles after the loop finishes
        return titles

    keywords=['Nombre d\'abonnés']
    table_titles = detect_tables(df, keywords)

    #Extraction de la table nombre d'abonnés
    start = table_titles[0]+2
    end = table_titles[0]+15
    nombre_abbo2013 = df.iloc[start:end]
    nombre_abbo2013= nombre_abbo2013.dropna(axis='columns',how='all')
    nombre_abbo2013 = nombre_abbo2013.iloc[:,:49]
    nombre_abbo2013.index = nombre_abbo2013[0]
    del nombre_abbo2013[0]
    nombre_abbo2013 = nombre_abbo2013[1:]
    nombre_abbo2013.index.name = None
    nombre_abbo2013.columns = nombre_abbo2013.iloc[0]
    nombre_abbo2013 = nombre_abbo2013[1:]
    nombre_abbo2013 = nombre_abbo2013.fillna(0)
    nombre_abbo2013.columns = pd.MultiIndex.from_arrays([
    ['Janvier'] * 4 + ['Février'] * 4+['Mars'] * 4+['Avril'] * 4+['Mai'] * 4+['Juin'] * 4+['Juillet'] * 4+['Août'] * 4+ ['Septembre'] * 4+['Octobre'] * 4+['Novembre'] * 4+['Decembre'] * 4,  
    nombre_abbo2013.columns 
    ])
    nombre_abbo2013.columns.name = None
    


    #Extraction de la table accroissement
    keywords=['Nombre d\'abonnés','Accroissement']
    table_titles = detect_tables(df, keywords)
    start = table_titles[1]
    end = table_titles[1]+13
    accroissement2013 = df.iloc[start:end]
    accroissement2013 = accroissement2013.dropna(axis='columns',how='all')
    accroissement2013 = accroissement2013.iloc[:,:53]
    accroissement2013 = accroissement2013[1:]
    accroissement2013.index = np.arange(0,12)
    accroissement2013.index = accroissement2013[0]
    accroissement2013.index.name = None
    accroissement2013.columns = accroissement2013.iloc[0]
    accroissement2013 = accroissement2013[1:]
    accroissement2013.index.name = None
    accroissement2013.columns.values[0]='placeholder'
    del accroissement2013['placeholder']
    accroissement2013.columns.name = None
    accroissement2013 = accroissement2013.fillna(0)
    accroissement2013.columns = pd.MultiIndex.from_arrays([
    ['Janvier'] * 4 + ['Février'] * 4+['Mars'] * 4+['Avril'] * 4+['Mai'] * 4+['Juin'] * 4+['Juillet'] * 4+['Août'] * 4+ ['Septembre'] * 4+['Octobre'] * 4+['Novembre'] * 4+['Decembre'] * 4+['CUMUL'] * 4,  
    accroissement2013.columns 
    ])







    return nombre_abbo2013,accroissement2013 





def get_cumul_accroissement_gaz(wil=wilayas):

    nombre_abb,accroissement2013 = load_old_data_gaz()
    accroissement2013 = accroissement2013.loc[:, pd.IndexSlice['CUMUL', ['BP', 'MP']]]
    accroissement2013.columns = accroissement2013.columns.droplevel(0)
    # for key, filename in zip(wilayas.keys(), file):
    # # Call your function and unpack the returned values
    #     nombre_abonne, accroissement, apport, apport_nouv = load_process_xl(filename)
        
    #     # Append the returned values to the respective key
    #     wilayas[key].append({
    #         "nombre_abonne": nombre_abonne,
    #         "accroissement": accroissement,
    #         "apport": apport,
    #         "apport_nouv": apport_nouv
    #     })
    df1 = pd.DataFrame()
    for key in wil:
        accroissement2024 = wil[key][0]["accroissement"]
        line = accroissement2024.loc['Total',pd.IndexSlice['gaz', ['BP', 'MP']]]
        df = pd.DataFrame(line)

        # Drop the top level ('électricité') from the index
        df.index = df.index.droplevel(0)
        df.index = ['BP','MP']
        df.columns = [key]
        df = df.T
        # line = line.T
        # line.index.values[0]= key
        df1 = pd.concat([df1, df], ignore_index=True)
    sum_row = df1.sum()
    df1.loc['SDC'] = sum_row
    list = accroissement2013.index.values
    df1.index = [list]
  
    df1['Total']=df1.sum(axis=1)
    df1.columns = pd.MultiIndex.from_arrays([
        ['cumul-24'] *3, df1.columns
    ])
    accroissement2013['Total']=accroissement2013.sum(axis=1)
    accroissement2013.columns = pd.MultiIndex.from_arrays([
        ['cumul-23'] *3, accroissement2013.columns
    ])
    df1.index = df1.index.map(lambda x: x if isinstance(x, str) else x[0])
    accroissement2013.index = accroissement2013.index.map(lambda x: x if isinstance(x, str) else x[0])
    combined = pd.concat([accroissement2013, df1], axis=1)
   
    return combined

def dataset_region_client_23_accroissement_gaz():
    ##This function has the goal of extracting the necessary data from the nombre abonné 2023 data frame and returning them as a dataframe to be then concatinated to the 2024 data
    nombre_abbo2013, accroissement2013 = load_old_data_gaz()
    mounth = last_month_with_data(accroissement2013)
    test =  accroissement2013.loc[:,pd.IndexSlice[[mounth], ['BP', 'MP']]]
    
       
    
    dataframe13 = pd.DataFrame()
    dataframe13 = test
    dataframe13.columns = dataframe13.columns.droplevel(0)
    dataframe13['Total']=dataframe13.sum(axis=1)
    # wilayas = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    # dataframe13.index = wilayas
    return dataframe13

def dataset_region_client_24_accroissement_gaz(wil=wilayas):

    cringe=[]
    for key in wil:
        # Access the 'nombre_abonne' DataFrame for the current key
        nombre_abonne = wil[key][0]['accroissement']  # Make sure this points to the right DataFrame

        # Find the index of the repeated row
        index_of_repetition = find_repeated_row(nombre_abonne)
        
        # Check if the index_of_repetition is valid
        if index_of_repetition is not None and index_of_repetition in nombre_abonne.index:
            # Access the desired data
            result = nombre_abonne.loc[index_of_repetition, pd.IndexSlice['gaz', ['BP', 'MP']]]
            flipped_result = result.unstack(level=0)  # Flip the first level of the MultiIndex

            # Optionally, reset the index for a flat DataFrame
            flipped_result.reset_index(drop=True, inplace=True)
            flipped_result.columns.name = None

            # Append the flipped result to cringe
            cringe.append(flipped_result)
        else:
            print(f"No valid index found for key: {key}")

    dataframe24 = pd.concat(cringe)
    wilayass = ['Blida','Bouira','Médéa','T.Ouzou','Djelfa','Tipaza','Boumerdes','Ain Defla','Chlef','Tissemssilt']
    dataframe24.index = wilayass
    dataframe24['Total']= dataframe24.sum(axis=1)
    dataframe24.loc['SDC'] = dataframe24.sum()
    return dataframe24



def region_accroissement_gaz():
    
    dataframe13 = dataset_region_client_23_accroissement_gaz()
    dataframe24 = dataset_region_client_24_accroissement_gaz()
    region_nombre_abonne_dataframe = pd.concat([dataframe13, dataframe24], axis=1)
    region_nombre_abonne_dataframe.columns = pd.MultiIndex.from_arrays([
        ['23'] * 3 + ['24'] * 3, region_nombre_abonne_dataframe.columns
    ])
     # Calculate evolution (%) for BT and MT
      # Calculate and round evolution (%) for BT and MT

    region_nombre_abonne_dataframe[('evolution (%)', 'Tx. Evol')] = round(
        ((region_nombre_abonne_dataframe[('24', 'Total')] - region_nombre_abonne_dataframe[('23', 'Total')]) 
         / region_nombre_abonne_dataframe[('23', 'Total')]) * 100, 1
    )
    region_nombre_abonne_dataframe=  region_nombre_abonne_dataframe.astype(float)
    
    combined= get_cumul_accroissement_gaz()
    region_nombre_abonne_dataframe = pd.concat([region_nombre_abonne_dataframe, combined], axis=1)
    region_nombre_abonne_dataframe[('evolution (%) cumul', 'Tx. Evol')] = round(
        ((region_nombre_abonne_dataframe[('cumul-24', 'Total')] - region_nombre_abonne_dataframe[('cumul-23', 'Total')]) 
         / region_nombre_abonne_dataframe[('cumul-23', 'Total')]) * 100, 1
    )
    # total_row = region_nombre_abonne_dataframe.sum(numeric_only=True)
    # total_row.name = 'RDC'

    # Append the total row to the DataFrame
    # region_nombre_abonne_dataframe = pd.concat([region_nombre_abonne_dataframe, total_row.to_frame().T])
    return region_nombre_abonne_dataframe