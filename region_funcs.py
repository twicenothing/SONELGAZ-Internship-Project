import pandas as pd
import numpy as np
from app import *
from data import *
from funcs import *
## function to load data from tableau de bord of 2023


def load_old_data(file=TB):
    df = pd.read_excel(file,sheet_name='élec')
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
    start = table_titles[1]+1
    end = table_titles[1]+14
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

    #Extraction de la table apport
    keywords=['Nombre d\'abonnés','Accroissement','Apport']
    table_titles = detect_tables(df, keywords)
    start = table_titles[2]+1
    end = table_titles[2]+13
    apport2023 = df.iloc[start:end]
    apport2023 = apport2023.dropna(axis='columns',how='all')
    apport2023 = apport2023.iloc[:,:53]
    apport2023.index = np.arange(0,12)
    apport2023.index = apport2023[0]
    apport2023.index.name = None
    apport2023.columns = apport2023.iloc[0]
    apport2023 =apport2023[1:]
    apport2023.columns.values[0]='placeholder'
    del apport2023['placeholder']
    apport2023.index.name = None
    apport2023.columns.name = None
    apport2023.columns = pd.MultiIndex.from_arrays([
    ['Janvier'] * 4 + ['Février'] * 4+['Mars'] * 4+['Avril'] * 4+['Mai'] * 4+['Juin'] * 4+['Juillet'] * 4+['Août'] * 4+ ['Septembre'] * 4+['Octobre'] * 4+['Novembre'] * 4+['Decembre'] * 4+['CUMUL'] * 4,  
    apport2023.columns 
    ])
    apport2023 = apport2023.fillna(0)





    return nombre_abbo2013,accroissement2013,apport2023



# this returns a dataframe for the latest mounth data for 2023 
def dataset_region_client_23():
    ##This function has the goal of extracting the necessary data from the nombre abonné 2023 data frame and returning them as a dataframe to be then concatinated to the 2024 data
    nombre_abbo2013, accroissement2013,_ = load_old_data()
    mounth = last_month_with_data(nombre_abbo2013)
    test =[]
    for i in range(len(nombre_abbo2013)-1):
        test.append(nombre_abbo2013.loc[nombre_abbo2013.index.values[i],pd.IndexSlice[mounth, ['BT', 'MT']]].unstack())
    i=0
    dataframe13 = pd.DataFrame()
    dataframe13 = pd.concat(test)
    wilayas = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    dataframe13.index = wilayas
    return dataframe13


def dataset_region_client_23_accroissement():
    ##This function has the goal of extracting the necessary data from the nombre abonné 2023 data frame and returning them as a dataframe to be then concatinated to the 2024 data
    nombre_abbo2013, accroissement2013,_ = load_old_data()
    mounth = last_month_with_data(accroissement2013)
    test =[]
    for i in range(len(accroissement2013)-1):
        test.append(accroissement2013.loc[accroissement2013.index.values[i],pd.IndexSlice['Août', ['BT', 'MT']]].unstack())
    i=0
    dataframe13 = pd.DataFrame()
    dataframe13 = pd.concat(test)
    wilayas = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    dataframe13.index = wilayas
    return dataframe13
















#this returns the data i need for 2024 to then concatinate it with 23
def dataset_region_client_24(wilaya=wilayas):

    
    cringe=[]
    for key in wilaya:
        # Access the 'nombre_abonne' DataFrame for the current key
        nombre_abonne = wilaya[key][0]['nombre_abonne']  # Make sure this points to the right DataFrame

        # Find the index of the repeated row
        index_of_repetition = find_repeated_row(nombre_abonne)
        
        # Check if the index_of_repetition is valid
        if index_of_repetition is not None and index_of_repetition in nombre_abonne.index:
            # Access the desired data
            result = nombre_abonne.loc[index_of_repetition, pd.IndexSlice['électricité', ['BT', 'MT']]]
            flipped_result = result.unstack(level=0)  # Flip the first level of the MultiIndex

            # Optionally, reset the index for a flat DataFrame
            flipped_result.reset_index(drop=True, inplace=True)
            flipped_result.columns.name = None

            # Append the flipped result to cringe
            cringe.append(flipped_result)
        else:
            print(f"No valid index found for key: {key}")

    dataframe24 = pd.concat(cringe)
    wilayass = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    dataframe24.index = wilayass
    return dataframe24


def region_nombre_abonne():
    
    dataframe13 = dataset_region_client_23()
    dataframe24 = dataset_region_client_24()
    region_nombre_abonne_dataframe = pd.concat([dataframe13, dataframe24], axis='columns')
    region_nombre_abonne_dataframe.columns = pd.MultiIndex.from_arrays([
        ['23'] * 2 + ['24'] * 2, region_nombre_abonne_dataframe.columns
    ])
     # Calculate evolution (%) for BT and MT
      # Calculate and round evolution (%) for BT and MT
    region_nombre_abonne_dataframe[('evolution (%)', 'BT')] = round(
        ((region_nombre_abonne_dataframe[('24', 'BT')] - region_nombre_abonne_dataframe[('23', 'BT')]) 
         / region_nombre_abonne_dataframe[('23', 'BT')]) * 100, 1
    )

    region_nombre_abonne_dataframe[('evolution (%)', 'MT')] = round(
        ((region_nombre_abonne_dataframe[('24', 'MT')] - region_nombre_abonne_dataframe[('23', 'MT')]) 
         / region_nombre_abonne_dataframe[('23', 'MT')]) * 100, 1
    )
    region_nombre_abonne_dataframe=  region_nombre_abonne_dataframe.astype(float)
    total_row = region_nombre_abonne_dataframe.sum(numeric_only=True)
    total_row.name = 'RDC'

    # Append the total row to the DataFrame
    region_nombre_abonne_dataframe = pd.concat([region_nombre_abonne_dataframe, total_row.to_frame().T])
    return region_nombre_abonne_dataframe




