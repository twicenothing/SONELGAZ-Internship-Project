import pandas as pd
import numpy as np
from data import *
from funcs import *


def load_old_data(file=TB):
    df = pd.read_excel(file,sheet_name='élec')
    df = df.dropna(how='all')
    first_row = df.columns.to_list()
    df.loc[0] = first_row  # Temporarily add the row with index -1
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
    table_titles
    #Extraction de la table nombre d'abonnés
    start = table_titles[0]+1
    end = table_titles[0]+14
    nombre_abbo2013 = df.iloc[start:end]
    nombre_abbo2013.columns = nombre_abbo2013.iloc[0]
    nombre_abbo2013 = nombre_abbo2013.iloc[1:]
    nombre_abbo2013.index = nombre_abbo2013['DD']
    del nombre_abbo2013['DD']
    nombre_abbo2013.index.name = None
    nombre_abbo2013.columns.naame = None
    nombre_abbo2013.columns = np.arange(0,58)
    nombre_abbo2013 = nombre_abbo2013.loc[:,:47]
    nombre_abbo2013.columns = nombre_abbo2013.iloc[0]
    nombre_abbo2013 = nombre_abbo2013[1:]
    nombre_abbo2013 = nombre_abbo2013.replace(0,np.nan)
    nombre_abbo2013 = nombre_abbo2013.dropna(axis=1, how='all' )
    french_months = [
        "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
    ]
    # Number of columns per month
    columns_per_month = 4

    # Calculate the number of months based on the columns
    num_months = (len(nombre_abbo2013.columns) + columns_per_month - 1) // columns_per_month

    # Generate month labels dynamically up to the current number of months
    months = french_months[:num_months]

    # Create the high-level (month) and low-level (original column names) indices
    high_level = []
    for month in months:
        high_level.extend([month] * columns_per_month)

    # Truncate high-level index to match the number of columns
    high_level = high_level[:len(nombre_abbo2013.columns)]

    # Create a MultiIndex
    multi_index = pd.MultiIndex.from_tuples(
        [(high, col) for high, col in zip(high_level, nombre_abbo2013.columns)]

    )

    # Set the MultiIndex as the new columns
    nombre_abbo2013.columns = multi_index
    nombre_abbo2013.columns.name = None
    nombre_abbo2013.index.name = None

    keywords=['Nombre d\'abonnés','Accroissement']
    table_titles = detect_tables(df, keywords)
    start = table_titles[1]+1
    end = table_titles[1]+14
    accroissement2013 = df.iloc[start:end]
    accroissement2013 = accroissement2013.dropna(axis='columns',how='all')
    accroissement2013 = accroissement2013.iloc[:,:53]
    accroissement2013 = accroissement2013.replace(0,np.nan)
    accroissement2013.columns = accroissement2013.iloc[1]
    accroissement2013 = accroissement2013.iloc[2:]
    accroissement2013 = accroissement2013.dropna(axis='columns',how='all')

    accroissement2013.columns.values[0]= 'placeholder'
    accroissement2013.index = accroissement2013['placeholder']
    del accroissement2013['placeholder']

    # Separate the last 4 columns for CUMUL
    num_regular_columns = len(accroissement2013.columns) - 4  # Exclude the last 4 for now
    num_months = (num_regular_columns + columns_per_month - 1) // columns_per_month

    # Generate month labels dynamically in French up to the current number of months
    months = french_months[:num_months]

    # Create the high-level index for the regular columns
    high_level = []
    for month in months:
        high_level.extend([month] * columns_per_month)

    # Add 'CUMUL' for the last 4 columns
    high_level.extend(['CUMUL'] * 4)

    # Ensure the high-level index matches the number of columns
    high_level = high_level[:len(accroissement2013.columns)]

    # Combine high-level index with existing column names (low-level index)
    multi_index = pd.MultiIndex.from_tuples(
        [(high, col) for high, col in zip(high_level, accroissement2013.columns)]

    )

    # Set the MultiIndex as the new columns
    accroissement2013.columns = multi_index
    accroissement2013.index.name = None

        #Extraction de la table apport
    keywords=['Nombre d\'abonnés','Accroissement','Apport']
    table_titles = detect_tables(df, keywords)
    start = table_titles[2]+1
    end = table_titles[2]+13
    apport2023 = df.iloc[start:end]
    apport2023 = apport2023.dropna(axis='columns',how='all')
    apport2023 = apport2023.iloc[:,:53]
    apport2023 = apport2023.replace(0,np.nan)
    apport2023.columns = apport2023.iloc[0]
    apport2023 = apport2023.iloc[1:]
    apport2023.columns.values[0] = 'placeholder'
    apport2023.index = apport2023['placeholder']
    del apport2023['placeholder']
    apport2023 = apport2023.dropna(axis=1,how='all')
    apport2023.index.name = None
    apport2023.columns.name = None

    # Separate the last 4 columns for CUMUL
    num_regular_columns = len(apport2023.columns) - 4  # Exclude the last 4 for now
    num_months = (num_regular_columns + columns_per_month - 1) // columns_per_month

    # Generate month labels dynamically in French up to the current number of months
    months = french_months[:num_months]

    # Create the high-level index for the regular columns
    high_level = []
    for month in months:
        high_level.extend([month] * columns_per_month)

    # Add 'CUMUL' for the last 4 columns
    high_level.extend(['CUMUL'] * 4)

    # Ensure the high-level index matches the number of columns
    high_level = high_level[:len(apport2023.columns)]

    # Combine high-level index with existing column names (low-level index)
    multi_index = pd.MultiIndex.from_tuples(
        [(high, col) for high, col in zip(high_level, apport2023.columns)]

    )

    # Set the MultiIndex as the new columns
    apport2023.columns = multi_index
    apport2023.index.name = None
    apport2023 = apport2023.replace(np.nan, 0)





    return nombre_abbo2013,accroissement2013,apport2023




def get_cumul_apport(wil=wilayas):

    nombre_abb,apport2023,apport2023 = load_old_data()
    apport2023 = apport2023.loc[:, pd.IndexSlice['CUMUL', ['BT', 'MT']]]
    apport2023.columns = apport2023.columns.droplevel(0)
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
        accroissement2024 = wil[key][0]["apport"]
        line = accroissement2024.loc['Total',pd.IndexSlice['électricité', ['BT', 'MT']]]
        df = pd.DataFrame(line)

        # Drop the top level ('électricité') from the index
        df.index = df.index.droplevel(0)
        df.index = ['BT','MT']
        df.columns = [key]
        df = df.T
        # line = line.T
        # line.index.values[0]= key
        df1 = pd.concat([df1, df], ignore_index=True)
    sum_row = df1.sum()
    df1.loc['SDC'] = sum_row
    list = apport2023.index.values
    df1.index = [list]
  
    df1['Total']=df1.sum(axis=1)
    df1.columns = pd.MultiIndex.from_arrays([
        ['cumul-24'] *3, df1.columns
    ])
    apport2023['Total']=apport2023.sum(axis=1)
    apport2023.columns = pd.MultiIndex.from_arrays([
        ['cumul-23'] *3, apport2023.columns
    ])
    df1.index = df1.index.map(lambda x: x if isinstance(x, str) else x[0])
    apport2023.index = apport2023.index.map(lambda x: x if isinstance(x, str) else x[0])
    combined = pd.concat([apport2023, df1], axis=1)
   
    return combined

def dataset_region_client_23_apport():
    ##This function has the goal of extracting the necessary data from the nombre abonné 2023 data frame and returning them as a dataframe to be then concatinated to the 2024 data
    nombre_abbo2013, accroissement2013,apport = load_old_data()
    mounth = last_month_with_data(apport)
    test =  apport.loc[:,pd.IndexSlice[[mounth], ['BT', 'MT']]]
    
       
    
    dataframe13 = pd.DataFrame()
    dataframe13 = test
    dataframe13.columns = dataframe13.columns.droplevel(0)
    dataframe13['Total']=dataframe13.sum(axis=1)
    # wilayas = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    # dataframe13.index = wilayas
    return dataframe13

def dataset_region_client_24_apport(wil=wilayas):

    cringe=[]
    for key in wil:
        # Access the 'nombre_abonne' DataFrame for the current key
        nombre_abonne = wil[key][0]['apport']  # Make sure this points to the right DataFrame

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
    wilayass = ['Blida','Bouira','Médéa','T.Ouzou','Djelfa','Tipaza','Boumerdes','Ain Defla','Chlef','Tissemssilt']
    dataframe24.index = wilayass
    dataframe24['Total']= dataframe24.sum(axis=1)
    dataframe24.loc['SDC'] = dataframe24.sum()
    return dataframe24


def region_apport():
    
    dataframe13 = dataset_region_client_23_apport()
    dataframe24 = dataset_region_client_24_apport()
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
    
    combined= get_cumul_apport()
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

