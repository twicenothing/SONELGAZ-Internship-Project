import pandas as pd
import numpy as np
from data import wilayas
from region_funcs import *
from funcs import *

def get_cumul_accroissement(wil=wilayas):
    nombre_abb,accroissement2013,_ = load_old_data()
    accroissement2013 = accroissement2013.loc[:, pd.IndexSlice['CUMUL', ['BT', 'MT']]]
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

def dataset_region_client_23_accroissement():
    ##This function has the goal of extracting the necessary data from the nombre abonné 2023 data frame and returning them as a dataframe to be then concatinated to the 2024 data
    nombre_abbo2013, accroissement2013,_ = load_old_data()
    mounth = last_month_with_data(accroissement2013)
    test =  accroissement2013.loc[:,pd.IndexSlice[[mounth], ['BT', 'MT']]]
    
       
    
    dataframe13 = pd.DataFrame()
    dataframe13 = test
    dataframe13.columns = dataframe13.columns.droplevel(0)
    dataframe13['Total']=dataframe13.sum(axis=1)
    # wilayas = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    # dataframe13.index = wilayas
    return dataframe13

def dataset_region_client_24_accroissement(wil=wilayas):

    cringe=[]
    for key in wil:
        # Access the 'nombre_abonne' DataFrame for the current key
        nombre_abonne = wil[key][0]['accroissement']  # Make sure this points to the right DataFrame

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

def region_accroissement():
    
    dataframe13 = dataset_region_client_23_accroissement()
    dataframe24 = dataset_region_client_24_accroissement()
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
    
    combined= get_cumul_accroissement()
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