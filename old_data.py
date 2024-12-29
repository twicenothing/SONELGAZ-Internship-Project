import pandas as pd
import numpy as np
from data import *


def load_solde_23(file=TB):
    df = pd.read_excel(file,sheet_name='calcul-DCC-TX-Encais old')
    def detect_tables(df, keywords):
        titles = []
        for index, row in df.iterrows():
                
                # Check if any cell in the row contains one of the keywords
            if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                    
                titles.append(index)
            
            # Return the list of detected titles after the loop finishes
        return titles
    keywords = ['Solde_23']
    titles = detect_tables(df,keywords)
    start_index = titles[0]+3
    end_index = titles[0]+15
    solde23 = df.loc[start_index:end_index]
    solde23.columns = solde23.iloc[0]
    solde23 = solde23[1:]
    solde23.columns.values[0]='placeholder'
    del solde23['placeholder']
    solde23.columns.values[0]='placeholder'
    solde23.index= solde23['placeholder']
    del solde23['placeholder']
    solde23.index.name = None 
    solde23.columns.name= None
    solde23 = solde23.iloc[:,:31]
    solde23.columns.values[-1]= 'Total Créances'
    solde23.columns.values[8] = 'Total Energie'
    solde23.columns.values[12]='Total Créances avec eaux'
    solde23.columns.values[26]='Total Energie sans eaux'
    solde23 = solde23.iloc[1:]
    return solde23