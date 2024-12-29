import pandas as pd 
import numpy as np

def load_process_creances(file='Final Aout-TDB Créances Aout 2024-Final(2).xlsx'):
    df = pd.read_excel(file,sheet_name='Créances Globale')
    def detect_tables(df, keywords):
        titles = []
        for index, row in df.iterrows():
            
            # Check if any cell in the row contains one of the keywords
            if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                
                titles.append(index)
        
        # Return the list of detected titles after the loop finishes
        return titles
    keywords = ['III- 2- Soldes des créances au mois de']
    titles = detect_tables(df,keywords)
    start_index = titles[0]+2
    end_index = titles[0]+27
    sold = df.loc[start_index:end_index]
    sold.columns = np.arange(0,19)
    del sold[5]
    del sold[6]
    del sold[7]
    del sold[8]
    del sold[9]
    del sold[10]
    del sold[11]
    del sold[12]
    del sold[13]
    del sold[14]
    del sold[16]
    del sold[17]
    del sold[18]
    del sold[15]
    del sold[2]
    sold.iloc[0,0] = 'placeholder'
    sold.columns = sold.iloc[0]
    sold = sold[1:]
    sold = sold[1:]
    sold.index = sold['placeholder']
    del sold['placeholder']
    sold.index.name = None
    sold.columns.name = None
    sold.index = sold.index.str.strip()
    return sold


def recette_df(file='Final Aout-TDB Créances Aout 2024-Final(2).xlsx'):
    df = pd.read_excel(file,sheet_name='Créances Globale')
    def detect_tables(df, keywords):
        titles = []
        for index, row in df.iterrows():
            
            # Check if any cell in the row contains one of the keywords
            if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                
                titles.append(index)
        
        # Return the list of detected titles after the loop finishes
        return titles
    keywords = ['III- 4- Recettes']
    titles = detect_tables(df,keywords)
    start_index = titles[0]+1
    end_index = titles[0]+16
    recettes_df = df.iloc[start_index:end_index]
    recettes_df = recettes_df.iloc[1:]
    recettes_df.columns = recettes_df.iloc[0]
    recettes_df = recettes_df.iloc[1:]
    recettes_df= recettes_df.iloc[:,:10]
    recettes_df.index = recettes_df[2024]
    del recettes_df[2024]
    recettes_df.index.name = None
    recettes_df.columns.name = None
    recettes_df.columns.values[-1]='Total Energie'
    return recettes_df
    