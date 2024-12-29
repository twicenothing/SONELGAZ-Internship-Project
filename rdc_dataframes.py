import pandas as pd 
import numpy as np



def branchement_dataframe_elec(file='RCN8.xlsx'):
    df = pd.read_excel(file,sheet_name='RCN')
    df = df.dropna(how='all')
    def detect_tables(df, keywords):
            titles = []
            for index, row in df.iterrows():
                
                # Check if any cell in the row contains one of the keywords
                if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                    
                    titles.append(index)
            
            # Return the list of detected titles after the loop finishes
            return titles
    keywords=['Branchement élec']
    table_titles = detect_tables(df, keywords)
    start_index = table_titles[0]+1
    end_index = table_titles[0]+16
    branchement_elec = df.loc[start_index:end_index]
    branchement_elec.columns = branchement_elec.iloc[0]
    branchement_elec = branchement_elec.iloc[1:]
    branchement_elec.index = branchement_elec['Bracht élec']
    del branchement_elec['Bracht élec']
    branchement_elec.index.name = None
    branchement_elec.columns.name = None
    branchement_elec.columns.values[-1] = 'placeholder'
    del branchement_elec['placeholder']

    return branchement_elec




def branchement_dataframe_gaz(file = 'RCN8.xlsx'):
    df = pd.read_excel(file,sheet_name='RCN')
    df = df.dropna(how='all')
    def detect_tables(df, keywords):
            titles = []
            for index, row in df.iterrows():
                
                # Check if any cell in the row contains one of the keywords
                if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                    
                    titles.append(index)
            
            # Return the list of detected titles after the loop finishes
            return titles
    keywords=['Branchement gaz']
    table_titles = detect_tables(df, keywords)
    start_index = table_titles[0]+1
    end_index = table_titles[0]+16
    branchement_gaz = df.loc[start_index:end_index]
    branchement_gaz.columns = branchement_gaz.iloc[0]
    branchement_gaz = branchement_gaz.iloc[1:]
    branchement_gaz.index = branchement_gaz['Bracht Gaz']
    del branchement_gaz['Bracht Gaz']
    branchement_gaz.index.name = None
    branchement_gaz.columns.name = None
    branchement_gaz.columns.values[-1] = 'placeholder'
    del branchement_gaz['placeholder']

    return branchement_gaz


def extension_dataframe_elec(file = 'RCN8.xlsx'):
    df = pd.read_excel(file,sheet_name='RCN')
    df = df.dropna(how='all')
    def detect_tables(df, keywords):
            titles = []
            for index, row in df.iterrows():
                
                # Check if any cell in the row contains one of the keywords
                if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                    
                    titles.append(index)
            
            # Return the list of detected titles after the loop finishes
            return titles
    keywords=['Extension élec']
    table_titles = detect_tables(df, keywords)
    start_index = table_titles[0]+1
    end_index = table_titles[0]+16
    extension_elec = df.loc[start_index:end_index]
    extension_elec.columns = extension_elec.iloc[0]
    extension_elec = extension_elec.iloc[1:]
    extension_elec.index = extension_elec['Extension électricité']
    del extension_elec['Extension électricité']
    extension_elec.index.name = None
    extension_elec.columns.name = None
    extension_elec.columns.values[-1] = 'placeholder'
    del extension_elec['placeholder']

    return extension_elec




def extension_dataframe_gaz(file = 'RCN8.xlsx'):
    df = pd.read_excel(file,sheet_name='RCN')
    df = df.dropna(how='all')
    def detect_tables(df, keywords):
            titles = []
            for index, row in df.iterrows():
                
                
                if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                    
                    titles.append(index)
            
            
            return titles
    keywords=['Extension Gaz']
    table_titles = detect_tables(df, keywords)
    start_index = table_titles[0]+1
    end_index = table_titles[0]+16
    extension_gaz = df.loc[start_index:end_index]
    extension_gaz.columns = extension_gaz.iloc[0]
    extension_gaz = extension_gaz.iloc[1:]
    extension_gaz.index = extension_gaz['Extension Gaz']
    del extension_gaz['Extension Gaz']
    extension_gaz.index.name = None
    extension_gaz.columns.name = None
    extension_gaz.columns.values[-1] = 'placeholder'
    del extension_gaz['placeholder']

    return extension_gaz