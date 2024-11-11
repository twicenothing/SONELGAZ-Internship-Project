import pandas as pd
import numpy as np
from flask import Flask,send_file, jsonify
from flask import request, send_file, jsonify
from flask_cors import CORS
from funcs import *
from region_funcs import *
from data import *
from region_funcs_gaz import *
from region_funcs_accroissement import *
from region_funcs_gaz_accroissement import *
from region_funcs_apport import *
app = Flask(__name__)
CORS(app)

#Functions

# made some changes to this function it should work all fine tho 
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




def load_process_ventes(file='DD Blida - TDB ventes final hor tb.xlsx'):
    xl = pd.ExcelFile(file)
    df = xl.parse(0)
    df = df.dropna(how='all')
    cpt = 0
    for index,row in df.iterrows():
        cpt = cpt+1
    df.index = np.arange(0,cpt)

    def detect_tables(df, keywords):
        titles = []
        for index, row in df.iterrows():
            
            # Check if any cell in the row contains one of the keywords
            if row.astype(str).str.contains('|'.join(keywords), case=False).any():
                
                titles.append(index)
        
        # Return the list of detected titles after the loop finishes
        return titles
    keywords=['vente', 'achats','chiffre']
    titles = detect_tables(df,keywords)

    start_row = titles[0]+1
    end_row = titles[0]+15
    ventes = df.loc[start_row+1:end_row].reset_index(drop=True)
    ventes.columns = ventes.iloc[0]
    ventes= ventes[1:]
    ventes.columns.values[0]='placeholder'
    ventes.index = ventes['placeholder']
    ventes.index.name = None
    del ventes['placeholder']
    ventes.columns = pd.MultiIndex.from_arrays([
    ['électricité'] * 5 + ['gaz'] * 5 +['gaz Exprimé en M3'] * 5,  # Higher level (First_5, Last_5)
    ventes.columns  # Lower level (A, B, C, D, etc.)
    ])
    start_row = titles[1]+1
    end_row= titles[1]+15
    achatsdf = df.loc[start_row+1:end_row].reset_index(drop=True)
    achatsdf = achatsdf.dropna(axis='columns',how='all')
    achatsdf.columns = achatsdf.iloc[0]
    achatsdf = achatsdf[1:]
    achatsdf.columns.values[0]='placeholder'
    achatsdf.index = achatsdf['placeholder']
    del achatsdf['placeholder']
    achatsdf.index.name = None
    achatsdf.columns = pd.MultiIndex.from_arrays([
    ['électricité'] * 4 + ['gaz'] * 4 +['gaz Exprimé en M3'] * 4,  # Higher level (First_5, Last_5)
    achatsdf.columns  # Lower level (A, B, C, D, etc.)
    ])
    one = achatsdf.loc[:, pd.IndexSlice['électricité',['Ventes','Achats']]]
    ventes = one.loc[:, pd.IndexSlice['électricité', 'Ventes']].squeeze()
    achats = one.loc[:, pd.IndexSlice['électricité', 'Achats']].squeeze()


    tp1 = (achats - ventes).div(achats.where(achats != 0))
    one = achatsdf.loc[:, pd.IndexSlice['gaz',['Ventes','Achats']]]
    ventes = one.loc[:, pd.IndexSlice['gaz', 'Ventes']].squeeze()
    achats = one.loc[:, pd.IndexSlice['gaz', 'Achats']].squeeze()


    tp2 = (achats - ventes).div(achats.where(achats != 0))
    one = achatsdf.loc[:, pd.IndexSlice['gaz Exprimé en M3',['Ventes','Achats']]]
    ventes = one.loc[:, pd.IndexSlice['gaz Exprimé en M3', 'Ventes']].squeeze()
    achats = one.loc[:, pd.IndexSlice['gaz Exprimé en M3', 'Achats']].squeeze()


    tp3 = (achats - ventes).div(achats.where(achats != 0))
    tp = pd.concat([tp1, tp2, tp3], axis=1)
    tp.columns = ['électricité','gaz','gaz Exprimé en M3']
    
    return ventes,achats,tp

def get_wilaya(wilaya_code):
    mapping = {
        "09":"blida",
        "02":"chlef",
        "10":"bouira",
        "15":"tiziouzou",
        "17":"djelfa",
        "26":"medea",
        "35":"boumerdes",
        "38":"tissemsilt",
        "42":"tipaza",
        "44":"aindefla"
    }
    return mapping.get(wilaya_code,"invalid")




##Routes

@app.route('/upload-ventes/<wilaya_code>', methods=['POST'])
def upload_file_ventes(wilaya_code):
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']

        # Process the file
        ventes, achats, tp = load_process_ventes(file)
        wilaya = get_wilaya(wilaya_code)
       

       
       
        if wilaya in wilayas:
            
            wilayas_ventes[f"{wilaya}-ventes"].append(achats)
            wilayas_ventes[f"{wilaya}-ventes"].append(ventes)
            wilayas_ventes[f"{wilaya}-ventes"].append(tp)
            # For other types like `tp`, append where 
            # appropriate in the wilayas dictionary
            wilayaname = wilaya + '-ventes'
            print(wilayas_ventes[wilayaname])
        else:
            return jsonify({'error': f'Wilaya code {wilaya} not found'}), 404, print(f'{wilaya} not found')

        return jsonify({'message': f'Data successfully appended for {wilaya}'})


    except Exception as e:
        return jsonify({'error': str(e)}), 500



@app.route('/upload/<wilaya_code>', methods=['POST'])
def upload_file(wilaya_code):
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']

        # Process the file
        nombre_abonne, accroissement, apport,apport_nouv = load_process_xl(file)
        wilaya = get_wilaya(wilaya_code)

       
       
        if wilaya in wilayas:
            
            wilayas[wilaya].append({
            "nombre_abonne": nombre_abonne,
            "accroissement": accroissement,
            "apport": apport,
            "apport_nouv": apport_nouv
        })
            # For other types like `tp`, append where 
            # appropriate in the wilayas dictionary
            # print(wilayas[wilaya][0]['nombre_abonne'])
        
            
        else:
            return jsonify({'error': f'Wilaya code {wilaya} not found'}), 404

        return jsonify({'message': f'Data successfully appended for {wilaya}'})


    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/region_clientel')
def getolddataframe():
    flag = dictionary_is_full(wilayas)
    if flag == True:
        region_nombre_abonne_dataframe = region_nombre_abonne()
        print(region_nombre_abonne_dataframe)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})


@app.route('/reegion_clientel_gaz')
def region_clientel_gaz():
    flag = dictionary_is_full(wilayas)
    if flag == True:
        region_nombre_abonne_dataframe = region_nombre_abonne_gaz()
        print(region_nombre_abonne_dataframe)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})   


@app.route('/region_accroissement')
def region_accroissement_dt():
    flag = dictionary_is_full(wilayas)
    if flag == True:
        region_nombre_abonne_dataframe = region_accroissement()
        print(region_nombre_abonne_dataframe)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})   


@app.route('/region_accroissement_gaz')
def region_accroissement_dt_gaz():
    flag = dictionary_is_full(wilayas)
    if flag == True:
        region_nombre_abonne_dataframe_gaz = region_accroissement_gaz()
        print(region_nombre_abonne_dataframe_gaz)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})   



@app.route('/region_apport')
def region_apport_dt():
    flag = dictionary_is_full(wilayas)
    if flag == True:
        region_nombre_abonne_dataframe_gaz = region_apport()
        print(region_nombre_abonne_dataframe_gaz)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})   





if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
