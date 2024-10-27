import pandas as pd
import numpy as np
from flask import Flask,send_file, jsonify
from flask import request, send_file, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)













def load_process_xl(file='clientèle.xlsx'):
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
        return titles
    
    keywords = ['Nombre d\'abonnés', 'résiliations', 'réabonnés']
    table_titles = detect_tables(df, keywords)
    
    # SETTING UP NOMBRE D'ABONNEES DATA FRAME
    start_row = table_titles[0] + 1
    end_row = table_titles[0] + 16
    nombre_abonne = df.iloc[start_row + 1:end_row].reset_index(drop=True)
    nombre_abonne.columns = nombre_abonne.iloc[0]
    nombre_abonne = nombre_abonne[1:]
    nombre_abonne = nombre_abonne.iloc[:,:11]
    nombre_abonne.columns.values[0] = 'placeholder'
    nombre_abonne.index = nombre_abonne['placeholder']
    del nombre_abonne['placeholder']
    nombre_abonne.index.name = None
    nombre_abonne.index.values[0] = 'déc-23'
    nombre_abonne.columns = pd.MultiIndex.from_arrays([
        ['élecricité'] * 5 + ['gaz'] * 5, nombre_abonne.columns
    ])
    nombre_abonne = nombre_abonne.astype(int)
    
    # SETTING UP ACCROISSEMENT
    accroissement = nombre_abonne.diff().fillna(0)[1:]
    accroissement.loc[len(accroissement)] = accroissement.sum()
    accroissement.index.values[-1] = 'Total'

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
    resiliation = resiliation.iloc[:,:10]
    resiliation.columns = pd.MultiIndex.from_arrays([
        ['élecricité'] * 5 + ['gaz'] * 5, resiliation.columns
    ])
    resiliation = resiliation.astype(int)
    resiliation = resiliation.fillna(0)

    apport = accroissement + resiliation
    
    #SEETTING UP APPORT NOUVEAU
    start_row = table_titles[3] + 1
    end_row = table_titles[3] + 16
    reabonne = df.iloc[start_row + 1:end_row].reset_index(drop=True)
    reabonne.columns = reabonne.iloc[0]
    reabonne = reabonne[1:]
    reabonne.columns.values[0] = 'placeholder'
    reabonne.index = reabonne['placeholder']
    del reabonne['placeholder']
    reabonne.index.name = None
    reabonne = reabonne.iloc[:,:8]
    reabonne.columns = pd.MultiIndex.from_arrays([
        ['élecricité'] * 4 + ['gaz'] * 4, reabonne.columns
    ])
    reabonne = reabonne.astype(int)
    reabonne = reabonne.fillna(0)

    apport_nouv = apport - reabonne





    return nombre_abonne, accroissement, apport,apport_nouv


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
    ['élecricité'] * 5 + ['gaz'] * 5 +['gaz Exprimé en M3'] * 5,  # Higher level (First_5, Last_5)
    ventes.columns  # Lower level (A, B, C, D, etc.)
    ])
    start_row = titles[1]+1
    end_row= titles[1]+15
    achats = df.loc[start_row+1:end_row].reset_index(drop=True)
    achats = achats.dropna(axis='columns',how='all')
    achats.columns = achats.iloc[0]
    achats = achats[1:]
    achats.columns.values[0]='placeholder'
    achats.index = achats['placeholder']
    del achats['placeholder']
    achats.index.name = None
    achats.columns = pd.MultiIndex.from_arrays([
    ['élecricité'] * 4 + ['gaz'] * 4 +['gaz Exprimé en M3'] * 4,  # Higher level (First_5, Last_5)
    achats.columns  # Lower level (A, B, C, D, etc.)
    ])
    tp = (achats - ventes) / achats.where(achats != 0)
    tp = tp.fillna(0)
    
    return ventes,achats,tp


@app.route('/download-dataframesventes', methods=['GET'])
def download_dataframesventes():
    try:
        ventes, achats, tp = load_process_ventes()
        
        output_file = 'ventes.xlsx'
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            ventes.to_excel(writer, sheet_name='Ventes', index=True)
            achats.to_excel(writer, sheet_name='Achats', index=True)
            tp.to_excel(writer, sheet_name='Tp', index=True)
            

        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500




@app.route('/upload-ventes', methods=['POST'])
def upload_file_ventes():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']

        # Process the file
        ventes, achats, tp = load_process_xl(file)

        # Write processed data to Excel file
        output_file = 'processed_data_ventes.xlsx'
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            ventes.to_excel(writer, sheet_name='Ventes', index=True)
            achats.to_excel(writer, sheet_name='Achats', index=True)
            tp.to_excel(writer, sheet_name='tp', index=True)
            

        # Send the processed Excel file as a response
        return send_file(output_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='processed_data_ventes.xlsx')

    except Exception as e:
        return jsonify({'error': str(e)}), 500

























@app.route('/download-dataframes', methods=['GET'])
def download_dataframes():
    try:
        nombre_abonne, accroissement, apport,apport_nouv = load_process_xl()
        
        output_file = 'nombre_abonnee.xlsx'
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            nombre_abonne.to_excel(writer, sheet_name='Nombre Abonnés', index=True)
            accroissement.to_excel(writer, sheet_name='Accroissement', index=True)
            apport.to_excel(writer, sheet_name='Apport', index=True)
            apport_nouv.to_excel(writer, sheet_name='Apport_nouveau', index=True)

        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500




@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']

        # Process the file
        nombre_abonne, accroissement, apport, apport_nouv = load_process_xl(file)

        # Write processed data to Excel file
        output_file = 'processed_data.xlsx'
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            nombre_abonne.to_excel(writer, sheet_name='Nombre Abonnés', index=True)
            accroissement.to_excel(writer, sheet_name='Accroissement', index=True)
            apport.to_excel(writer, sheet_name='Apport', index=True)
            apport_nouv.to_excel(writer, sheet_name='Apport_nouveau', index=True)

        # Send the processed Excel file as a response
        return send_file(output_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='processed_data.xlsx')

    except Exception as e:
        return jsonify({'error': str(e)}), 500



    

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
