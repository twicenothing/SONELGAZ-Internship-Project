import pandas as pd
import numpy as np
from flask import Flask,send_file, jsonify

app = Flask(__name__)

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

if __name__ == '__main__':
    app.run(debug=True)