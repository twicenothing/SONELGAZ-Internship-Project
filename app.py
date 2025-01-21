import pandas as pd
import numpy as np
import os
from openpyxl.utils.dataframe import dataframe_to_rows
from flask import Flask,send_file, jsonify
from flask import request, send_file, jsonify
from flask_cors import CORS
from funcs import *
from data import *
from creances.creance import load_process_creances, recette_df
from RCN.rdc_dataframes import *
from RCN.rdc_TB import branchements_simples_elecrticite,branchements_simples_gaz,extension_portfeuille_elec,extension_portfeuille_gaz,extension_affaires_elec,branchment_affaires_elec,extension_affaires_gaz, branchment_affaires_gaz
from TB_creation import TB_clientele, TB_ventes, TB_ventes_gaz, TB_RCN_elec, TB_RCN_gaz, TB_solde, TB_elec
from elec.elec_tb import nombre_abo_tb, accroissement_tb, apport_tb, apport_nv_tb, ventes_tb, achats_tb, resiliation_tb, reabonne_tb, chiffre_aff_tb,prix_tb
pd.set_option('display.float_format', '{:.2f}'.format)





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

    return nombre_abonne, accroissement, apport, apport_nouv,resiliation, reabonne



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
    keywords=['vente', 'achats','chiffre','II- 3- Chiffres d\'affaires HT','II- 4- Prix moyens']
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
    tventes = one.loc[:, pd.IndexSlice['électricité', 'Ventes']].squeeze()
    tachats = one.loc[:, pd.IndexSlice['électricité', 'Achats']].squeeze()


    tp1 = (tachats - tventes).div(tachats.where(tachats != 0))
    one = achatsdf.loc[:, pd.IndexSlice['gaz',['Ventes','Achats']]]
    tventes = one.loc[:, pd.IndexSlice['gaz', 'Ventes']].squeeze()
    tachats = one.loc[:, pd.IndexSlice['gaz', 'Achats']].squeeze()


    tp2 = (tachats - tventes).div(tachats.where(tachats != 0))
    one = achatsdf.loc[:, pd.IndexSlice['gaz Exprimé en M3',['Ventes','Achats']]]
    tventes = one.loc[:, pd.IndexSlice['gaz Exprimé en M3', 'Ventes']].squeeze()
    tachats = one.loc[:, pd.IndexSlice['gaz Exprimé en M3', 'Achats']].squeeze()


    tp3 = (tachats - tventes).div(tachats.where(tachats != 0))
    tp = pd.concat([tp1, tp2, tp3], axis=1)
    tp.columns = ['électricité','gaz','gaz Exprimé en M3']



    start_index = titles[3]+1
    end_index = titles[3]+15
    chiffre_aff = df.loc[start_index:end_index]
    chiffre_aff = chiffre_aff.dropna(axis=1, how='all')
    chiffre_aff = chiffre_aff.iloc[1:]
    chiffre_aff.columns = chiffre_aff.iloc[0]
    chiffre_aff = chiffre_aff.iloc[1:]
    chiffre_aff.columns.values[0] = 'placeholder'
    chiffre_aff.index = chiffre_aff['placeholder']
    del chiffre_aff['placeholder']
    chiffre_aff.index.name = None
    chiffre_aff.columns.name = None
    chiffre_aff.columns = pd.MultiIndex.from_arrays([
    ['électricité'] * 5 + ['gaz'] * 5,  # Higher level (First_5, Last_5)
    chiffre_aff.columns  # Lower level (A, B, C, D, etc.)
    ])
    




    start_index = titles[4]+1
    end_index = titles[4]+15
    prix = df.loc[start_index:end_index]
    prix = prix.dropna(axis=1, how='all')
    prix = prix.iloc[1:]
    prix.columns = prix.iloc[0]
    prix = prix.iloc[1:]
    prix.columns.values[0] = 'placeholder'
    prix.index = prix['placeholder']
    del prix['placeholder']
    prix.index.name = None
    prix.columns.name = None
    prix.columns = pd.MultiIndex.from_arrays([
    ['électricité'] * 5 + ['gaz'] * 5,  # Higher level (First_5, Last_5)
    prix.columns  # Lower level (A, B, C, D, etc.)
    ])
    return ventes,achatsdf,tp, chiffre_aff,prix

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


wilayas_rnc_file = ['RCN8.xlsx','RCN8.xlsx','RCN8.xlsx','RCN8.xlsx','RCN8.xlsx','RCN8.xlsx','RCN8.xlsx','RCN8.xlsx','RCN8.xlsx','RCN8.xlsx']

def fill_dict():
    for key, filename in zip(wilayas_rnc.keys(), wilayas_rnc_file):
# Call your function and unpack the returned values
        branchement = branchement_dataframe_elec(filename)
        extension = extension_dataframe_elec(filename)
        
        # Append the returned values to the respective key
        wilayas_rnc[key].append({
            "branchement": branchement,
            "extension": extension,  
        })

def fill_dict_gaz():
    for key, filename in zip(wilayas_rnc.keys(), wilayas_rnc_file):
# Call your function and unpack the returned values
        branchement = branchement_dataframe_gaz(filename)
        extension = extension_dataframe_gaz()
        
        
        # Append the returned values to the respective key
        wilayas_rnc_gaz[key].append({
            "branchement": branchement,
            "extension": extension,
            
            
        })


@app.route('/fill')
def filldictionary():
    fill_dict()
        
    return jsonify({'message':'filled dict',}),200

    

@app.route('/fillgaz')
def filldictionarygaz():
    try :
        fill_dict_gaz()
        
        return jsonify({'message':'filled dict',}),200
    except Exception as e:
        return jsonify({'message':str(e)}),500





    

@app.route('/tableauteshting')
def tableauteshting():
    result = branchment_affaires_gaz()
    print(result)
    return jsonify({'message':'done!',}),200



















##Routes

@app.route('/upload-ventes/<wilaya_code>', methods=['POST'])
def upload_file_ventes(wilaya_code):
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400
       
        file = request.files['file']
      
        # Process the file
        ventes, achats, tp, chiffre_aff,prix = load_process_ventes(file)
     
        wilaya = get_wilaya(wilaya_code)
        if wilaya in wilayas:
            
            wilayas_ventes[wilaya].append({
            "ventes": ventes,
            "achaats": achats,
            "tp": tp,
            'chiffre_aff': chiffre_aff,
            'prix': prix,
        })
            


            # For other types like `tp`, append where 
            # appropriate in the wilayas dictionary
            # print(wilayas[wilaya][0]['nombre_abonne'])
            print(wilayas_ventes[wilaya][0]['chiffre_aff'])
        else:
        
            return jsonify({'error': f'Wilaya code {wilaya} not found'}), 404, print(f'{wilaya} not found')

        return jsonify({'message': f'Data successfully appended for {wilaya}'})


    except Exception as e:
        print(e)
        return jsonify({'error': str(e)}), 500



@app.route('/upload/<wilaya_code>', methods=['POST'])
def upload_file(wilaya_code):
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']

        # Process the file
        nombre_abonne, accroissement, apport,apport_nouv,resiliation,reabonne = load_process_xl(file)
        wilaya = get_wilaya(wilaya_code)

        if wilaya in wilayas:
            
            wilayas[wilaya].append({
            "nombre_abonne": nombre_abonne,
            "accroissement": accroissement,
            "apport": apport,
            "apport_nouv": apport_nouv,
            "resiliation": resiliation,
            "reabonne": reabonne,
        })
            print(wilayas[wilaya][0]['reabonne'])
            
        else:
            return jsonify({'error': f'Wilaya code {wilaya} not found'}), 404

        return jsonify({'message': f'Data successfully appended for {wilaya}'})


    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/upload-solde/<wilaya_code>', methods=['POST'])
def upload_file_solde(wilaya_code):
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']

        # Process the file
        solde = load_process_creances(file)
        recette = recette_df(file)
        wilaya = get_wilaya(wilaya_code)

        if wilaya in wilayas:
            
            wilayas_creance[wilaya].append({
            "solde": solde,
            "recette": recette,
        })
            # For other types like `tp`, append where 
            # appropriate in the wilayas dictionary
            # print(wilayas_creance[wilaya][0]['solde'])
            # print('\n')
            # print(wilayas_creance[wilaya][0]['recette'])
        
            
        else:
            return jsonify({'error': f'Wilaya code {wilaya} not found'}), 404

        return jsonify({'message': f'Data successfully appended for {wilaya}'})


    except Exception as e:
        return jsonify({'error': str(e)}), 500







@app.route('/tableau_de_borde', methods=['POST'])
def upload_file_TB():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']
        directory = os.path.join(os.getcwd(), 'tableau_de_bord')
        os.makedirs(directory, exist_ok=True)  # Create the folder if it doesn't exist

        # Define the file path
        file_path = os.path.join(directory, 'TB.xlsx')

        # Save the file to disk
        file.save(file_path)

        return jsonify({'success': True})


    except Exception as e:
        return jsonify({'error': str(e)}), 500





@app.route('/date')
def datee():
    print(total_days)
    return jsonify({'message':'done!'})



@app.route('/get_tb', methods=['GET'])
def get_tableau_de_bord():
    try:
        TB_clientele()
        TB_ventes()
        TB_ventes_gaz()
        TB_RCN_elec()
        TB_RCN_gaz()
        TB_solde()
        TB_elec()
        
        # Return success response immediately
        return jsonify({"status": "success", "message": "Tableau de bord generated successfully"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500



@app.route('/upload-rcn/<wilaya_code>', methods=['POST'])
def upload_file_rcn(wilaya_code):
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request'}), 400

        file = request.files['file']
        wilaya = get_wilaya(wilaya_code)

        if wilaya in wilayas:
            try:
                # Process the file for electricity
                branchement = branchement_dataframe_elec(file)
                # Need to seek back to start of file as it was read in previous operation
                file.seek(0)
                extension = extension_dataframe_elec(file)
                
                wilayas_rnc[wilaya].append({
                    "branchement": branchement,
                    "extension": extension
                })

                # Reset file pointer again
                file.seek(0)
                
                # Process the file for gas
                branchement_gaz = branchement_dataframe_gaz(file)
                file.seek(0)
                extension_gaz = extension_dataframe_gaz()  # Note: This doesn't seem to need a file parameter
                
                wilayas_rnc_gaz[wilaya].append({
                    "branchement": branchement_gaz,
                    "extension": extension_gaz
                })
                
                return jsonify({'message': f'Data successfully appended for {wilaya}'})
            except Exception as processing_error:
                print(f"Error processing file: {str(processing_error)}")  # For debugging
                return jsonify({'error': f'Error processing file: {str(processing_error)}'}), 500
        else:
            return jsonify({'error': f'Wilaya code {wilaya} not found'}), 404

    except Exception as e:
        print(f"General error: {str(e)}")  # For debugging
        return jsonify({'error': str(e)}), 500


@app.route('/check-wilaya/<wilaya_code>')
def check_wilaya(wilaya_code):
    try:
        wilaya = get_wilaya(wilaya_code)
        if wilaya == "invalid":
            return jsonify({'error': 'Invalid wilaya code'}), 400
            
        status = check_wilaya_status(wilaya)
        
        return jsonify({
            'wilaya': wilaya,
            'status': status,
            'is_complete': all(status.values())
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/check-all-wilayas')
def check_all_wilayas():
    try:
        all_wilayas_status = {}
        
        for wilaya in wilayas.keys():
            status = check_wilaya_status(wilaya)
            all_wilayas_status[wilaya] = {
                'status': status,
                'is_complete': all(status.values())
            }
        
        return jsonify({
            'wilayas': all_wilayas_status,
            'all_complete': all(wilaya_data['is_complete'] for wilaya_data in all_wilayas_status.values())
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/nombre_abo_tb')
def nombre_abo_tb_test():
    result = prix_tb()
    print(result)
    return jsonify({'message':'done!'}),200




if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')