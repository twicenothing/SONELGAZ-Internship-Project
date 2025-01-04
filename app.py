import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from flask import Flask,send_file, jsonify
from flask import request, send_file, jsonify
from openpyxl import load_workbook
from flask_cors import CORS
from funcs import *
from region_funcs import *
from data import *
from region_funcs_gaz import *
from region_funcs_accroissement import *
from region_funcs_gaz_accroissement import *
from region_funcs_apport import *
from region_funcs_gaz_apport import *
from creances.creance import load_process_creances, recette_df
from creances.solde_df import solde_dataframe
from creances.prise_en_charge_df import prise_en_charge_dataframe
from creances.encaissement_df import encaissement_dataframe
from creances.solde23_df import load_old_data_creance
from creances.recettes_df import recettes_tb
from creances.ddc_df import ddc_tb
from creances.old_data import load_solde_23
from ventes.region_ventes import dataset_region_ventes_24, dataset_region_ventes_23, region_ventes,get_cumul_ventes
from ventes.region_ventes_gaz import dataset_region_ventes_24_gaz, dataset_region_ventes_23_gaz,region_ventes_gaz,get_cumul_ventes_gaz
from RCN.rdc_dataframes import *
from RCN.rdc_TB import branchements_simples_elecrticite,branchements_simples_gaz,extension_portfeuille_elec,extension_portfeuille_gaz,extension_affaires_elec,branchment_affaires_elec,extension_affaires_gaz, branchment_affaires_gaz


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
    
    return ventes,achatsdf,tp

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
        ventes, achats, tp = load_process_ventes(file)
     
        wilaya = get_wilaya(wilaya_code)
       
       
       
        if wilaya in wilayas:
            
            wilayas_ventes[wilaya].append({
            "ventes": ventes,
            "achaats": achats,
            "tp": tp,
           
        })
            


            # For other types like `tp`, append where 
            # appropriate in the wilayas dictionary
            # print(wilayas[wilaya][0]['nombre_abonne'])
            wilayaname = wilaya + '-ventes'
            # print(wilayas_ventes[wilayaname][0]['ventes'])
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
            print(wilayas[wilaya][0]['nombre_abonne'])
        
            
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
















@app.route('/region_clientel')
def getolddataframe():
    flag = dictionary_is_full(wilayas)
    if flag == True:
        region_nombre_abonne_dataframe = region_nombre_abonne()
        print(region_nombre_abonne_dataframe)
       # Create a workbook and sheet
        wb = Workbook()
        ws = wb.active

        # Add DataFrame rows to the sheet
        for r_idx, row in enumerate(dataframe_to_rows(region_nombre_abonne_dataframe, index=False, header=True), start=1):
            ws.append(row)

        # Define a simple border style
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        # Apply borders to all cells
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border

        # Save the workbook
        wb.save("simple_table_with_outlines.xlsx")
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

@app.route('/region_apport_gaz')
def region_apport_gaz_dt():
    flag = dictionary_is_full(wilayas)
    if flag == True:
        region_nombre_abonne_dataframe_gaz = region_apport_gaz()
        print(region_nombre_abonne_dataframe_gaz)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})   

@app.route('/soldedf')
def soldedf():
    flag = dictionary_is_full(wilayas_creance)
    if flag == True:
        sold_df = solde_dataframe()
        print(sold_df)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})


@app.route('/pcdf')
def pcdf():
    flag = dictionary_is_full(wilayas_creance)
    if flag == True:
        pc_df = prise_en_charge_dataframe()
        print(pc_df)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})   


@app.route('/encaissement')
def encaissement():
    flag = dictionary_is_full(wilayas_creance)
    if flag == True:
        encaissement_df = encaissement_dataframe()
        print(encaissement_df)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})   

@app.route('/testing')
def testing():
    
    sold23 = load_old_data_creance()
    print(sold23)
  
    

    return jsonify({'message':'done!'}) 

@app.route('/taux')
def taux():
    
    sold = solde_dataframe()
    solde23= load_old_data_creance()
    encaissement = encaissement_dataframe()
    solde23.columns = sold.columns
    prise_en_charge = prise_en_charge_dataframe()
    solde23 = solde23.droplevel([0, 1], axis=1)
    encaissement = encaissement.droplevel([0, 1], axis=1)
    prise_en_charge = prise_en_charge.droplevel([0, 1], axis=1)
    prise_en_charge.replace(0,np.nan,inplace=True)
    solde23.replace(0,np.nan,inplace=True)


    tx_encaissement = (encaissement / (solde23 + prise_en_charge))* 100
    tx_encaissement = tx_encaissement.fillna(0)
    print(tx_encaissement)
    
    

    return jsonify({'message':'done!'}) 



@app.route('/recettes')
def recettes_tb_link():
    flag = dictionary_is_full(wilayas_creance)
    if flag == True:
        recettes_df = recettes_tb()
        print(recettes_df)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})   


@app.route('/ddc')
def ddc_link():
    flag = dictionary_is_full(wilayas_creance)
    if flag == True:
        ddc_df = ddc_tb()
        print(ddc_df)
    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'}) 


@app.route('/date')
def datee():
    print(total_days)
    return jsonify({'message':'done!'})



@app.route('/evolution')
def evolutionsolde():
    flag = dictionary_is_full(wilayas_creance)
    if flag == True:
        solde = solde_dataframe()
        solde_old = load_solde_23()
        solde_old.columns = solde.columns
        solde.replace(0,np.nan,inplace=True)
        solde_old.replace(0,np.nan,inplace=True)
        evolution = (solde - solde_old)/solde_old*100
        print(evolution)


    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'})




@app.route('/evolution2')
def evolutionsoldefinannee():
    flag = dictionary_is_full(wilayas_creance)
    if flag == True:
        solde = solde_dataframe()
        solde_old = load_old_data_creance()
        solde_old.columns = solde.columns
        solde.replace(0,np.nan,inplace=True)
        solde_old.replace(0,np.nan,inplace=True)
        evolution = (solde - solde_old)/solde_old*100
        print(evolution)


    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'}) 


@app.route('/workspace')
def workspace():
    flag = dictionary_is_full(wilayas_ventes)
    if flag == True:
        client = region_ventes()
        print(client)


    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'}) 


@app.route('/workspace2')
def workspacetwo():
    flag = dictionary_is_full(wilayas_ventes)
    if flag == True:
        client = get_cumul_ventes()
        print(client)


    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'}) 


@app.route('/workspace3')
def workspacethree():
    flag = dictionary_is_full(wilayas_ventes)
    if flag == True:
        w = get_cumul_ventes_gaz()
        print(w)
            


    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'}) 




@app.route('/workspace4')
def workspacefour():
    flag = dictionary_is_full(wilayas_ventes)
    if flag == True:
        for key in wilayas_ventes:
            print(wilayas_ventes[key][0]["ventes"])


    else:
        print('Please fill all the wilayas data')
    

    return jsonify({'message':'done!'}) 




@app.route('/get_tb')
def get_tableau_de_bord():
    flag = dictionary_is_full(wilayas)
    if flag == True:
        w = region_nombre_abonne()
        nombre_abbo_gaz = region_nombre_abonne_gaz()

        # Save the first DataFrame to Excel
        file_name = 'TB_test.xlsx'
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            w.to_excel(writer, sheet_name='Clientele', startrow=4)

        # Customize the first DataFrame in the Excel file
        customize_excel_table(
            file_name=file_name,
            sheet_name='Clientele',
            start_row=5,  # Data starts from row 5
            title="Nombre d'abonnés électricité",
            last_row_color="87CEEB"  # Light blue for the last row
        )

        # Add a gap of 4 rows before the second table
        gap_row_2nd = w.shape[0] + 9  # Adds a 4-row gap between the two tables

        # Save the second DataFrame to the same Excel sheet, with a gap between tables row-wise
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            nombre_abbo_gaz.to_excel(writer, sheet_name='Clientele', startrow=gap_row_2nd, startcol=0)

        # Customize the second DataFrame in the Excel file
        customize_excel_table(
            file_name=file_name,
            sheet_name='Clientele',
            start_row=gap_row_2nd + 1,  # Adjust the row number after the second DataFrame
            title="Nombre d'abonnés gaz",
            last_row_color="90EE90"  # Light green for the last row
        )

        # Add a gap of 4 rows before the third table
        gap_row_3rd = gap_row_2nd + nombre_abbo_gaz.shape[0] + 5  # Adds a 4-row gap between the second and third table

        accroissement_elec = region_accroissement()
        # Save the third DataFrame to the same Excel sheet, with a gap between tables row-wise
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            accroissement_elec.to_excel(writer, sheet_name='Clientele', startrow=gap_row_3rd, startcol=0)

        # Customize the third DataFrame in the Excel file
        customize_excel_table(
            file_name=file_name,
            sheet_name='Clientele',
            start_row=gap_row_3rd + 1,  # Adjust the row number after the third DataFrame
            title="Accroissement abonnées électricité",
            last_row_color="FFC0CB"  # Light pink for the last row (change this color as needed)
        )

        # Add a gap of 4 rows before the fourth table
        gap_row_4rth = gap_row_3rd + accroissement_elec.shape[0] + 5  # Adds a 4-row gap between the third and fourth table

        accroissement_gaz = region_accroissement_gaz()
        # Save the fourth DataFrame to the same Excel sheet, with a gap between tables row-wise
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            accroissement_gaz.to_excel(writer, sheet_name='Clientele', startrow=gap_row_4rth, startcol=0)

        # Customize the fourth DataFrame in the Excel file
        customize_excel_table(
            file_name=file_name,
            sheet_name='Clientele',
            start_row=gap_row_4rth + 1,  # Adjust the row number after the fourth DataFrame
            title="Accroissement abonnées gaz",
            last_row_color="D8B7DD"  # Light pink for the last row (change this color as needed)
        )

        # Add a gap of 4 rows before the fifth table
        gap_row_5th = gap_row_4rth + accroissement_gaz.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

        apport_elec = region_apport()
        # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            apport_elec.to_excel(writer, sheet_name='Clientele', startrow=gap_row_5th, startcol=0)

        # Customize the fifth DataFrame in the Excel file
        customize_excel_table(
            file_name=file_name,
            sheet_name='Clientele',
            start_row=gap_row_5th + 1,  # Adjust the row number after the fifth DataFrame
            title="Nouveaux abonnées électricité",
            last_row_color="FFFFE0"  # Light yellow for the last row
        )
        gap_row_6th = gap_row_5th + apport_elec.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

        apport_gaz = region_apport_gaz()
        # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            apport_gaz.to_excel(writer, sheet_name='Clientele', startrow=gap_row_6th, startcol=0)

        # Customize the fifth DataFrame in the Excel file
        customize_excel_table(
            file_name=file_name,
            sheet_name='Clientele',
            start_row=gap_row_6th + 1,  # Adjust the row number after the fifth DataFrame
            title="Nouveaux abonnées gaz",
            last_row_color="FFA500"  # Light yellow for the last row
        )

    else:
        print('Please fill all the wilayas data')
    
    return jsonify({'message': 'done!'})

def customize_excel_table(file_name, sheet_name, start_row, title, last_row_color):
    # Load the workbook and select the sheet
    wb = load_workbook(file_name)
    ws = wb[sheet_name]

    # Define styles
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    title_font = Font(bold=True, size=14)
    centered_alignment = Alignment(horizontal='center', vertical='center')
    last_row_fill = PatternFill(start_color=last_row_color, end_color=last_row_color, fill_type='solid')

    # Add a title
    ws.merge_cells(start_row=start_row - 1, start_column=1, end_row=start_row - 1, end_column=ws.max_column)
    title_cell = ws.cell(row=start_row - 1, column=1)
    title_cell.value = title
    title_cell.font = title_font
    title_cell.alignment = centered_alignment

    # Adjust column width and row height for the table
    for col in ws.iter_cols(min_col=1, max_col=ws.max_column):
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 15  # Adjust column width

    for row in range(start_row, ws.max_row + 1):
        ws.row_dimensions[row].height = 20  # Adjust row height

    # Apply borders and style to the table
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

    # Color the last row
    for cell in ws[ws.max_row]:
        cell.fill = last_row_fill

    # Save the workbook
    wb.save(file_name)
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
