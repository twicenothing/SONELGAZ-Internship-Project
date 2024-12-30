import pandas as pd
import numpy as np
from data import wilayas_rnc, wilayas_rnc_gaz
from funcs import last_month_with_data_RNC




def branchements_simples_elecrticite():

    cringe=[]
    for key in wilayas_rnc:
        branchement = wilayas_rnc[key][0]['branchement']
        tb = pd.DataFrame()
        tb.loc[0,'Portfeuille ancien']= branchement.loc['Portefeuille  au 31/12/2022','janvier']
        tb.loc[0,'Demandes reçues'] = branchement.loc['reçues','Total']
        tb.loc[0,'Branchements Réalisés'] = branchement.loc['mises en service','Total']
        tb.loc[0,'Affaires Annulées'] = branchement.loc['annulées','Total']
        tb.loc[0,'Portfeuille récent'] = branchement.loc['Portefeuille au mois','Total']
        tb.loc[0,'Tx évolution'] = tb.loc[0,'Portfeuille récent'] / tb.loc[0,'Portfeuille ancien']*100-100
        cringe.append(tb)
    result = pd.concat(cringe, axis=0, ignore_index=True)
    result.index = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    result.loc['SDC']= result.sum(axis=0)
    result.loc['SDC','Portfeuille récent'] = result.loc['SDC','Portfeuille ancien'] + result.loc['SDC','Demandes reçues'] +  result.loc['SDC','Branchements Réalisés'] + result.loc['SDC','Affaires Annulées']
    result.loc['SDC','Tx évolution'] = result.loc['SDC','Portfeuille récent'] / result.loc['SDC','Portfeuille ancien'] * 100 - 100
    return result


def branchements_simples_gaz():

    cringe=[]
    for key in wilayas_rnc_gaz:
        branchement = wilayas_rnc_gaz[key][0]['branchement']
        tb = pd.DataFrame()
        tb.loc[0,'Portfeuille ancien']= branchement.loc['Portefeuille  au 31/12/2022','janvier']
        tb.loc[0,'Demandes reçues'] = branchement.loc['reçues','Total']
        tb.loc[0,'Branchements Réalisés'] = branchement.loc['mises en service','Total']
        tb.loc[0,'Affaires Annulées'] = branchement.loc['annulées','Total']
        tb.loc[0,'Portfeuille récent'] = branchement.loc['Portefeuille au mois','Total']
        tb.loc[0,'Tx évolution'] = tb.loc[0,'Portfeuille récent'] / tb.loc[0,'Portfeuille ancien']*100-100
        cringe.append(tb)
    result = pd.concat(cringe, axis=0, ignore_index=True)
    result.index = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    result.loc['SDC']= result.sum(axis=0)
    result.loc['SDC','Portfeuille récent'] = result.loc['SDC','Portfeuille ancien'] + result.loc['SDC','Demandes reçues'] +  result.loc['SDC','Branchements Réalisés'] + result.loc['SDC','Affaires Annulées']
    result.loc['SDC','Tx évolution'] = result.loc['SDC','Portfeuille récent'] / result.loc['SDC','Portfeuille ancien'] * 100 - 100
    return result




def extension_portfeuille_elec():

    cringe=[]
    for key in wilayas_rnc:
        ext = wilayas_rnc[key][0]['extension']
        tb = pd.DataFrame()
        tb.loc[0,'Portfeuille ancien']= ext.loc['Portefeuille  M-1','janvier']
        tb.loc[0,'Demandes reçues'] = ext.loc['reçues','Total']
        tb.loc[0,'Demandes Annulées'] = ext.loc['annulées','Total']
        tb.loc[0,'Affaires MES'] = ext.loc['mises en service','Total']
        tb.loc[0,'Portfeuille récent'] = ext.loc['Portefeuille au mois','Total']
        tb.loc[0,'Tx évolution'] = tb.loc[0,'Portfeuille récent'] / tb.loc[0,'Portfeuille ancien']*100-100
        cringe.append(tb)
    result = pd.concat(cringe, axis=0, ignore_index=True)
    result.index = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    result.loc['SDC']= result.sum(axis=0)
    result.loc['SDC','Portfeuille récent'] = result.loc['SDC','Portfeuille ancien'] + result.loc['SDC','Demandes reçues'] +  result.loc['SDC','Demandes Annulées'] + result.loc['SDC','Affaires MES']
    result.loc['SDC','Tx évolution'] = result.loc['SDC','Portfeuille récent'] / result.loc['SDC','Portfeuille ancien'] * 100 - 100
    return result



def extension_portfeuille_gaz():

    cringe=[]
    for key in wilayas_rnc_gaz:
        ext = wilayas_rnc_gaz[key][0]['extension']
        tb = pd.DataFrame()
        tb.loc[0,'Portfeuille ancien']= ext.loc['Portefeuille  au M-1','janvier']
        tb.loc[0,'Demandes reçues'] = ext.loc['reçues','Total']
        tb.loc[0,'Demandes Annulées'] = ext.loc['annulées','Total']
        tb.loc[0,'Affaires MES'] = ext.loc['mises en service','Total']
        tb.loc[0,'Portfeuille récent'] = ext.loc['Portefeuille au mois','Total']
        tb.loc[0,'Tx évolution'] = tb.loc[0,'Portfeuille récent'] / tb.loc[0,'Portfeuille ancien']*100-100
        cringe.append(tb)
    result = pd.concat(cringe, axis=0, ignore_index=True)
    result.index = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    result.loc['SDC']= result.sum(axis=0)
    result.loc['SDC','Portfeuille récent'] = result.loc['SDC','Portfeuille ancien'] + result.loc['SDC','Demandes reçues'] +  result.loc['SDC','Demandes Annulées'] + result.loc['SDC','Affaires MES']
    result.loc['SDC','Tx évolution'] = result.loc['SDC','Portfeuille récent'] / result.loc['SDC','Portfeuille ancien'] * 100 - 100
    return result



def extension_affaires_elec():
    cringe=[]
    for key in wilayas_rnc:
        ext = wilayas_rnc[key][0]['extension']
        ext.fillna(0,inplace=True)
        ext = ext.round()
        desired = last_month_with_data_RNC(ext)
        data = {
            "Nbre": ext.loc['étude',desired],
            "Poids": ext.loc['étude',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row = pd.DataFrame([data])
        row.columns = pd.MultiIndex.from_arrays([
            ['Etude'] * 2, 
            row.columns 
            ])
        data = {
            "Nbre": ext.loc['accord',desired],
            "Poids": ext.loc['accord',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row2 = pd.DataFrame([data])
        row2.columns = pd.MultiIndex.from_arrays([
            ['Accord'] * 2, 
            row2.columns 
            ])
        affaires_ext = pd.concat([row,row2],axis=1)
        data = {
            "Nbre": ext.loc['réalisation',desired],
            "Poids": ext.loc['réalisation',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row3 = pd.DataFrame([data])
        row3.columns = pd.MultiIndex.from_arrays([
            ['Réalisation'] * 2, 
            row3.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row3],axis=1)
        data = {
            "Nbre": ext.loc['en cours de réalisation',desired],
            "Poids": ext.loc['en cours de réalisation',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row4 = pd.DataFrame([data])
        row4.columns = pd.MultiIndex.from_arrays([
            ['Cours Réalisation'] * 2, 
            row4.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row4],axis=1)
        data = {
            "Nbre": ext.loc['MES',desired],
            "Poids": ext.loc['MES',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row5 = pd.DataFrame([data])
        row5.columns = pd.MultiIndex.from_arrays([
            ['MES'] * 2, 
            row5.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row5],axis=1)
        data = {
            "Portfeuille": ext.loc['Portefeuille au mois','Total'],
        }
        row6 = pd.DataFrame([data])
        row6.columns = pd.MultiIndex.from_arrays([
            ['Total'] * 1, 
            row6.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row6],axis=1)
        cringe.append(affaires_ext)
    affaires = pd.concat(cringe)
    wilayass = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    affaires.index = wilayass
    affaires.loc['SDC']= affaires.sum()
    affaires.loc['SDC',('Etude','Poids')] = affaires.loc['SDC',('Etude','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Accord','Poids')] = affaires.loc['SDC',('Accord','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Réalisation','Poids')] = affaires.loc['SDC',('Réalisation','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Cours Réalisation','Poids')] = affaires.loc['SDC',('Cours Réalisation','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('MES','Poids')] = affaires.loc['SDC',('MES','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100

    return affaires


def branchment_affaires_elec():
    cringe=[]
    for key in wilayas_rnc:
        ext = wilayas_rnc[key][0]['branchement']
        ext.fillna(0,inplace=True)
        ext = ext.round()
        desired = last_month_with_data_RNC(ext)
        data = {
            "Nbre": ext.loc['étude',desired],
            "Poids": ext.loc['étude',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row = pd.DataFrame([data])
        row.columns = pd.MultiIndex.from_arrays([
            ['Etude'] * 2, 
            row.columns 
            ])
        data = {
            "Nbre": ext.loc['accord',desired],
            "Poids": ext.loc['accord',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row2 = pd.DataFrame([data])
        row2.columns = pd.MultiIndex.from_arrays([
            ['Accord'] * 2, 
            row2.columns 
            ])
        affaires_ext = pd.concat([row,row2],axis=1)
        data = {
            "Nbre": ext.loc['réalisation',desired],
            "Poids": ext.loc['réalisation',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row3 = pd.DataFrame([data])
        row3.columns = pd.MultiIndex.from_arrays([
            ['Réalisation'] * 2, 
            row3.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row3],axis=1)
        data = {
            "Nbre": ext.loc['en cours de réalisation',desired],
            "Poids": ext.loc['en cours de réalisation',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row4 = pd.DataFrame([data])
        row4.columns = pd.MultiIndex.from_arrays([
            ['Cours Réalisation'] * 2, 
            row4.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row4],axis=1)
        data = {
            "Nbre": ext.loc['MES',desired],
            "Poids": ext.loc['MES',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row5 = pd.DataFrame([data])
        row5.columns = pd.MultiIndex.from_arrays([
            ['MES'] * 2, 
            row5.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row5],axis=1)
        data = {
            "Portfeuille": ext.loc['Portefeuille au mois','Total'],
        }
        row6 = pd.DataFrame([data])
        row6.columns = pd.MultiIndex.from_arrays([
            ['Total'] * 1, 
            row6.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row6],axis=1)
        cringe.append(affaires_ext)
    affaires = pd.concat(cringe)
    wilayass = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    affaires.index = wilayass
    affaires.loc['SDC']= affaires.sum()
    affaires.loc['SDC',('Etude','Poids')] = affaires.loc['SDC',('Etude','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Accord','Poids')] = affaires.loc['SDC',('Accord','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Réalisation','Poids')] = affaires.loc['SDC',('Réalisation','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Cours Réalisation','Poids')] = affaires.loc['SDC',('Cours Réalisation','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('MES','Poids')] = affaires.loc['SDC',('MES','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100

    return affaires



def extension_affaires_gaz():
    cringe=[]
    for key in wilayas_rnc_gaz:
        ext = wilayas_rnc_gaz[key][0]['extension']
        ext.fillna(0,inplace=True)
        ext = ext.round()
        desired = last_month_with_data_RNC(ext)
        data = {
            "Nbre": ext.loc['étude',desired],
            "Poids": ext.loc['étude',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row = pd.DataFrame([data])
        row.columns = pd.MultiIndex.from_arrays([
            ['Etude'] * 2, 
            row.columns 
            ])
        data = {
            "Nbre": ext.loc['accord',desired],
            "Poids": ext.loc['accord',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row2 = pd.DataFrame([data])
        row2.columns = pd.MultiIndex.from_arrays([
            ['Accord'] * 2, 
            row2.columns 
            ])
        affaires_ext = pd.concat([row,row2],axis=1)
        data = {
            "Nbre": ext.loc['réalisation',desired],
            "Poids": ext.loc['réalisation',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row3 = pd.DataFrame([data])
        row3.columns = pd.MultiIndex.from_arrays([
            ['Réalisation'] * 2, 
            row3.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row3],axis=1)
        data = {
            "Nbre": ext.loc['en cours de réalisation',desired],
            "Poids": ext.loc['en cours de réalisation',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row4 = pd.DataFrame([data])
        row4.columns = pd.MultiIndex.from_arrays([
            ['Cours Réalisation'] * 2, 
            row4.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row4],axis=1)
        data = {
            "Nbre": ext.loc['MES',desired],
            "Poids": ext.loc['MES',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row5 = pd.DataFrame([data])
        row5.columns = pd.MultiIndex.from_arrays([
            ['MES'] * 2, 
            row5.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row5],axis=1)
        data = {
            "Portfeuille": ext.loc['Portefeuille au mois','Total'],
        }
        row6 = pd.DataFrame([data])
        row6.columns = pd.MultiIndex.from_arrays([
            ['Total'] * 1, 
            row6.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row6],axis=1)
        cringe.append(affaires_ext)
    affaires = pd.concat(cringe)
    wilayass = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    affaires.index = wilayass
    affaires.loc['SDC']= affaires.sum()
    affaires.loc['SDC',('Etude','Poids')] = affaires.loc['SDC',('Etude','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Accord','Poids')] = affaires.loc['SDC',('Accord','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Réalisation','Poids')] = affaires.loc['SDC',('Réalisation','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Cours Réalisation','Poids')] = affaires.loc['SDC',('Cours Réalisation','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('MES','Poids')] = affaires.loc['SDC',('MES','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100

    return affaires


def branchment_affaires_gaz():
    cringe=[]
    for key in wilayas_rnc_gaz:
        ext = wilayas_rnc_gaz[key][0]['branchement']
        ext.fillna(0,inplace=True)
        ext = ext.round()
        desired = last_month_with_data_RNC(ext)
        data = {
            "Nbre": ext.loc['étude',desired],
            "Poids": ext.loc['étude',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row = pd.DataFrame([data])
        row.columns = pd.MultiIndex.from_arrays([
            ['Etude'] * 2, 
            row.columns 
            ])
        data = {
            "Nbre": ext.loc['accord',desired],
            "Poids": ext.loc['accord',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row2 = pd.DataFrame([data])
        row2.columns = pd.MultiIndex.from_arrays([
            ['Accord'] * 2, 
            row2.columns 
            ])
        affaires_ext = pd.concat([row,row2],axis=1)
        data = {
            "Nbre": ext.loc['réalisation',desired],
            "Poids": ext.loc['réalisation',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row3 = pd.DataFrame([data])
        row3.columns = pd.MultiIndex.from_arrays([
            ['Réalisation'] * 2, 
            row3.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row3],axis=1)
        data = {
            "Nbre": ext.loc['en cours de réalisation',desired],
            "Poids": ext.loc['en cours de réalisation',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row4 = pd.DataFrame([data])
        row4.columns = pd.MultiIndex.from_arrays([
            ['Cours Réalisation'] * 2, 
            row4.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row4],axis=1)
        data = {
            "Nbre": ext.loc['MES',desired],
            "Poids": ext.loc['MES',desired] / ext.loc['Portefeuille au mois','Total']*100
        }
        row5 = pd.DataFrame([data])
        row5.columns = pd.MultiIndex.from_arrays([
            ['MES'] * 2, 
            row5.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row5],axis=1)
        data = {
            "Portfeuille": ext.loc['Portefeuille au mois','Total'],
        }
        row6 = pd.DataFrame([data])
        row6.columns = pd.MultiIndex.from_arrays([
            ['Total'] * 1, 
            row6.columns 
            ])
        affaires_ext = pd.concat([affaires_ext,row6],axis=1)
        cringe.append(affaires_ext)
    affaires = pd.concat(cringe)
    wilayass = ['blida','bouira','medea','tiziouzou','djelfa','tipaza','boumerdes','aindefla','chlef','tissemsilt']
    affaires.index = wilayass
    affaires.loc['SDC']= affaires.sum()
    affaires.loc['SDC',('Etude','Poids')] = affaires.loc['SDC',('Etude','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Accord','Poids')] = affaires.loc['SDC',('Accord','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Réalisation','Poids')] = affaires.loc['SDC',('Réalisation','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('Cours Réalisation','Poids')] = affaires.loc['SDC',('Cours Réalisation','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100
    affaires.loc['SDC',('MES','Poids')] = affaires.loc['SDC',('MES','Nbre')]/affaires.loc['SDC',('Total','Portfeuille')]*100

    return affaires
