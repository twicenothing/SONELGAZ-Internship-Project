import pandas as pd
import numpy as np
from data import wilayas_rnc





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