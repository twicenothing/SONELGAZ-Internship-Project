import pandas as pd 
import numpy as np
from data import *

pd.set_option('display.float_format', '{:.2f}'.format)


def encaissement_dataframe():
#     for key, filename in zip(wilayas_solde.keys(), file_creance):
# # Call your function and unpack the returned values
#         solde = load_process_creances(filename)
        
#         # Append the returned values to the respective key
#         wilayas_solde[key].append({
#             "solde": solde,
#         })
    cringe=[]
    for key in wilayas_creance:
        sold = wilayas_creance[key][0]['solde']  # Make sure this points to the right DataFrame
        data = {
        "AO": sold.loc["A.O", "Encaissement"],
        "MT/MP Privées": sold.loc["MT/MP Privé", "Encaissement"],
        "FRM": sold.loc["FRM", "Encaissement"],
        "HT/HP": sold.loc["HT/HP", "Encaissement"],
        "Total Privés": (
            sold.loc["MT/MP Privé", "Encaissement"] +
            sold.loc["A.O", "Encaissement"] +
            sold.loc["HT/HP", "Encaissement"] +
            sold.loc["FRM", "Encaissement"]
        ),
        }
        teshting = pd.DataFrame([data])
        teshting.columns = pd.MultiIndex.from_arrays([
        ['Privées'] * 5, 
        teshting.columns 
        ])
        data2 = {
        "MT/MP": sold.loc["MT/MP ADM", "Encaissement"],
        "FSM": sold.loc["FSM avec les eaux", "Encaissement"],
        "Total ADM": (
            sold.loc["MT/MP ADM", "Encaissement"] +
            sold.loc["FSM avec les eaux", "Encaissement"]
        ),
    }
        teshting2 = pd.DataFrame([data2])
        teshting2.columns = pd.MultiIndex.from_arrays([
        ['ADM'] * 3, 
        teshting2.columns 
        ])
        solde_avec_eau = pd.concat([teshting,teshting2],axis =1)
        new_column = pd.DataFrame(
        {
            "Total Energie": [
                (
                    sold.loc["MT/MP ADM", "Encaissement"] +
                    sold.loc["FSM avec les eaux", "Encaissement"]
                ) +
                (
                    sold.loc["MT/MP Privé", "Encaissement"] +
                    sold.loc["A.O", "Encaissement"] +
                    sold.loc["HT/HP", "Encaissement"] +
                    sold.loc["FRM", "Encaissement"]
                )
            ]
        },  # Assign the desired value
        )
        # Update hierarchical columns for the new data
        new_column.columns = pd.MultiIndex.from_tuples([("", "Total Energie")])
        solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)
        data3 = {
        "ADM": sold.loc["Travaux ADM", "Encaissement"],
        "Privé": sold.loc["Travaux prives", "Encaissement"],
        "Total": (
            sold.loc["Travaux ADM", "Encaissement"] +
            sold.loc["Travaux prives", "Encaissement"]
        ),
        }
        teshting3= pd.DataFrame([data3])
        teshting3.columns = pd.MultiIndex.from_arrays([
        ['Travaux'] * 3, 
        teshting3.columns 
        ])
        solde_avec_eau = pd.concat([solde_avec_eau, teshting3], axis=1)
        new_column = pd.DataFrame(
        {
                "Total Créances": [
                ((
                    sold.loc["MT/MP ADM", "Encaissement"] +
                    sold.loc["FSM avec les eaux", "Encaissement"]
                ) +
                (
                    sold.loc["MT/MP Privé", "Encaissement"] +
                    sold.loc["A.O", "Encaissement"] +
                    sold.loc["HT/HP", "Encaissement"] +
                    sold.loc["FRM", "Encaissement"]
                ))
                +
                (sold.loc["Travaux ADM", "Encaissement"] +
                    sold.loc["Travaux prives", "Encaissement"]
                )
            ]
        },  # Assign the desired value
        )
        new_column.columns = pd.MultiIndex.from_tuples([("", "Total Créances")])
        solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)
        #loooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
        data4 = {
        "HTA privés": sold.loc["Les eaux privée", "Encaissement"],
        "HTB": sold.loc["les eaux HT-HP", "Encaissement"],
        "HTA ADM": sold.loc["Les eaux ADM", "Encaissement"],
        "FRM": np.random.random(),
        "FSM": sold.loc['les eaux fsm','Encaissement']	
        }
        teshting4 = pd.DataFrame([data4])
        teshting4.columns = pd.MultiIndex.from_arrays([
        ['eaux'] * 5, 
        teshting4.columns 
        ])
        solde_avec_eau = pd.concat([solde_avec_eau, teshting4], axis=1)
        solde_avec_eau = solde_avec_eau.fillna(0)
        data5 = {
        "AO": sold.loc["A.O", "Encaissement"],
        "MT/MP Privées": solde_avec_eau[('Privées','MT/MP Privées')].iloc[0] - solde_avec_eau[('eaux','HTA privés')].iloc[0],
        "FRM": solde_avec_eau[('Privées','FRM')].iloc[0] - solde_avec_eau[('eaux','FRM')].iloc[0],
        "HT/HP": solde_avec_eau[('Privées','HT/HP')].iloc[0] - solde_avec_eau[('eaux','HTB')].iloc[0],
        }
        data5["Total Privés"] = sum(data5[key] for key in ["AO", "MT/MP Privées", "FRM", "HT/HP"])
        teshting5 = pd.DataFrame([data5])
        teshting5.columns = pd.MultiIndex.from_arrays([
        ['Privées'] * 5, 
        teshting5.columns 
        ])
        data6 = {
        "MT/MP": solde_avec_eau[('ADM','MT/MP')].iloc[0] - solde_avec_eau[('eaux','HTA ADM')].iloc[0],
        "FSM": solde_avec_eau[('ADM','FSM')].iloc[0] - solde_avec_eau[('eaux','FSM')].iloc[0],
            
        }
        data6["Total ADM"] = sum(data6[key] for key in ["MT/MP", "FSM"])
        teshting6 = pd.DataFrame([data6])
        teshting6.columns = pd.MultiIndex.from_arrays([
        ['ADM'] * 3, 
        teshting6.columns 
        ])
        data = pd.concat([teshting5, teshting6], axis=1)
        solde_avec_eau = pd.concat([solde_avec_eau, data], axis=1)
        original_columns = solde_avec_eau.columns

        new_highest_level = [
            'Encaissement' if i < 13 else 'Encaissement sans les eaux' 
            for i in range(len(original_columns))
        ]
        new_columns = pd.MultiIndex.from_tuples(
        [(group, col[0], col[1]) for col, group in zip(original_columns, new_highest_level)]
        )
        solde_avec_eau.columns = new_columns
        solde_avec_eau[('Encaissement sans les eaux', 'Total', 'Total Energie')] = solde_avec_eau[('Encaissement sans les eaux', 'ADM', 'Total ADM')] + solde_avec_eau[('Encaissement sans les eaux', 'Privées', 'Total Privés')]
        teshting8 = teshting3
        new_columns = pd.MultiIndex.from_tuples(
        [('Encaissement sans les eaux', col[0], col[1]) for col in teshting8.columns],
        names=[None, None, None]  # Avoid displaying column names
        )
        teshting8.columns = new_columns
        solde_avec_eau = pd.concat([solde_avec_eau,teshting8], axis=1)
        solde_avec_eau[('Encaissement sans les eaux', 'Créance ', 'Total Créances')] = (
        solde_avec_eau[('Encaissement sans les eaux', 'Total', 'Total Energie')] +
        solde_avec_eau[('Encaissement sans les eaux', 'Travaux', 'Total')]
        )
        cringe.append(solde_avec_eau)
    sold_sans_eau24 = pd.concat(cringe)
    wilayass = ['Blida','Bouira','Médéa','T.Ouzou','Djelfa','Tipaza','Boumerdes','Ain Defla','Chlef','Tissemssilt']
    sold_sans_eau24.index = wilayass
    sum_row = sold_sans_eau24.sum()
    sold_sans_eau24.loc['RD Blida']= sum_row
    sold_sans_eau24 = sold_sans_eau24/1000
    sold_sans_eau24 = sold_sans_eau24.round(2)

    return sold_sans_eau24