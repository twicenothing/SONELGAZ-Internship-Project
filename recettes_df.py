import pandas as pd
import numpy as np
# from prise_en_charge_df import prise_en_charge_dataframe_unique
from data import * 

def prise_en_charge_dataframe_unique(df):
    sold = df
    data = {
    "AO": sold.loc["AO services commerciaux", "Prise en chatge Net "],
    "MT/MP Privées": sold.loc["MT/MP Privé", "Prise en chatge Net "],
    "FRM": sold.loc["FRM", "Prise en chatge Net "],
    "HT/HP": sold.loc["HT/HP", "Prise en chatge Net "],
    "Total Privés": (
        sold.loc["MT/MP Privé", "Prise en chatge Net "] +
        sold.loc["AO services commerciaux", "Prise en chatge Net "] +
        sold.loc["HT/HP", "Prise en chatge Net "] +
        sold.loc["FRM", "Prise en chatge Net "]
    ),
    }
    teshting = pd.DataFrame([data])
    teshting.columns = pd.MultiIndex.from_arrays([
    ['Privées'] * 5, 
    teshting.columns 
    ])
    data2 = {
    "MT/MP": sold.loc["MT/MP ADM", "Prise en chatge Net "],
    "FSM": sold.loc["FSM avec les eaux", "Prise en chatge Net "],
    "Total ADM": (
        sold.loc["MT/MP ADM", "Prise en chatge Net "] +
        sold.loc["FSM avec les eaux", "Prise en chatge Net "]
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
                sold.loc["MT/MP ADM", "Prise en chatge Net "] +
                sold.loc["FSM avec les eaux", "Prise en chatge Net "]
            ) +
            (
                sold.loc["MT/MP Privé", "Prise en chatge Net "] +
                sold.loc["AO services commerciaux", "Prise en chatge Net "] +
                sold.loc["HT/HP", "Prise en chatge Net "] +
                sold.loc["FRM", "Prise en chatge Net "]
            )
        ]
    },  # Assign the desired value
    )
    # Update hierarchical columns for the new data
    new_column.columns = pd.MultiIndex.from_tuples([("", "Total Energie")])
    solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)
    data3 = {
    "ADM": sold.loc["Travaux ADM", "Prise en chatge Net "],
    "Privé": sold.loc["Travaux prives", "Prise en chatge Net "],
    "Total": (
        sold.loc["Travaux ADM", "Prise en chatge Net "] +
        sold.loc["Travaux prives", "Prise en chatge Net "]
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
                sold.loc["MT/MP ADM", "Prise en chatge Net "] +
                sold.loc["FSM avec les eaux", "Prise en chatge Net "]
            ) +
            (
                sold.loc["MT/MP Privé", "Prise en chatge Net "] +
                sold.loc["AO services commerciaux", "Prise en chatge Net "] +
                sold.loc["HT/HP", "Prise en chatge Net "] +
                sold.loc["FRM", "Prise en chatge Net "]
            ))
            +
            (sold.loc["Travaux ADM", "Prise en chatge Net "] +
                sold.loc["Travaux prives", "Prise en chatge Net "]
            )
        ]
    },  # Assign the desired value
    )
    new_column.columns = pd.MultiIndex.from_tuples([("", "Total Créances")])
    solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)
    data4 = {
        
    "HTA privés": sold.loc["Les eaux privée", "Prise en chatge Net "],
    "HTB": sold.loc["les eaux HT-HP", "Prise en chatge Net "],
    "HTA ADM": sold.loc["Les eaux ADM", "Prise en chatge Net "],
    "FRM": np.random.random(),
    "FSM": sold.loc['les eaux fsm','Prise en chatge Net ']	
    }
    teshting4 = pd.DataFrame([data4])
    teshting4.columns = pd.MultiIndex.from_arrays([
    ['eaux'] * 5, 
    teshting4.columns 
    ])
    solde_avec_eau = pd.concat([solde_avec_eau, teshting4], axis=1)
    solde_avec_eau= solde_avec_eau.fillna(0)
    data5 = {

     #sldaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa 
    "AO": sold.loc["AO services commerciaux", "Prise en chatge Net "],
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
        'Prise en charges' if i < 13 else 'Prise en charges sans les eaux' 
        for i in range(len(original_columns))
    ]
    new_columns = pd.MultiIndex.from_tuples(
    [(group, col[0], col[1]) for col, group in zip(original_columns, new_highest_level)]
    )
    solde_avec_eau.columns = new_columns
    solde_avec_eau[('Prise en charges sans les eaux', 'Total', 'Total Energie')] = solde_avec_eau[('Prise en charges sans les eaux', 'ADM', 'Total ADM')] + solde_avec_eau[('Prise en charges sans les eaux', 'Privées', 'Total Privés')]
    teshting8 = teshting3
    new_columns = pd.MultiIndex.from_tuples(
    [('Prise en charges sans les eaux', col[0], col[1]) for col in teshting8.columns],
    names=[None, None, None]  # Avoid displaying column names
    )
    teshting8.columns = new_columns
    solde_avec_eau = pd.concat([solde_avec_eau,teshting8], axis=1)
    solde_avec_eau[('Prise en charges sans les eaux', 'Créance ', 'Total Créances')] = (
    solde_avec_eau[('Prise en charges sans les eaux', 'Total', 'Total Energie')] +
    solde_avec_eau[('Prise en charges sans les eaux', 'Travaux', 'Total')]
    )
    solde_avec_eau= solde_avec_eau.fillna(0)
    return solde_avec_eau





def recettes_tb():
    cringe=[]
    for key in wilayas_creance :
        sold = wilayas_creance[key][0]['recette']  # Make sure this points to the right DataFrame
        nmi = wilayas_creance[key][0]['solde']
        sold.columns = sold.columns.str.strip()
        data = {
        "AO": sold.loc["Total année", "AO"],
        "MT/MP Privées": sold.loc["Total année", "MT/MP Privées"],
        "FRM": sold.loc["Total année", "FRM"],
        "HT/HP": sold.loc["Total année", "HT/HP"],
        "Total Privés": (
            sold.loc["Total année", "AO"] +
            sold.loc["Total année", "MT/MP Privées"] +
            sold.loc["Total année", "FRM"] +
            sold.loc["Total année", "HT/HP"]
        ),
        }
        teshting = pd.DataFrame([data])
        teshting.columns = pd.MultiIndex.from_arrays([
        ['Privées'] * 5, 
        teshting.columns 
        ])
        data2 = {
        "MT/MP": sold.loc["Total année", "MT/MP"],
        "FSM": sold.loc["Total année", "FSM"],
        "Total ADM": sold.loc['Total année', "Total ADM"]
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
                sold.loc["Total année", "Total Energie"]
            ]
        },  # Assign the desired value
        )
        pc = prise_en_charge_dataframe_unique(nmi)
        # Update hierarchical columns for the new data
        new_column.columns = pd.MultiIndex.from_tuples([("", "Total Energie")])
        solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)
        data3 = {
        "ADM": pc[('Prise en charges','Travaux','ADM')].loc[0],
        "Privé": pc[('Prise en charges','Travaux','Privé')].loc[0],
        "Total": (
            pc[('Prise en charges','Travaux','ADM')].loc[0] +
            pc[('Prise en charges','Travaux','Privé')].loc[0]
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
                    pc[('Prise en charges','Travaux','ADM')].loc[0] +
                    pc[('Prise en charges','Travaux','Privé')].loc[0]
                ) +
                (
                    sold.loc["Total année", "AO"] +
                    sold.loc["Total année", "MT/MP Privées"] +
                    sold.loc["Total année", "FRM"] +
                    sold.loc["Total année", "HT/HP"]
                ))
                +
                (sold.loc["Total année", "MT/MP"] +
                    sold.loc["Total année", "FSM"]
                )
            ]
        },  # Assign the desired value
        )
        new_column.columns = pd.MultiIndex.from_tuples([("", "Total Créances")])
        solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)
        #loooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
        data4 = {
        "HTA privés": pc[('Prise en charges sans les eaux','eaux','HTA privés')].loc[0],
        "HTB": pc[('Prise en charges sans les eaux','eaux','HTB')].loc[0],
        "HTA ADM": pc[('Prise en charges sans les eaux','eaux','HTA ADM')].loc[0],
        "FRM": pc[('Prise en charges sans les eaux','eaux','FRM')].loc[0],
        "FSM": pc[('Prise en charges sans les eaux','eaux','FSM')].loc[0]	
        }
        teshting4 = pd.DataFrame([data4])
        teshting4.columns = pd.MultiIndex.from_arrays([
        ['eaux'] * 5, 
        teshting4.columns 
        ])
        solde_avec_eau = pd.concat([solde_avec_eau, teshting4], axis=1)
        solde_avec_eau = solde_avec_eau.fillna(0)
        data5 = {
        "AO": sold.loc["Total année", "AO"],
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
            'Recettes' if i < 13 else 'Recettes sans les eaux' 
            for i in range(len(original_columns))
        ]
        new_columns = pd.MultiIndex.from_tuples(
        [(group, col[0], col[1]) for col, group in zip(original_columns, new_highest_level)]
        )
        solde_avec_eau.columns = new_columns
        solde_avec_eau[('Recettes sans les eaux', 'Total', 'Total Energie')] = solde_avec_eau[('Recettes sans les eaux', 'ADM', 'Total ADM')] + solde_avec_eau[('Recettes sans les eaux', 'Privées', 'Total Privés')]
        teshting8 = teshting3
        new_columns = pd.MultiIndex.from_tuples(
        [('Recettes sans les eaux', col[0], col[1]) for col in teshting8.columns],
        names=[None, None, None]  # Avoid displaying column names
        )
        teshting8.columns = new_columns
        solde_avec_eau = pd.concat([solde_avec_eau,teshting8], axis=1)
        solde_avec_eau[('Recettes sans les eaux', 'Créance ', 'Total Créances')] = (
        solde_avec_eau[('Recettes sans les eaux', 'Total', 'Total Energie')] +
        solde_avec_eau[('Recettes sans les eaux', 'Travaux', 'Total')]
        )
        cringe.append(solde_avec_eau)
    sold_sans_eau24 = pd.concat(cringe)
    wilayass = ['Blida','Bouira','Médéa','T.Ouzou','Djelfa','Tipaza','Boumerdes','Ain Defla','Chlef','Tissemssilt']
    sold_sans_eau24.index = wilayass
    sum_row = sold_sans_eau24.sum()
    sold_sans_eau24.loc['RD Blida']= sum_row
    return sold_sans_eau24