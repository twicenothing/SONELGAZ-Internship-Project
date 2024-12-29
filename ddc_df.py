import pandas as pd 
import numpy as np 
from data import *

def solde_dataframe_unique(df):
    sold = df
    data = {
    "AO": sold.loc["AO services commerciaux", "Solde Final"],
    "MT/MP Privées": sold.loc["MT/MP Privé", "Solde Final"],
    "FRM": sold.loc["FRM", "Solde Final"],
    "HT/HP": sold.loc["HT/HP", "Solde Final"],
    "Total Privés": (
        sold.loc["MT/MP Privé", "Solde Final"] +
        sold.loc["AO services commerciaux", "Solde Final"] +
        sold.loc["HT/HP", "Solde Final"] +
        sold.loc["FRM", "Solde Final"]
    ),
    }
    teshting = pd.DataFrame([data])
    teshting.columns = pd.MultiIndex.from_arrays([
    ['Privées'] * 5, 
    teshting.columns 
    ])
    data2 = {
    "MT/MP": sold.loc["MT/MP ADM", "Solde Final"],
    "FSM": sold.loc["FSM avec les eaux", "Solde Final"],
    "Total ADM": (
        sold.loc["MT/MP ADM", "Solde Final"] +
        sold.loc["FSM avec les eaux", "Solde Final"]
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
                sold.loc["MT/MP ADM", "Solde Final"] +
                sold.loc["FSM avec les eaux", "Solde Final"]
            ) +
            (
                sold.loc["MT/MP Privé", "Solde Final"] +
                sold.loc["AO services commerciaux", "Solde Final"] +
                sold.loc["HT/HP", "Solde Final"] +
                sold.loc["FRM", "Solde Final"]
            )
        ]
    },  # Assign the desired value
    )
    # Update hierarchical columns for the new data
    new_column.columns = pd.MultiIndex.from_tuples([("", "Total Energie")])
    solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)
    data3 = {
    "ADM": sold.loc["Travaux ADM", "Solde Final"],
    "Privé": sold.loc["Travaux prives", "Solde Final"],
    "Total": (
        sold.loc["Travaux ADM", "Solde Final"] +
        sold.loc["Travaux prives", "Solde Final"]
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
                sold.loc["MT/MP ADM", "Solde Final"] +
                sold.loc["FSM avec les eaux", "Solde Final"]
            ) +
            (
                sold.loc["MT/MP Privé", "Solde Final"] +
                sold.loc["AO services commerciaux", "Solde Final"] +
                sold.loc["HT/HP", "Solde Final"] +
                sold.loc["FRM", "Solde Final"]
            ))
            +
            (sold.loc["Travaux ADM", "Solde Final"] +
                sold.loc["Travaux prives", "Solde Final"]
            )
        ]
    },  # Assign the desired value
    )
    new_column.columns = pd.MultiIndex.from_tuples([("", "Total Créances")])
    solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)

    #HTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
    data4 = {
    "HTA privés": sold.loc["Les eaux privée", "Solde Final"],
    "HTB": sold.loc["les eaux HT-HP", "Solde Final"],
    "HTA ADM": sold.loc["Les eaux ADM", "Solde Final"],
    "FRM": np.random.random(),
    "FSM": sold.loc['les eaux fsm','Solde Final']	
    }
    teshting4 = pd.DataFrame([data4])
    teshting4.columns = pd.MultiIndex.from_arrays([
    ['eaux'] * 5, 
    teshting4.columns 
    ])
    solde_avec_eau = pd.concat([solde_avec_eau, teshting4], axis=1)
    #fixaggeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee
    data5 = {
    "AO": sold.loc["AO services commerciaux", "Solde Final"],
    "MT/MP Privées": solde_avec_eau[('Privées','MT/MP Privées')].iloc[0] - solde_avec_eau[('eaux','HTA privés')].iloc[0],
    "FRM": solde_avec_eau[('Privées','FRM')].iloc[0] - solde_avec_eau[('eaux','FRM')].iloc[0],
    "HT/HP": solde_avec_eau[('Privées','FRM')].iloc[0] - solde_avec_eau[('eaux','FRM')].iloc[0],
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
        'solde avec eaux' if i < 13 else 'solde sans les eaux' 
        for i in range(len(original_columns))
    ]
    new_columns = pd.MultiIndex.from_tuples(
    [(group, col[0], col[1]) for col, group in zip(original_columns, new_highest_level)]
    )
    solde_avec_eau.columns = new_columns
    solde_avec_eau[('solde sans les eaux', 'Total', 'Total Energie')] = solde_avec_eau[('solde sans les eaux', 'ADM', 'Total ADM')] + solde_avec_eau[('solde sans les eaux', 'Privées', 'Total Privés')]
    teshting8 = teshting3
    new_columns = pd.MultiIndex.from_tuples(
    [('solde sans les eaux', col[0], col[1]) for col in teshting8.columns],
    names=[None, None, None]  # Avoid displaying column names
    )
    teshting8.columns = new_columns
    solde_avec_eau = pd.concat([solde_avec_eau,teshting8], axis=1)
    solde_avec_eau[('solde sans les eaux', 'Créance ', 'Total Créances')] = (
    solde_avec_eau[('solde sans les eaux', 'Total', 'Total Energie')] +
    solde_avec_eau[('solde sans les eaux', 'Travaux', 'Total')]
    )
    return solde_avec_eau



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

























def recettes_tb_unique(df,df1):
    sold = df
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
    pc = prise_en_charge_dataframe_unique(df1)
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
    return solde_avec_eau



















def ddc_tb():
    cringe=[]
    for key in wilayas_creance:
        recette_df = wilayas_creance[key][0]['recette']
        solde_df = wilayas_creance[key][0]['solde']
        recette = recettes_tb_unique(recette_df,solde_df)
        solde = solde_dataframe_unique(solde_df)
        recette = recette.replace(0, np.nan)
        recette = recette/1000
        solde = solde.replace(0, np.nan)
        solde = solde/1000
        n = 243
        data = {
        "AO": solde[('solde avec eaux','Privées','AO')].loc[0]/recette[('Recettes','Privées','AO')].loc[0]*n,
        "MT/MP Privées": solde[('solde avec eaux','Privées','MT/MP Privées')].loc[0]/recette[('Recettes','Privées','MT/MP Privées')].loc[0]*n,
        "FRM": solde[('solde avec eaux','Privées','FRM')].loc[0]/recette[('Recettes','Privées','FRM')].loc[0]*n,
        "HT/HP": solde[('solde avec eaux','Privées','HT/HP')].loc[0]/recette[('Recettes','Privées','HT/HP')].loc[0]*n,
        "Total Privés": solde[('solde avec eaux','Privées','Total Privés')].loc[0]/recette[('Recettes','Privées','Total Privés')].loc[0]*n,
        }
        teshting = pd.DataFrame([data])
        teshting.columns = pd.MultiIndex.from_arrays([
        ['Privées'] * 5, 
        teshting.columns 
        ])
        data2 = {
        "MT/MP": solde[('solde avec eaux','ADM','MT/MP')].loc[0]/recette[('Recettes','ADM','MT/MP')].loc[0]*n,
        "FSM": solde[('solde avec eaux','ADM','FSM')].loc[0]/recette[('Recettes','ADM','FSM')].loc[0]*n,
        "Total ADM": solde[('solde avec eaux','ADM','Total ADM')].loc[0]/recette[('Recettes','ADM','Total ADM')].loc[0]*n,
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
                solde[('solde avec eaux','','Total Energie')].loc[0]/recette[('Recettes','','Total Energie')].loc[0]*n,
            ]
        },  # Assign the desired value
        )
        # Update hierarchical columns for the new data
        new_column.columns = pd.MultiIndex.from_tuples([("", "Total Energie")])
        solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)
        data3 = {
        "ADM": solde[('solde avec eaux','Travaux','ADM')].loc[0]/recette[('Recettes','Travaux','ADM')].loc[0]*n,
        "Privé": solde[('solde avec eaux','Travaux','Privé')].loc[0]/recette[('Recettes','Travaux','Privé')].loc[0]*n,
        "Total": 
            solde[('solde avec eaux','Travaux','Total')].loc[0]/recette[('Recettes','Travaux','Total')].loc[0]*n,
        
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
                solde[('solde avec eaux','','Total Créances')].loc[0]/recette[('Recettes','','Total Créances')].loc[0]*n,
            ]
        },  # Assign the desired value
        )
        new_column.columns = pd.MultiIndex.from_tuples([("", "Total Créances")])
        solde_avec_eau = pd.concat([solde_avec_eau, new_column], axis=1)
        #loooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
        data4 = {
        "HTA privés": solde[('solde sans les eaux','eaux','HTA privés')].loc[0]/recette[('Recettes sans les eaux','eaux','HTA privés')].loc[0]*n,
        "HTB": solde[('solde sans les eaux','eaux','HTB')].loc[0]/recette[('Recettes sans les eaux','eaux','HTB')].loc[0]*n,
        "HTA ADM": solde[('solde sans les eaux','eaux','HTA ADM')].loc[0]/recette[('Recettes sans les eaux','eaux','HTA ADM')].loc[0]*n,
        "FRM": solde[('solde sans les eaux','eaux','FRM')].loc[0]/recette[('Recettes sans les eaux','eaux','FRM')].loc[0]*n,
        "FSM": solde[('solde sans les eaux','eaux','FSM')].loc[0]/recette[('Recettes sans les eaux','eaux','FSM')].loc[0]*n,	
        }
        teshting4 = pd.DataFrame([data4])
        teshting4.columns = pd.MultiIndex.from_arrays([
        ['eaux'] * 5, 
        teshting4.columns 
        ])
        solde_avec_eau = pd.concat([solde_avec_eau, teshting4], axis=1)
        solde_avec_eau = solde_avec_eau.fillna(0)
        data5 = {
        "AO": solde[('solde sans les eaux','Privées','AO')].loc[0]/recette[('Recettes sans les eaux','Privées','AO')].loc[0]*n,
        "MT/MP Privées": solde[('solde sans les eaux','Privées','MT/MP Privées')].loc[0]/recette[('Recettes sans les eaux','Privées','MT/MP Privées')].loc[0]*n,
        "FRM": solde[('solde sans les eaux','Privées','FRM')].loc[0]/recette[('Recettes sans les eaux','Privées','FRM')].loc[0]*n,
        "HT/HP": solde[('solde sans les eaux','Privées','HT/HP')].loc[0]/recette[('Recettes sans les eaux','Privées','HT/HP')].loc[0]*n,
        }
        data5["Total Privés"] = sum(data5[key] for key in ["AO", "MT/MP Privées", "FRM", "HT/HP"])
        teshting5 = pd.DataFrame([data5])
        teshting5.columns = pd.MultiIndex.from_arrays([
        ['Privées'] * 5, 
        teshting5.columns 
        ])
        data6 = {
        "MT/MP": solde[('solde sans les eaux','ADM','MT/MP')].loc[0]/recette[('Recettes sans les eaux','ADM','MT/MP')].loc[0]*n,
        "FSM": solde[('solde sans les eaux','ADM','FSM')].loc[0]/recette[('Recettes sans les eaux','ADM','FSM')].loc[0]*n,
            
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
        solde_avec_eau = solde_avec_eau.round()
        cringe.append(solde_avec_eau)
    sold_sans_eau24 = pd.concat(cringe)
    wilayass = ['Blida','Bouira','Médéa','T.Ouzou','Djelfa','Tipaza','Boumerdes','Ain Defla','Chlef','Tissemssilt']
    sold_sans_eau24.index = wilayass
    sum_row = sold_sans_eau24.sum()
    sold_sans_eau24.loc['RD Blida']= sum_row
    return sold_sans_eau24