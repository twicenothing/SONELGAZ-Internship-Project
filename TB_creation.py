import pandas as pd
import numpy as np 
from region_funcs import *
from region_funcs_accroissement import * 
from region_funcs_apport import * 
from region_funcs_gaz import * 
from region_funcs_gaz_accroissement import * 
from region_funcs_gaz_apport import *
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from ventes.region_ventes import *
from ventes.region_ventes_gaz import * 
from RCN.rdc_TB import *
from creances.solde_df import *
from creances.prise_en_charge_df import *
from creances.encaissement_df import *
from creances.solde23_df import * 
from creances.recettes_df import *
from creances.ddc_df import ddc_tb
from creances.old_data import *
from elec.elec_tb import *



pd.set_option('display.float_format', '{:.2f}'.format)


def TB_clientele():
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



def TB_ventes():
    # Adding a new sheet for the 7th DataFrame
        new_sheet_name = 'ventes_achats_élec'  # The name for the new sheet (can be customized)
        gap_row_7th = 4  # Start at a custom row number for the new sheet

        # Fetch the 7th DataFrame (replace with your actual data function)
        new_sheet_data = region_ventes()  # Replace with the function for your 7th DataFrame
        with pd.ExcelWriter('TB_test.xlsx', engine='openpyxl', mode='a') as writer:
            # Write the new DataFrame to the new sheet
            new_sheet_data.to_excel(writer, sheet_name=new_sheet_name, startrow=gap_row_7th, startcol=0)

        # Customize the new DataFrame in the new sheet
        customize_excel_table(
            file_name='TB_test.xlsx',
            sheet_name=new_sheet_name,
            start_row=gap_row_7th + 1,  # Adjust the row number after the 7th DataFrame
            title="Ventes & Achats électricité (BT/MT en GWh) du mois",  # Title for the new DataFrame
            last_row_color="C8A2D4"  # Light yellow for the last row
        )
        # Assuming the 7th DataFrame was saved to the new sheet 'OtherSheet' previously
        gap_row_8th = gap_row_7th + new_sheet_data.shape[0] + 5  # Adds a 4-row gap after the 7th DataFrame
        ventes_cumul = get_cumul_ventes()
        # Save the 8th DataFrame to the new sheet, right below the 7th DataFrame
        with pd.ExcelWriter('TB_test.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            ventes_cumul.to_excel(writer, sheet_name=new_sheet_name, startrow=gap_row_8th, startcol=0)

        # Customize the 8th DataFrame in the Excel file
        customize_excel_table(
            file_name='TB_test.xlsx',
            sheet_name=new_sheet_name,
            start_row=gap_row_8th + 1,  # Adjust the row number after the 8th DataFrame
            title="Ventes & Achats électricité (BT/MT en GWh) cumulés",  # Update this title accordingly
            last_row_color="C8A2D4"  # Yellow gold, or change to any color you prefer
        )

def TB_ventes_gaz():
    # Adding a new sheet for the 7th DataFrame
        new_sheet_name = 'ventes_achats_gaz'  # The name for the new sheet (can be customized)
        gap_row_7th = 4  # Start at a custom row number for the new sheet

        # Fetch the 7th DataFrame (replace with your actual data function)
        new_sheet_data = region_ventes_gaz()  # Replace with the function for your 7th DataFrame
        with pd.ExcelWriter('TB_test.xlsx', engine='openpyxl', mode='a') as writer:
            # Write the new DataFrame to the new sheet
            new_sheet_data.to_excel(writer, sheet_name=new_sheet_name, startrow=gap_row_7th, startcol=0)

        # Customize the new DataFrame in the new sheet
        customize_excel_table(
            file_name='TB_test.xlsx',
            sheet_name=new_sheet_name,
            start_row=gap_row_7th + 1,  # Adjust the row number after the 7th DataFrame
            title="Ventes & Achats GAZ (BP/MP en Mth) du mois",  # Title for the new DataFrame
            last_row_color="C8A2D4"  # Light yellow for the last row
        )
        # Assuming the 7th DataFrame was saved to the new sheet 'OtherSheet' previously
        gap_row_8th = gap_row_7th + new_sheet_data.shape[0] + 5  # Adds a 4-row gap after the 7th DataFrame
        ventes_cumul = get_cumul_ventes_gaz()
        # Save the 8th DataFrame to the new sheet, right below the 7th DataFrame
        with pd.ExcelWriter('TB_test.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            ventes_cumul.to_excel(writer, sheet_name=new_sheet_name, startrow=gap_row_8th, startcol=0)

        # Customize the 8th DataFrame in the Excel file
        customize_excel_table(
            file_name='TB_test.xlsx',
            sheet_name=new_sheet_name,
            start_row=gap_row_8th + 1,  # Adjust the row number after the 8th DataFrame
            title="Ventes  & Achats GAZ (BP/ MP en Mth) cumulés",  # Update this title accordingly
            last_row_color="C8A2D4"  # Yellow gold, or change to any color you prefer
        )

def TB_RCN_elec():
    w = extension_portfeuille_elec()
    nombre_abbo_gaz = branchements_simples_elecrticite()
    name='RCN_Elec'
    # Save the first DataFrame to Excel
    file_name = 'TB_test.xlsx'
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
        w.to_excel(writer, sheet_name=name, startrow=4)

    # Customize the first DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=5,  # Data starts from row 5
        title="Extension Portfeuille",
        last_row_color="87CEEB"  # Light blue for the last row
    )

    # Add a gap of 4 rows before the second table
    gap_row_2nd = w.shape[0] + 9  # Adds a 4-row gap between the two tables

    # Save the second DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        nombre_abbo_gaz.to_excel(writer, sheet_name=name, startrow=gap_row_2nd, startcol=0)

    # Customize the second DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_2nd + 1,  # Adjust the row number after the second DataFrame
        title="Branchements Simples Electricité",
        last_row_color="90EE90"  # Light green for the last row
    )
     # Add a gap of 4 rows before the third table
    gap_row_3rd = gap_row_2nd + nombre_abbo_gaz.shape[0] + 5  # Adds a 4-row gap between the second and third table

    accroissement_elec = extension_affaires_elec()
    # Save the third DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        accroissement_elec.to_excel(writer, sheet_name=name, startrow=gap_row_3rd, startcol=0)

    # Customize the third DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_3rd + 1,  # Adjust the row number after the third DataFrame
        title="Extension Affaires",
        last_row_color="FFC0CB"  # Light pink for the last row (change this color as needed)
    )
    # Add a gap of 4 rows before the fourth table
    gap_row_4rth = gap_row_3rd + accroissement_elec.shape[0] + 5  # Adds a 4-row gap between the third and fourth table

    accroissement_gaz = branchment_affaires_elec()
    # Save the fourth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        accroissement_gaz.to_excel(writer, sheet_name=name, startrow=gap_row_4rth, startcol=0)

    # Customize the fourth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_4rth + 1,  # Adjust the row number after the fourth DataFrame
        title="Branchement affaires",
        last_row_color="D8B7DD"  # Light pink for the last row (change this color as needed)
    )


def TB_RCN_gaz():
    w = extension_portfeuille_gaz()
    nombre_abbo_gaz = branchements_simples_gaz()
    name='RCN_Gaz'
    # Save the first DataFrame to Excel
    file_name = 'TB_test.xlsx'
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
        w.to_excel(writer, sheet_name=name, startrow=4)

    # Customize the first DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=5,  # Data starts from row 5
        title="Extension Portfeuille",
        last_row_color="87CEEB"  # Light blue for the last row
    )

    # Add a gap of 4 rows before the second table
    gap_row_2nd = w.shape[0] + 9  # Adds a 4-row gap between the two tables

    # Save the second DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        nombre_abbo_gaz.to_excel(writer, sheet_name=name, startrow=gap_row_2nd, startcol=0)

    # Customize the second DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_2nd + 1,  # Adjust the row number after the second DataFrame
        title="Branchements Simples Gaz",
        last_row_color="90EE90"  # Light green for the last row
    )
 # Add a gap of 4 rows before the third table
    gap_row_3rd = gap_row_2nd + nombre_abbo_gaz.shape[0] + 5  # Adds a 4-row gap between the second and third table

    accroissement_elec = extension_affaires_gaz()
    # Save the third DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        accroissement_elec.to_excel(writer, sheet_name=name, startrow=gap_row_3rd, startcol=0)

    # Customize the third DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_3rd + 1,  # Adjust the row number after the third DataFrame
        title="Extension Affaires",
        last_row_color="FFC0CB"  # Light pink for the last row (change this color as needed)
    )
    # Add a gap of 4 rows before the fourth table
    gap_row_4rth = gap_row_3rd + accroissement_elec.shape[0] + 5  # Adds a 4-row gap between the third and fourth table

    accroissement_gaz = branchment_affaires_gaz()
    # Save the fourth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        accroissement_gaz.to_excel(writer, sheet_name=name, startrow=gap_row_4rth, startcol=0)

    # Customize the fourth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_4rth + 1,  # Adjust the row number after the fourth DataFrame
        title="Branchement affaires",
        last_row_color="D8B7DD"  # Light pink for the last row (change this color as needed)
    )


def TB_solde():
    w = solde_dataframe()
    name='calcul-DCC-TX-Encais'
    # Save the first DataFrame to Excel
    file_name = 'TB_test.xlsx'
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
        w.to_excel(writer, sheet_name=name, startrow=4)

    # Customize the first DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=5,  # Data starts from row 5
        title="2024",
        last_row_color="87CEEB"  # Light blue for the last row
    )
     # Add a gap of 4 rows before the second table
    gap_row_2nd = w.shape[0] + 9  # Adds a 4-row gap between the two tables
    nombre_abbo_gaz = prise_en_charge_dataframe()
    # Save the second DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        nombre_abbo_gaz.to_excel(writer, sheet_name=name, startrow=gap_row_2nd, startcol=0)

    # Customize the second DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_2nd + 1,  # Adjust the row number after the second DataFrame
        title="2024",
        last_row_color="90EE90"  # Light green for the last row
    )
    gap_row_3rd = gap_row_2nd + nombre_abbo_gaz.shape[0] + 5  # Adds a 4-row gap between the second and third table

    accroissement_elec = encaissement_dataframe()
    # Save the third DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        accroissement_elec.to_excel(writer, sheet_name=name, startrow=gap_row_3rd, startcol=0)

    # Customize the third DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_3rd + 1,  # Adjust the row number after the third DataFrame
        title="2024",
        last_row_color="FFC0CB"  # Light pink for the last row (change this color as needed)
    )

  

    solde23= load_old_data_creance()
    solde23.columns = w.columns
    solde23 = solde23.droplevel([0, 1], axis=1)
    accroissement_elec = accroissement_elec.droplevel([0, 1], axis=1)
    nombre_abbo_gaz = nombre_abbo_gaz.droplevel([0, 1], axis=1)
    nombre_abbo_gaz.replace(0,np.nan,inplace=True)
    solde23.replace(0,np.nan,inplace=True)


    tx_encaissement = (accroissement_elec / (solde23 + nombre_abbo_gaz))* 100
    tx_encaissement = tx_encaissement.fillna(0)

    # Correct calculation for gap_row_4rth
    gap_row_4rth = gap_row_3rd + accroissement_elec.shape[0] + 5  # Adds a 4-row gap after the third DataFrame

    # Save the fourth DataFrame (tx_encaissement) to the same Excel sheet
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        tx_encaissement.to_excel(writer, sheet_name=name, startrow=gap_row_4rth, startcol=0)

    # Customize the fourth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_4rth + 1,  # Adjust the row number for the title
        title="2024",
        last_row_color="D8B7DD"  # Light purple for the last row
    )
    # Add a gap of 4 rows before the fifth table
    gap_row_5th = gap_row_4rth + tx_encaissement.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    apport_elec = recettes_tb()
    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        apport_elec.to_excel(writer, sheet_name=name, startrow=gap_row_5th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_5th + 1,  # Adjust the row number after the fifth DataFrame
        title="2024",
        last_row_color="FFFFE0"  # Light yellow for the last row
    )
    gap_row_6th = gap_row_5th + apport_elec.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    apport_gaz = ddc_tb()
    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        apport_gaz.to_excel(writer, sheet_name=name, startrow=gap_row_6th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_6th + 1,  # Adjust the row number after the fifth DataFrame
        title="2024 DDC",
        last_row_color="FFA500"  # Light yellow for the last row
    )
    solde23 = load_solde_23()
    solde23.columns = w.columns
    w.replace(0,np.nan,inplace=True)
    solde23.replace(0,np.nan,inplace=True)
    evolution = (w - solde23)/solde23*100

    gap_row_7th = gap_row_6th + apport_gaz.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        evolution.to_excel(writer, sheet_name=name, startrow=gap_row_7th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_7th + 1,  # Adjust the row number after the fifth DataFrame
        title="Evolution solde /mois",
        last_row_color="00FFFF"  # Light yellow for the last row
    )

    solde_old = load_old_data_creance()
    solde_old.columns = w.columns
    w.replace(0,np.nan,inplace=True)
    solde_old.replace(0,np.nan,inplace=True)
    evolution2 = (w - solde_old)/solde_old*100
    gap_row_8th = gap_row_7th + evolution.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        evolution2.to_excel(writer, sheet_name=name, startrow=gap_row_8th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_8th + 1,  # Adjust the row number after the fifth DataFrame
        title="Evolution solde /fin d'année",
        last_row_color="D3D3D3"  # Light yellow for the last row
    )



def TB_elec():
    w = nombre_abo_tb()
    name='élec'
    # Save the first DataFrame to Excel
    file_name = 'TB_test.xlsx'
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
        w.to_excel(writer, sheet_name=name, startrow=4)

    # Customize the first DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=5,  # Data starts from row 5
        title="I-  Nombre d'abonnés",
        last_row_color="87CEEB"  # Light blue for the last row
    )
     # Add a gap of 4 rows before the second table
    gap_row_2nd = w.shape[0] + 9  # Adds a 4-row gap between the two tables
    nombre_abbo_gaz = accroissement_tb()
    # Save the second DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        nombre_abbo_gaz.to_excel(writer, sheet_name=name, startrow=gap_row_2nd, startcol=0)

    # Customize the second DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_2nd + 1,  # Adjust the row number after the second DataFrame
        title="II-  Accroissement abonnés",
        last_row_color="90EE90"  # Light green for the last row
    )
    gap_row_3rd = gap_row_2nd + nombre_abbo_gaz.shape[0] + 5  # Adds a 4-row gap between the second and third table

    accroissement_elec = resiliation_tb()
    # Save the third DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        accroissement_elec.to_excel(writer, sheet_name=name, startrow=gap_row_3rd, startcol=0)

    # Customize the third DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_3rd + 1,  # Adjust the row number after the third DataFrame
        title="II-  Résiliation",
        last_row_color="FFC0CB"  # Light pink for the last row (change this color as needed)
    )

  



    tx_encaissement = apport_tb()

    # Correct calculation for gap_row_4rth
    gap_row_4rth = gap_row_3rd + accroissement_elec.shape[0] + 5  # Adds a 4-row gap after the third DataFrame

    # Save the fourth DataFrame (tx_encaissement) to the same Excel sheet
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        tx_encaissement.to_excel(writer, sheet_name=name, startrow=gap_row_4rth, startcol=0)

    # Customize the fourth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_4rth + 1,  # Adjust the row number for the title
        title="II-  Apport",
        last_row_color="D8B7DD"  # Light purple for the last row
    )
    # Add a gap of 4 rows before the fifth table
    gap_row_5th = gap_row_4rth + tx_encaissement.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    apport_elec = reabonne_tb()
    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        apport_elec.to_excel(writer, sheet_name=name, startrow=gap_row_5th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_5th + 1,  # Adjust the row number after the fifth DataFrame
        title="II-  Réabonnement",
        last_row_color="FFFFE0"  # Light yellow for the last row
    )
    gap_row_6th = gap_row_5th + apport_elec.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    apport_gaz = apport_nv_tb()
    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        apport_gaz.to_excel(writer, sheet_name=name, startrow=gap_row_6th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_6th + 1,  # Adjust the row number after the fifth DataFrame
        title="Apport nouveau",
        last_row_color="FFA500"  # Light yellow for the last row
    )

    evolution = ventes_tb()

    gap_row_7th = gap_row_6th + apport_gaz.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        evolution.to_excel(writer, sheet_name=name, startrow=gap_row_7th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_7th + 1,  # Adjust the row number after the fifth DataFrame
        title="III-  Ventes électricité (GWh)",
        last_row_color="00FFFF"  # Light yellow for the last row
    )

    evolution2 = achats_tb()
    gap_row_8th = gap_row_7th + evolution.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        evolution2.to_excel(writer, sheet_name=name, startrow=gap_row_8th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_8th + 1,  # Adjust the row number after the fifth DataFrame
        title="IV-  Achats électricité & pertes d'énergie (GWh/%)",
        last_row_color="D3D3D3"  # Light yellow for the last row
    )
    chiffre_aff = chiffre_aff_tb()
    gap_row_9th = gap_row_8th + evolution2.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        chiffre_aff.to_excel(writer, sheet_name=name, startrow=gap_row_9th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_9th + 1,  # Adjust the row number after the fifth DataFrame
        title="V-  Chiffre d'affaires électricité (MDA)",
        last_row_color="D3D3D3"  # Light yellow for the last row
    )

    prix = prix_tb()
    gap_row_10th = gap_row_9th + chiffre_aff.shape[0] + 5  # Adds a 4-row gap between the fourth and fifth table

    # Save the fifth DataFrame to the same Excel sheet, with a gap between tables row-wise
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        prix.to_excel(writer, sheet_name=name, startrow=gap_row_10th, startcol=0)

    # Customize the fifth DataFrame in the Excel file
    customize_excel_table(
        file_name=file_name,
        sheet_name=name,
        start_row=gap_row_10th + 1,  # Adjust the row number after the fifth DataFrame
        title="VI- Prix de Ventes Moyen électricité (cDA)",
        last_row_color="D3D3D3"  # Light yellow for the last row
    )