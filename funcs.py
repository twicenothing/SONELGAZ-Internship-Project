import pandas as pd 
import numpy as np


## this serves to return the index of the last updated value in the frame as an int
## this serves to return the index of the last updated value in the frame as an int

def find_repeated_index(df):
    """
    Finds the index of the first row where values start repeating or become blank.
    Assumes values are identical from this row onward until the end of the DataFrame.
    
    Parameters:
    - df: The DataFrame to analyze.

    Returns:
    - Index of the first row of interest or None if none are found.
    """
    for i in range(len(df) - 1):
        current_row = df.iloc[i, :]

        # Check for blank row (NaN values)
        if current_row.isnull().all() or (current_row == 0).all():
            return i  # Return the index of the first blank row

        next_row = df.iloc[i + 1, :]

        # Check if the current row equals the next row for all columns
        if current_row.equals(next_row):
            return i  # Return the index of the first repeated row
            
    return None  # In case no repeated or blank rows are found



# this serves the same thing as the last function but returns the mounth instead of the int
def find_repeated_row(df):
    """
    Finds the index of the first row where values start repeating.
    Assumes values are identical from this row onward until the end of the DataFrame.
    
    Parameters:
    - df: The DataFrame to analyze.

    Returns:
    - Index of the first repeated row.
    """
   
    for i in range(len(df) - 1):
        # Select the current row and the next row for comparison
        if df.index.values[i] != df.index.values[-1]:
            current_row = df.iloc[i, :]
            next_row = df.iloc[i + 1, :]

            # Check if the current row equals the next row for all columns
            if current_row.equals(next_row):
                return df.index.values[i-1]   # Return the first repeated row index
        else : 
            return df.index.values[i-1]

    return None  # In case no repeated rows are found

# checks if you set all the values in the wilayas dictionary

def dictionary_is_full(dictionary):
    flag = True
    for key in dictionary.keys():
        if dictionary[key] == [] :
            flag = False
    return flag



# finds the last mounth with uploaded data in the tableau de bord to use it in the region
def last_month_with_data(df):
    # Get unique top-level columns
    months = df.columns.get_level_values(0).unique()
    
    # Check if the last column is 'CUMUL' and exclude it if present
    if months[-1] == 'CUMUL':
        months = months[:-1]
    
    # Loop in reverse to find the last month with non-zero data
    for month in reversed(months):
        # Select columns for the current month
        month_data = df[month]
        
        # Check if any value in this month's columns is non-zero
        if (month_data != 0).any().any():
            return month  # Return the month name as soon as we find non-zero data
    
    return None  # Return None if all months have only zero data



def last_month_with_data_RNC(df):
    months = df.columns
    desired = ''
    for month in reversed(months):
        # Select columns for the current month
        month_data = df.loc['étude',month]
        
        # Check if any value in this month's columns is non-zero
        if (month_data != 0):
            desired = month
            break
    return desired


def remove_null_rows_below_threshold(df):
    """
    Removes rows below the first completely null row in the DataFrame.
    Assumes non-numerical index.

    Parameters:
        df (pd.DataFrame): Input DataFrame.

    Returns:
        pd.DataFrame: Filtered DataFrame with non-null rows.
    """
    # Identify rows where all values are NaN
    null_rows = df.isnull().all(axis=1)

    # Find the first index where all values are NaN
    first_null_row_index = null_rows.idxmax() if null_rows.any() else None

    if first_null_row_index is not None:
        # Get the position of the first null row
        first_null_row_position = df.index.get_loc(first_null_row_index)

        # Keep rows up to but not including the first null row
        return df.iloc[:first_null_row_position]
    else:
        # If no completely null row, return the original DataFrame
        return df
    


def remove_zero_rows_below_threshold(df):
    """
    Removes rows below the first row where all values are zero in the DataFrame.
    Assumes non-numerical index.

    Parameters:
        df (pd.DataFrame): Input DataFrame.

    Returns:
        pd.DataFrame: Filtered DataFrame with rows before the all-zero row.
    """
    # Identify rows where all values are 0
    zero_rows = (df == 0).all(axis=1)

    # Find the first index where all values are 0
    first_zero_row_index = zero_rows.idxmax() if zero_rows.any() else None

    if first_zero_row_index is not None:
        # Get the position of the first zero row
        first_zero_row_position = df.index.get_loc(first_zero_row_index)

        # Keep rows up to but not including the first zero row
        return df.iloc[:first_zero_row_position]
    else:
        # If no completely zero row, return the original DataFrame
        return df
