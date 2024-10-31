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
