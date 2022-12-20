# This is a script to format the BOM excel file
# before it can be imported in Windchill Risk and Reliability
# Author: Syed Ali Hasany

# To run this file from anywhere (in any directory)
# add the path of this script file to the PATH environment variable
# simply run >> format_script.py

import numpy
import pandas as pd
import os
import math

# The functions below, valid_row and valid_column, return the row and column
# numbers after which the valid data starts

def valid_row(df):
    rows    = len(df.axes[0])
    columns = len(df.axes[1])
    count = 0
    for row in range(rows):
        count = 0
        for column in range(columns):
            if( type(df.iloc[row,column]) == str ):
                count+=1
                if (count >= 4):
                    return row

def valid_column(df):
    rows    = len(df.axes[0])
    columns = len(df.axes[1])
    count = 0
    for column in range(columns):
        count = 0
        for row in range(rows):
            if( math.isnan(df.iloc[row,column]) ):
                count+=1
                if (count >= 5):
                    return column


df = pd.read_excel(r'unformatted.xlsx')

rows    = len(df.axes[0])
columns = len(df.axes[1])

start_row = valid_row(df)
start_column = valid_column(df)


# dropping rows before the valid row
for row in range(0,start_row):
    df = df.drop(labels=[row], axis=0)

# dropping the first column (this part requires more sophistication)
df = df.drop(df.columns[start_column], axis=1)


df.to_excel('new_formatted.xlsx', header = False, index = False)