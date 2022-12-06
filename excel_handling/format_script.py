# This is a script to format an excel file 
# to allow for easier data processing



import numpy
import pandas as pd
import os
import math

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

print(start_row)
print(start_column)

for row in range(0,start_row):
    df = df.drop(labels=[row], axis=0)

df = df.drop(df.columns[start_column], axis=1)

print(df)


df.to_excel('newformatted.xlsx', header = False, index = False)