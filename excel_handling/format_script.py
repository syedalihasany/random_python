# This is a script to format an excel file 
# to allow for easier data processing



import numpy
import pandas as pd
import os 


df = pd.read_excel(r'unformatted.xlsx')
df.to_excel('midformatted.xlsx')
df = pd.read_excel(r'midformatted.xlsx')