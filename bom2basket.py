# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 12:33:49 2019

@author: jpelleti
"""
import os.path
import pandas as pd

from pandas import ExcelWriter
from pandas import ExcelFile

from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename


#root = Tk()
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing

file = askopenfilename(title='Choose a file')
prod_dir = os.path.dirname(file)
os.chdir(prod_dir)

prod_data = pd.read_excel(file)
files = prod_data.loc[:,'BOM files']
Qties = prod_data.loc[:,'Qty']
Refs = prod_data.loc[:,'Cust Ref']

# Create an empty dataframe
df = pd.DataFrame()

#rows, columns = df.shape

# Append all the BOMs together
for i in range(files.size):
    data = pd.read_excel(files[i])

    # Add refs and total quantities
    data['Refs'] = Refs[i]
    data['Total Quantity'] = data['Quantity'] * Qties[i]
    df = df.append(data, ignore_index=True)
    
# Once all merged, separate by suppliers    
df_Digikey = df[df.Dist == 'Digi-Key']
df_Newark = df[df.Dist == 'Newark']
df_Others = df[(df.Dist != 'Digi-Key') & (df.Dist != 'Newark')]

# Combine quantities by DPN, MPN or PN
# Fonctionne
#df_D = df_Digikey.pivot_table(index = 'DPN',values = 'Total Quantity',aggfunc='sum')
#df_N = df_Newark.pivot_table(index = 'DPN',values = 'Total Quantity',aggfunc='sum')
#df_O = df_Others.pivot_table(index = 'DPN',values = 'Total Quantity',aggfunc='sum')
df_D = df_Digikey.groupby(['DPN'],agg({'Total Quantity':{'Qty2':'sum'}})
df_N = df_Newark.groupby(['DPN'],agg({'Total Quantity':{'Qty3':'sum'}})
df_O = df_Others.groupby(['DPN'],agg({'Total Quantity':{'Qty4':'sum'}})
    
fileout = asksaveasfilename(title = "Select file",filetypes = (("Excel files","*.xlsx"),("all files","*.*")))

# Write resulting file
writer = ExcelWriter(fileout, engine='xlsxwriter')
df_D.to_excel(writer, sheet_name='Digi-Key')
df_N.to_excel(writer, sheet_name='Newark')
df_O.to_excel(writer, sheet_name='Others')
writer.save()

# Close Tk properly
#root.destroy()
