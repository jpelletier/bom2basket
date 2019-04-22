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

# Add columns Buy and Stock
def CheckStock(datafr, stockfr):
    stock_list = []
    buy_list = []
    
    for i in range (datafr.size):
        buy = datafr['Total Quantity'][i]
        stock = 0
        
        # Check if current item is in stock, if so,
        # calculate buying quantities
        for stock_item in stockfr.index:

            # Item found in stock data
            if (datafr.index[i] == stockfr['DPN'][stock_item]):
                # get production and stock quantities
                qty = datafr['Total Quantity'][i]
                stock = stockfr['Quantity'][stock_item]

                # if quantity is negative, we ship back the items
                # to the suppliers :D                
                buy = qty - stock
                if (buy < 0):
                    buy = 0
                    
                #print ('----')
                #print (stock_item, stockfr['DPN'][stock_item],qty,stock,buy)
                #print ('yes')

        stock_list.append(stock)
        buy_list.append(buy)
        #print (i)

    # append the list to the dataframe
    datafr['Stock'] = stock_list
    datafr['Buy'] = buy_list
    return datafr
#------------------------------------------------------
# we don't want a full GUI, so keep the root window from appearing
Tk().withdraw()

# Get the production file specifying the BOM files and quantities
# to be produced
prod_file = askopenfilename(title='Select production file')

# Get the directory where the production file is.
prod_dir = os.path.dirname(prod_file)
os.chdir(prod_dir)

# Read the production file
prod_data = pd.read_excel(prod_file)

# Make a list of BOM files, with the corresponding production
# quantity and other data
files = prod_data.loc[:,'BOM files']
Qties = prod_data.loc[:,'Qty']
Refs = prod_data.loc[:,'Cust Ref']

# Create an empty dataframe
df = pd.DataFrame()

# Append all the BOMs together
for i in range(files.size):
    print ('Processing file: %s\n',files[i])
    
    # Open each BOM files
    data = pd.read_excel(files[i])

    # Add refs and total quantities
    data['Refs'] = Refs[i]
    data['Total Quantity'] = data['Quantity'] * Qties[i]
    df = df.append(data, ignore_index=True, sort = False)
    
# Once all merged, separate by suppliers    
#print ('separate suppliers')
df_Digikey = df[df.Dist == 'Digi-Key']
df_Newark = df[df.Dist == 'Newark']
df_Others = df[(df.Dist != 'Digi-Key') & (df.Dist != 'Newark')]

# Combine quantities by DPN and compute the total quantity for each part
#print ('Combining quantities')
df_D = df_Digikey.pivot_table(index = 'DPN',values = 'Total Quantity',aggfunc='sum')
df_N = df_Newark.pivot_table(index = 'DPN',values = 'Total Quantity',aggfunc='sum')

# For others, we include the distributors
df_O = df_Others.pivot_table(index = ['Dist','DPN'],values = 'Total Quantity',aggfunc='sum')
    
# Check if we have some stock already
print ('Check stock')
stock_file = askopenfilename(title='Select stock file')
if (stock_file != None):
    stock_data = pd.read_excel(stock_file)

    df_D_Buy = CheckStock(df_D,stock_data)
    df_N_Buy = CheckStock(df_N,stock_data)
    df_O_Buy = CheckStock(df_O,stock_data)
    
# Basket is ready, save it
fileout = asksaveasfilename(title = "Save basket file",filetypes = (("Excel files","*.xlsx"),("all files","*.*")))

# Write resulting file
writer = ExcelWriter(fileout, engine='xlsxwriter')

if (stock_file != None):
    df_D_Buy.to_excel(writer, sheet_name='Digi-Key')
    df_N_Buy.to_excel(writer, sheet_name='Newark')
    df_O_Buy.to_excel(writer, sheet_name='Others')
else:
    df_D.to_excel(writer, sheet_name='Digi-Key')
    df_N.to_excel(writer, sheet_name='Newark')
    df_O.to_excel(writer, sheet_name='Others')
    
writer.save()

print('Done\n')