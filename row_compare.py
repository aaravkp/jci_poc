import pandas as pd
import numpy as np
import openpyxl
import os
import xlrd

#df = pd.read_excel('/Users/arvindkumar/Documents/JCI/Input/Book1.xlsx', sheetname='Sheet1')

df = pd.read_excel('/Users/arvindkumar/Documents/JCI/Input/Book1.xlsx', sheetname='Sheet1')
#print(df)
sdf= df[['Part name', 'Product Name', 'Part Cost','Part Number']]
print(sdf)

#Count Rows
#print(len(df.axes[0]))

#Count Columns
#print(len(df.axes[1]))

print(len(sdf.axes[0]))

totalRows = len(sdf.axes[0])


#firstRow = selectedColumns.loc[index,:]
#list1 = list(column1)
#column2 = df.loc[2,:]
#list2 = list(column2)


for i in range(0, totalRows):
    for j in range(0, totalRows):

        if list(sdf.loc[i,:]) == list(sdf.loc[j,:]):
            print(list(sdf.loc[i, :]))
            print(list(sdf.loc[j, :]))
            print("Number was found")
        else:
            print(list(sdf.loc[i, :]))
            print(list(sdf.loc[j, :]))
            print("Number was not found")

#if set(list1) & set(list2):
 #   print("Number was found")
#else:
 #   print("Number not in list")

# Declare a list that is to be converted into a column
#address = ['Delhi', 'Bangalore', 'Chennai', 'Patna']

# Using 'Address' as the column name
# and equating it to the list
#df['Address'] = address
