import pandas as pd
import openpyxl

# File Path
inputFile = '/Users/arvindkumar/Documents/JCI/Input/Book1.xlsx'
outputFile = '/Users/arvindkumar/Documents/JCI/Output/OutputFile.xlsx'

sheetName = 'Sheet1'

# Fetching the excel file, Hence please mention the proper excel file path and sheet
df = pd.read_excel(inputFile, sheetname=sheetName)
wb = openpyxl.load_workbook(inputFile)
sheet1 = wb.get_sheet_by_name(sheetName)

# Selecting the columns which are required to be compared
sdf= df[['Column1', 'Column2', 'Column3','Column4','Column5','Column6','Column7','Column8','Column9','Column10']]

# Finding number of rows and columns
totalRows = len(sdf.axes[0])
totalcols = len(df.axes[1])

colIndex = totalcols + 1

sheet1.cell(row=1, column=colIndex).value = 'Match Found'

for i in range(0, totalRows):

    for j in range(i+1, totalRows):

        if list(sdf.loc[i,:]) == list(sdf.loc[j,:]):
            if sheet1.cell(row=i + 2, column=colIndex).value is None or  sheet1.cell(row=j+2, column=colIndex).value is None:
                sheet1.cell(row=i+2, column=colIndex).value = i +1
                sheet1.cell(row=j+2, column=colIndex).value = i + 1

# Saving excel output file, provide path for save the output file
wb.save(outputFile)
