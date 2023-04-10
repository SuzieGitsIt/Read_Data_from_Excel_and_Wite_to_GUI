# pip install xlrd
# pip install pandas
# pip install openpyxl
# https://www.geeksforgeeks.org/reading-excel-file-using-python/
# https://www.geeksforgeeks.org/creating-a-dataframe-using-excel-files/

# import pandas lib as pd
import pandas as pd         # librarie for reading from an excel file
import tkinter

req_col = [0,3]

# read by default 1st sheet of an excel file
Sheet1 = pd.read_excel('data_entry.xlsx')                                           # This is Sheet1, no need to call out the sheet_name.
Sheet2 = pd.read_excel('data_entry.xlsx', sheet_name= 1)                            # This is Sheet2
Sheet3 = pd.read_excel('data_entry.xlsx', sheet_name= 2, usecols= req_col)          # This is Sheet3, required columns A and D only [0,3]
Sheet4 = pd.read_excel('data_entry.xlsx', na_values= "Missing", sheet_name= 3)      # This is Sheet4, Missing in blank fields. Failed
Sheet5 = pd.read_excel('data_entry.xlsx', sheet_name= 4, skiprows= 2)               # This is Sheet5, Skip first 2 rows.
Sheet6 = pd.read_excel('data_entry.xlsx', sheet_name= 5, header= 2)                 # This is Sheet6, skip first 2 header columns
Sheet7 = pd.read_excel('data_entry.xlsx', na_values= "Missing", sheet_name= [5,6])  # This is Sheet6 and 7, Missing in blank fields. Failed
all_sheets= pd.read_excel('data_entry.xlsx', na_values= "Missing", sheet_name= None)# Open all worksheets 

print(Sheet1)
print(Sheet2)
print(Sheet3)
print(Sheet4)
print(Sheet5)
print(Sheet6)
print(Sheet7)
print(all_sheets)
