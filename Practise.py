import pandas as pd
import numpy as np
import os

df = pd.read_excel(r'C:\Users\pasha\Desktop\untitled1\Book1.xlsx')
UPC = pd.read_excel(r'C:\Users\pasha\Desktop\untitled1\UPC_Template.xlsx')
# Analyst = int(input("What is your Byr#? "))


DC_list = []
PO_list = []
Name_list = []

for index, row in df.iterrows():
    if row['Analyst'] == 17:
        if row['NAME'] == "Gildan Apparel Canada LP" or row['NAME'] == "GILDAN APPAREL (OTC)(LB)" or row[
            'NAME'] == "GILDAN APPAREL(LB)":
            if row['Past Due'] == True:
                DC_list.append(row['DC'])
                PO_list.append(row['PO'])
                Name_list.append(row['NAME'])
df2 = pd.DataFrame({'DC': DC_list, 'PO': PO_list, 'Vendor Name': Name_list})

df2.sort_values('PO', ascending=False)
df3 = df2.drop_duplicates(subset='PO', keep='first')
print(df3)

'''Pivot Table
x = df[(df["VEND#"] == 3003999) | (df["VEND#"] == 20661201) | (df["VEND#"] == 20662201)]
y = x.drop_duplicates("PO#")
z = df.pivot_table(values="PO", index="NAME", aggfunc=lambda x: (x.unique()))
print(z)
df2 = df.pivot_table(index=["Analyst", "NAME"], values="PO", columns="Past Due", margins=True, aggfunc=lambda x: len(x.unique()))
t = df2.loc[[17], ["Gildan Apparel Canada LP", "GILDAN APPAREL (OTC)(LB)", "GILDAN APPAREL(LB)"], :]
print(t)
'''


''' UPC Finder'''
df = pd.read_excel(r'C:\Users\pasha\Desktop\untitled1\Book1.xlsx')
UPC = pd.read_excel(r'C:\Users\pasha\Desktop\untitled1\UPC_Template.xlsx')

DC_list = []
PO_list = []
Name_list = []

dc_for_upc = []
upc_to_find = []
concat = []
temp = []

for index, row in UPC.iterrows():
    dc_for_upc.append(row['dc'])
    upc_to_find.append(row['Sku#'])
    concat.append(str(row['dc']) + str(row['Sku#']))

df2 = pd.DataFrame({'DC': dc_for_upc, 'UPC': upc_to_find, 'Concat': concat})
shorts = df2.drop_duplicates()
#print(shorts)
use = df[['Sku#', 'DC', 'PO', 'DueDate', 'ApptDate', 'EventCode', 'Past Due', 'Description', 'POQty', 'QtyRecd']]
#print(use)
for index, row in use.iterrows():
    dc = str(row['DC'])
    sku = str(row['Sku#'])
    y = dc + sku
    temp.append(y)

use['Concat'] = temp
#print(use)
# add how for joins indicator=True
answer = pd.merge(use, shorts, on='Concat')

answer.to_excel (r'C:\Users\pasha\Desktop\untitled1\export_dataframe.xlsx', index = False, header=True)

'''Shorts'''

import pandas as pd
import glob as glob
import os
from datetime import datetime

os.chdir('C:/Users/pasha/Desktop/untitled1/Shorts')
fileNames = glob.glob('*.xlsx')
Mass = {}

for f in fileNames:
    x = pd.read_excel(f)
    x['Concat'] = x['DC'].astype(str) + x['upc'].astype(str)
    file_name = os.path.splitext(f)[0]
    date_string = '/'.join(file_name.split(" ")[1::])
    parsed_date = datetime.strptime(date_string, "%m/%d/%Y")
    x['TimeStamp'] = parsed_date
    parsed_date = parsed_date.weekday()

    if parsed_date == 0:
        x['Day'] = 'Monday'
    elif parsed_date == 1:
        x['Day'] = 'Tuesday'
    elif parsed_date == 2:
        x['Day'] = 'Wednesday'
    elif parsed_date == 3:
        x['Day'] = 'Thursday'
    elif parsed_date == 4:
        x['Day'] = 'Friday'

    x.set_index('Concat', drop=True, inplace=True)

    Mass[f] = x

for x in Mass:
    print(x)

print(Mass['Shorts 06 01 2020.xlsx'])
# x = pd.read_excel('Shorts 06 01 2020.xlsx')
# x['Concat'] = x['DC'].astype(str) + x['upc'].astype(str)
# file_name = os.path.splitext('Shorts 06 01 2020.xlsx')[0]
# date_string = '/'.join(file_name.split(" ")[1::])
# parsed_date = datetime.strptime(date_string, "%m/%d/%Y")
# parsed_date = parsed_date.weekday()
#
# if parsed_date == 0:
#     x['Day'] = 'Monday'
# elif parsed_date == 1:
#     x['Day'] = 'Tuesday'
# elif parsed_date == 2:
#     x['Day'] = 'Wednesday'
# elif parsed_date == 3:
#     x['Day'] = 'Thursday'
# elif parsed_date == 4:
#     x['Day'] = 'Friday'


# for f in fileNames:
#     x = pd.read_excel(f)
#
#
# # dataframes = [pd.read_excel(f) for f in fileNames]
#
# for f in fileNames:
#     file_name = os.path.splitext(f)[0]
#     date_string = '/'.join(file_name.split(" ")[1::])
#     parsed_date = datetime.strptime(date_string, "%m/%d/%Y")
#     print(parsed_date.weekday())
#
#
# print(dataframes)
