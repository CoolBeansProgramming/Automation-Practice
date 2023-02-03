import pandas as pd
from openpyxl import load_workbook


df=pd.read_excel(r'sales.xlsx', sheet_name='Sheet1')


df=df[['Gender','Product line','Total']]

pivot_table = df.pivot_table(index='Gender', columns='Product line', values='Total', aggfunc='sum').round(0)

pivot_table.to_excel('pivot_table.xlsx', 'Report', startrow=4)
