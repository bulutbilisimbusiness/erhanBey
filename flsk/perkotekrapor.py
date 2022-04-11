
import pandas as pd
import numpy as np
import random
import string
import openpyxl
from yeni import *

def perkotek() -> str:
    columns = ['Personel Adı','Kart','Sicil No','Soyadı']
    df = pd.read_excel('data/Mart Ayı Giriş Çıkış.xlsx', index_col=False,sheet_name='Mart Ayı Giriş Çıkış',header=None,usecols=[0,1,2,3,4,5,6,7,8,9,10,11],skiprows=[1,2,3],keep_default_na=False)

    print(df)
    df = df[~df(columns[10]).astype(str).str.startswith('Devamsız')]
    df = df.drop(labels='Devamsız', axis=0)
    data_with_index = data_with_index.drop('Devamsız')
    df=data_with_index
    df.to_excel("Output.xlsx")
    df.head()
    wb_obj = openpyxl.load_workbook('data/Mart Ayı Giriş Çıkış.xlsx')
    sheet_obj = wb_obj.active
    row = sheet_obj.max_row
    column = sheet_obj.max_column
    m_row = sheet_obj.max_row
    print("Total Rows:", row)
    print("Total Columns:", column)
    print("\nValue of first column")
    """ for i in range(6731, row + 1): 
        cell_obj = sheet_obj.cell(row = i, column = 1) 
        print(cell_obj.value) 
    print("\nValue of first row")
    for i in range(11, m_row + 1):
        cell_obj = sheet_obj.cell(row = i, column = 1)
        print(cell_obj.value)   """ 
    df.info()
    
    df.columns

    for x in df:
        print(x)
    #print(df['Kart','Sicil No','Personel Adı'].tolist()) 
    #print('Excel Sheet to JSON:', df.to_json(orient='records')) 
    dframe = pd.DataFrame(df, columns= ['Personel Adı','Kart','Sicil No'])
    print(dframe) 
perkotek()

 