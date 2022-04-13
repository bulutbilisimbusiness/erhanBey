
import pandas as pd
import numpy as np
import random
import string
import openpyxl
from yeni import *

def perkotek() -> str:
    #columns = ['Personel Adı','Kart','Sicil No','Soyadı']
    #df = pd.read_excel('data/Mart Ayı Giriş Çıkış.xlsx', index_col=False,sheet_name='Mart Ayı Giriş Çıkış',usecols=[0,1,2,3,4,5,6,7,8,9,10,11],skiprows=[1,2,3],keep_default_na=False)
    #print(df)
    df= pd.read_html("data/Mart Ayı Giriş Çıkış.xls")
    #print(df[1])
    df1 = pd.DataFrame(df[1])
    
    df1.columns = df1.columns.get_level_values(0)
    #print(df1.columns)
    #print(df1.Durumu)

    df = df1[(df1.Durumu != 'Devamsız') & (df1.Durumu != "HAFTA TATİLİ")]
    print(df.isnull().sum().sort_values(ascending=False))
    df = df[~df['Sicil No'].isnull()]
    df = df.reset_index(drop=True)
    print('Eksik hücre sayısı')
    print(df.isnull().sum())
    df = df.sort_values(by=['Personel Adı'])
    print("personel aına göre sırala")
    print(df)
    df.to_excel("Output.xlsx")
    
    
perkotek()

 