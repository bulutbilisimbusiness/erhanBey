import pandas as pd
import numpy as np
import random
import string
def okuma():
    df= pd.read_html("data/Mart Ayı Giriş Çıkış.xls")
    #print(df[1])
    df1 = pd.DataFrame(df[1])
    
    df1.columns = df1.columns.get_level_values(0)
    #print(df1.columns)
    #print(df1.Durumu)

    df = df1[df1.Durumu != 'Devamsız']
    df = df.reset_index(drop=True)
    print(df)
    df.to_excel("Output.xlsx")
okuma()    