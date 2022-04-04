from flask import Flask
import pandas as pd
import numpy as np
import random
import string
from flask_cors import CORS, cross_origin
def random_giris_saati_ekle():
    giris_saat = random.randint(8, 9)
    if giris_saat == 8:
        giris_dakika = random.randint(45, 59) 
    else:
        giris_dakika = random.randint(0, 14)    

    giris_saati = str(giris_saat) + ':' + str(giris_dakika) + ':' + '00'
    return giris_saati

def random_cikis_saati_ekle():
    cikis_saat = random.randint(17, 18)
    if cikis_saat == 17:
        cikis_dakika = random.randint(50, 59)
    else:
        cikis_dakika = random.randint(0, 45)
        
    cikis_saati = str(cikis_saat) + ':' + str(cikis_dakika) + ':' + '00'
    return cikis_saati

def random_cikis_saati_azalt():
    cikis_saat = random.randint(18, 19)
    if cikis_saat == 18:
        cikis_dakika = random.randint(0, 59)
    else:
        cikis_dakika = random.randint(0, 30)
        
    cikis_saati = str(cikis_saat) + ':' + str(cikis_dakika) + ':' + '00'
    return cikis_saati

def cikis_saati_arttir():
    cikis_saat = 18
    cikis_dakika = random.randint(15, 59)        
    cikis_saati = str(cikis_saat) + ':' + str(cikis_dakika) + ':' + '00'
    return cikis_saati

def cikis_saati_azalt():
    cikis_saat = 17
    cikis_dakika = random.randint(50, 56)        
    cikis_saati = str(cikis_saat) + ':' + str(cikis_dakika) + ':' + '00'
    return cikis_saati

def calisilan_saat_hesapla(all_data):
    all_data['Giriş Saati'] = all_data['Tarih'] +' '+ all_data['Giriş Saati']
    all_data['Çıkış Saati'] = all_data['Tarih'] +' '+ all_data['Çıkış Saati']

    all_data['Giriş Saati'] = pd.to_datetime(all_data['Giriş Saati'])
    all_data['Çıkış Saati'] = pd.to_datetime(all_data['Çıkış Saati'])

    all_data['Çalışılan Saatler'] = (all_data['Çıkış Saati'] - all_data['Giriş Saati']) / pd.Timedelta(hours=1)

    all_data['Giriş Saati'] = pd.to_datetime(all_data['Giriş Saati']).dt.strftime('%H:%M:%S')
    all_data['Çıkış Saati'] = pd.to_datetime(all_data['Çıkış Saati']).dt.strftime('%H:%M:%S')
    return(all_data)