from flask import Flask,send_from_directory
import pandas as pd
import numpy as np
import random
import string
from flask_cors import CORS, cross_origin
from yeni import *
app = Flask(__name__)


cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'
@app.route("/")
def hello() -> str:
    return "Hello World"


@cross_origin()
@app.route("/excel")
def excel() -> str:

    df = pd.read_excel('data/TimeReport_TEAM.xlsx', index_col=False)
    
    izinler = pd.read_excel('data/Ağustos izin.XLSX')

    resmi_tatil = pd.read_excel('data/resmi_tatil_2021.xlsx', index_col=False)
    df = df[~df['Tarih'].isin(resmi_tatil['Tarih'])]
    
    df.to_excel('data/modified.xlsx', columns=['Tarih', 'Çalışan Adı', 'Giriş Saati', 'Çıkış Saati','Çalışılan Saatler'], index=False)
    df = pd.read_excel('data/modified.xlsx', index_col=False)
    
    # Aynı tarihte bir çalışan iki kere giriş ve çıkış yapmış.
    veri = df.duplicated(['Tarih', 'Çalışan Adı'])
    if [veri == 'True']:
        df.drop_duplicates(subset=['Tarih', 'Çalışan Adı'], keep='first', inplace=True)
    else:
        pass

    # Çalışan isimlerinin düzenlenmesi
    df['Çalışan Adı'] = df['Çalışan Adı'].str.strip()
    df_name = pd.read_excel('data/Arge İsim.xlsx', index_col=False)
    names = df_name.iloc[:,0].values
    new_name = df_name.iloc[:,1].values
    old_name = df['Çalışan Adı'].values

    for i in range(len(old_name)):
        for n in range(len(names)):
            if old_name[i] == names[n]:
                old_name[i] = new_name[n]
            else:
                pass  
    df['Çalışan Adı'] = old_name     
        
    df['Tarih'] = pd.to_datetime(df['Tarih']).dt.strftime('%Y-%m-%d')
    df['Giriş Saati'] = pd.to_datetime(df['Giriş Saati']).dt.strftime('%H:%M:%S')
    df['Çıkış Saati'] = pd.to_datetime(df['Çıkış Saati']).dt.strftime('%H:%M:%S')


    # Giriş saati 09:00 olanlara rastgele değer atama
    df.loc[df['Giriş Saati'] == '09:00:00', 'Giriş Saati'] = df.apply(lambda row : random_giris_saati_ekle(), axis = 1)
    
    # Çıkış saati 18:00 olanlara rastgele değer atama
    df.loc[df['Çıkış Saati'] == '18:00:00', 'Çıkış Saati'] = df.apply(lambda row : random_cikis_saati_ekle(), axis = 1)
    
    calisilan_saat_hesapla(df)

    df['Çıkış Saati'] =np.where((df['Çalışılan Saatler'] < 8.75) & (df['Giriş Saati'] < '08:30:00'), df.apply(lambda row : random_cikis_saati_ekle(),axis = 1), df['Çıkış Saati'])
    df.loc[df['Çalışılan Saatler'] > 11, 'Çıkış Saati'] = df.apply(lambda row: random_cikis_saati_ekle(), axis=1)
    
    calisilan_saat_hesapla(df)

    # Giriş Yapmayan Personel Var Mı?
    try:
        df['İzin Türü'] = 'Yok'

        teams = df['Çalışan Adı'].unique().tolist()

        arge_personelleri = pd.read_excel('data/Ar-Ge Personelleri.xlsx', dtype={'Durumu': str})
        arge_personelleri = arge_personelleri[['Kişiler', 'İşe Başlama Tarihi', 'Durumu']]
        arge_personelleri = arge_personelleri.loc[arge_personelleri['Durumu'] == 'Çalışıyor\xa0']
        arge_personelleri['Kişiler'] = arge_personelleri['Kişiler'].apply(lambda x: string.capwords(x))
        arge_personelleri['İşe Başlama Tarihi'] = pd.to_datetime(arge_personelleri['İşe Başlama Tarihi']).dt.strftime('%Y-%m-%d')
        arge_personelleri = arge_personelleri.reset_index(drop=True)
        arge_personelleri = arge_personelleri.rename(columns={'Kişiler': "Çalışan Adı"}) # Kolon ismi değiştirme

        arge_personelleri_list = arge_personelleri['Çalışan Adı'].unique().tolist()

        not_arge = []
        for i in teams:
            if i not in arge_personelleri_list:
                not_arge.append(i)
                print(i)
        df = df[~df['Çalışan Adı'].isin(not_arge)]

        izinler = izinler[['Tarih','Çalışanın/başvuranın adı', 'İzin Türü']]
        izinler = izinler.rename(columns={'Çalışanın/başvuranın adı': "Çalışan Adı"}) # Kolon ismi değiştirme
        izinler['Tarih'] = pd.to_datetime(izinler['Tarih']).dt.strftime('%Y-%m-%d') 

        izinli_list = izinler['Çalışan Adı'].unique().tolist()

        not_arge = []
        for i in izinli_list:
            if i not in arge_personelleri_list:
                not_arge.append(i)
        izinler = izinler[~izinler['Çalışan Adı'].isin(not_arge)]

        table = pd.merge(df,arge_personelleri , on="Çalışan Adı", how="left")
        table2 = pd.merge(table, izinler , on=['Çalışan Adı','Tarih'], how="left")

        table2.loc[table2['İzin Türü_y'].isnull(), 'İzin Türü_y'] = 'Yok'
        table2 = table2.drop(['İzin Türü_x'],axis=1)

        tarihler = sorted(list(set(list(table2['Tarih'].values))))
        df_tarih = pd.DataFrame(columns=['Tarih'])
        df_tarih['Tarih'] = tarihler

        personel = sorted(list(set(list(df['Çalışan Adı'].values))))
        # tarihler = sorted(list(set(list(table2['Tarih'].values))))
        personels = pd.DataFrame(columns=['Çalışan Adı'])
        personels['Çalışan Adı'] = personel

        df_tarih['tmp'] = 1
        personels['tmp'] = 1

        table_null = pd.merge(df_tarih, personels,on='tmp')
        table_null = table_null.drop('tmp', axis=1)

        table3 = pd.merge(table_null, table2, on=['Çalışan Adı','Tarih'], how="left")
        table4 = pd.merge(table3, izinler , on=['Çalışan Adı','Tarih'], how="left")
        table5 = pd.merge(table4, arge_personelleri, on=['Çalışan Adı'], how="left")

        table5 = table5.drop(['İşe Başlama Tarihi_x','Durumu_x','İzin Türü_y'], axis=1)
        table5.loc[table5['İzin Türü'].isnull(), 'İzin Türü'] = 'Çalışma'

        all_data = table5.loc[(table5['Tarih'] >= table5['İşe Başlama Tarihi_y'])]

        # hafta sonu bulma
        all_data['Tarih'] = pd.to_datetime(all_data['Tarih'])
        all_data['haftasonu'] = all_data['Tarih'].dt.day_name().isin(['Saturday', 'Sunday'])

        haftasonu = all_data.loc[(all_data['haftasonu'] == True) & (all_data['Giriş Saati'].isnull())]
        all_data = all_data[~(all_data.isin(haftasonu))]

        all_data['haftasonu'] = all_data['Tarih'].dt.day_name().isin(['Sunday'])
        all_data = all_data[all_data['haftasonu'] != True]

        all_data['Tarih'] = pd.to_datetime(all_data['Tarih']).dt.strftime('%Y-%m-%d')
        all_data['İşe Başlama Tarihi'] = pd.to_datetime(all_data['İşe Başlama Tarihi_y']).dt.strftime('%Y-%m-%d')
        all_data['Giriş Saati'] = pd.to_datetime(all_data['Giriş Saati']).dt.strftime('%H:%M:%S')
        all_data['Çıkış Saati'] = pd.to_datetime(all_data['Çıkış Saati']).dt.strftime('%H:%M:%S')
    except:
        pass

    # izin kuralları
    all_data = all_data[all_data['İzin Türü'] != 'Mazeret']

    all_data.loc[all_data['Giriş Saati'].isnull(), 'Giriş Saati'] = all_data.apply(lambda row : random_giris_saati_ekle(), axis = 1)
    all_data.loc[all_data['Çıkış Saati'].isnull(), 'Çıkış Saati'] = all_data.apply(lambda row : random_cikis_saati_ekle(), axis = 1)

    all_data.loc[all_data['İzin Türü'] == 'Yıllık İzin', 'Giriş Saati'] = '09:00:00'
    all_data.loc[all_data['İzin Türü'] == 'Yıllık İzin', 'Çıkış Saati'] = '18:00:00'

    all_data.loc[all_data['İzin Türü'] == 'Rapor', 'Giriş Saati'] = '09:00:00'
    all_data.loc[all_data['İzin Türü'] == 'Rapor', 'Çıkış Saati'] = '18:00:00'
    
    calisilan_saat_hesapla(all_data)
    
    # Bir yılı dolmamış kişiler yıllık izin kullandıysa dahil etme
    all_data['Tarih'] = pd.to_datetime(all_data['Tarih'])
    all_data['İşe Başlama Tarihi'] = pd.to_datetime(all_data['İşe Başlama Tarihi'])

    all_data['gün'] = (all_data['Tarih'] - all_data['İşe Başlama Tarihi']).dt.days 

    all_data =  all_data.loc[~((all_data['İzin Türü'] == 'Yıllık İzin' ) & (all_data['gün'] < 365))]

    all_data['Tarih'] = pd.to_datetime(all_data['Tarih']).dt.strftime('%Y-%m-%d')        

    all_data['Giriş Saati'] = np.where((all_data['Çalışılan Saatler'] < 8.75) & (all_data['İzin Türü'] != 'Yarım Gün İzin'), all_data.apply(lambda row : random_giris_saati_ekle(),axis = 1), all_data['Giriş Saati'])
    
    calisilan_saat_hesapla(all_data)
    
    all_data['Çıkış Saati'] = np.where((all_data['Çalışılan Saatler'] < 8.75) & (all_data['İzin Türü'] != 'Yarım Gün İzin'), all_data.apply(lambda row : random_cikis_saati_ekle(),axis = 1), all_data['Çıkış Saati'])
    all_data.loc[all_data['Çalışılan Saatler'] > 11, 'Çıkış Saati'] = all_data.apply(lambda row : random_cikis_saati_azalt(),axis = 1)

    calisilan_saat_hesapla(all_data)

    all_data['Çıkış Saati'] = np.where((all_data['Çalışılan Saatler'] < 8.75) & (all_data['İzin Türü'] != 'Yarım Gün İzin'), all_data.apply(lambda row : cikis_saati_arttir(),axis = 1), all_data['Çıkış Saati'])
    all_data['Çıkış Saati'] = np.where((all_data['Çalışılan Saatler'] > 11) & (all_data['Giriş Saati'] < '08:30:00'), all_data.apply(lambda row : cikis_saati_azalt(),axis = 1), all_data['Çıkış Saati'])

    # Yarım Gün İzinli Kuralları
    all_data['Giriş Saati'] = pd.to_datetime(all_data['Giriş Saati'])
    all_data['Çıkış Saati'] = pd.to_datetime(all_data['Çıkış Saati'])
    all_data.loc[all_data['İzin Türü'] == 'Yarım Gün İzin', 'Çıkış Saati'] = all_data['Giriş Saati'] + pd.to_timedelta('04:30:00')

    veri1 = all_data.loc[(all_data['İzin Türü'] == 'Yarım Gün İzin') & (all_data['Giriş Saati'] < '13:00:00')]
    veri1['Giriş Saati'] = veri1['Çıkış Saati']
    veri1['Çıkış Saati'] = veri1['Giriş Saati'] + pd.to_timedelta('04:30:00')
    veri1['İzin Türü'] = np.nan

    all_data = all_data.append(veri1, ignore_index=False)
    
    veri2 = all_data.loc[(all_data['İzin Türü'] == 'Yarım Gün İzin') & (all_data['Giriş Saati'] >= '13:00:00')]
    veri2['Çıkış Saati'] = veri2['Giriş Saati']
    veri2['Giriş Saati'] = veri2['Çıkış Saati'] - pd.to_timedelta('04:30:00')
    veri2['İzin Türü'] = np.nan

    all_data = all_data.append(veri2, ignore_index=False)

    all_data.loc[all_data['İzin Türü'] == 'Yarım Gün İzin', 'İzin Türü'] = 'Çalışma'
    all_data.loc[all_data['İzin Türü'].isnull(), 'İzin Türü'] = 'Yarım Gün İzin'

    all_data = all_data.sort_values(by='Tarih')
    
    all_data['Giriş Saati'] = pd.to_datetime(all_data['Giriş Saati']).dt.strftime('%H:%M:%S')
    all_data['Çıkış Saati'] = pd.to_datetime(all_data['Çıkış Saati']).dt.strftime('%H:%M:%S')

    calisilan_saat_hesapla(all_data)

    all_data['Tarih'] = pd.to_datetime(all_data['Tarih']).dt.strftime('%Y-%m-%d')
    all_data['Tarih'] = pd.to_datetime(all_data['Tarih']).dt.date

    # Boş satırları dahil etme
    all_data = all_data[~all_data['Giriş Saati'].isnull()]

    all_data = all_data.rename(columns={'İzin Türü': 'Açıklama'})

    all_data['Çalışılan Saatler'] = all_data['Çalışılan Saatler'].apply(lambda x:round(x,2))

    all_data.to_excel('data/modified.xlsx', columns=['Tarih', 'Çalışan Adı', 'Giriş Saati', 'Çıkış Saati', 'Çalışılan Saatler', 'Açıklama'], sheet_name='Sayfa1', index=False)
    all_data = all_data.groupby(['Tarih','Çalışan Adı',]).sum()

    pivot_tbl = pd.pivot_table(all_data, index='Çalışan Adı', columns='Tarih', values='Çalışılan Saatler')
    with pd.ExcelWriter('data/modified.xlsx', engine='openpyxl', mode='a') as writer:
        pivot_tbl.to_excel(writer, sheet_name='Pivot_tablo')
    
    #return_df = pd.read_excel('data/modified.xlsx', index_col=False)
    #return_df = pd.read_excel('data/Ar-Ge Personelleri.xlsx', index_col=False)
    return_df=pd.read_excel('data/Mart Ayı Giriş Çıkış.xls',index_col=False)
    return return_df.to_json(orient='table')
    #return send_from_directory(path='modified.xlsx',directory='data', as_attachment=False)
    



if __name__ == "__main__":
    app.run(debug=True, port=3000)





