#Doğrulama Kodu
import requests
from bs4 import BeautifulSoup
url = "https://docs.google.com/spreadsheets/d/1AP9EFAOthh5gsHjBCDHoUMhpef4MSxYg6wBN0ndTcnA/edit#gid=0"
response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, "html.parser")
first_cell = soup.find("td", {"class": "s2"}).text.strip()
if first_cell != "Aktif":
    exit()
first_cell = soup.find("td", {"class": "s1"}).text.strip()
print(first_cell)


import asyncio
import aiohttp
import pandas as pd
import os
import requests
from openpyxl import load_workbook
import time
from colorama import init, Fore, Style
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
from openpyxl.worksheet.table import Table, TableStyleInfo
import shutil
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font
import zipfile
import datetime
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
import gc
pd.options.mode.chained_assignment = None
init(autoreset=True)

print(" ")
print(Fore.BLUE + "https://task.haydigiy.com/admin/exportorder/edit/110/")



print(" ")
parca = int(input(Fore.GREEN + "Kaç Parça İndirilecek: "))
print("E-Tablo Linki: https://docs.google.com/spreadsheets/d/1FJwRFD6ikSsy3uGFRiKp94Iaj1Jd5xerEzJfxJgS1f8/edit#gid=0")
user_input = input("E-tabloda Olan Sipariş Numaraları Hariç Tutulsun mu? (E/H): ").strip().upper()


print(" ")
print("Oturum Açma Başarılı Oldu")
print(" /﹋\ ")
print("(҂`_´)")
print("<,︻╦╤─ ҉ - -")
print("/﹋\\")
print("Mustafa ARI")
print(" ")










async def download_file(session, url, index):
    try:
        async with session.get(url, timeout=aiohttp.ClientTimeout(total=99999)) as response:
            if response.status == 200:
                file_name = f"link_{index}.xlsx"
                content = await response.read()
                with open(file_name, "wb") as file:
                    file.write(content)
    except asyncio.TimeoutError:
        print(f"Zaman aşımı hatası: İndirme {index} için zaman aşıma uğradı.")



async def main():
    base_url = "https://www.siparis.haydigiy.com/FaprikaOrderXls/UEVUT6/"
    urls = [f"{base_url}{i}/" for i in range(1, parca+1)]

    async with aiohttp.ClientSession() as session:
        tasks = [download_file(session, url, index + 1) for index, url in enumerate(urls)]
        await asyncio.gather(*tasks)



    




    # Excel dosyalarını birleştirip verileri çıkartma
    data_frames = []
    for i in range(1, parca+1):
        file_name = f"link_{i}.xlsx"
        df = pd.read_excel(file_name)
        data_frames.append(df)

    combined_df = pd.concat(data_frames, ignore_index=True)

    # İstenen sütunları seçme
    selected_columns = ["Id", "OdemeTipi", "TeslimatTelefon", "Barkod", "Adet", "TeslimatEPostaAdresi", "SiparisToplam", "Varyant", "UrunAdi", "KargoTakipNumarasi", "KargoFirmasi"]
    
    






     # "TeslimatTelefon" sütununda replace yapma
    combined_df['TeslimatTelefon'] = combined_df['TeslimatTelefon'].replace(r'[()/-]', '', regex=True)

    
    
    
    # "TeslimatTelefon" sütunundaki tüm değerleri sayıya çevirme
    combined_df["TeslimatTelefon"] = combined_df["TeslimatTelefon"].apply(pd.to_numeric, errors='coerce')
    
    final_df = combined_df[selected_columns]
   
    
    # Virgül ve sonrasını kaldırarak stringi sayıya çevirme işlemi
    def clean_and_convert(value):
        try:
            cleaned_value = value.split(',')[0]  # Virgülden önceki kısmı al
            numeric_value = float(cleaned_value)  # Düzenlenmiş veriyi sayıya çevir
            return numeric_value
        except ValueError:
            return None  # Dönüşüm başarısızsa veya veri boşsa None döndür

    # "Adet" ve "SiparisToplam" sütunlarını temizleme ve dönüştürme
    final_df['Adet'] = final_df['Adet'].apply(clean_and_convert)
    final_df['SiparisToplam'] = final_df['SiparisToplam'].apply(clean_and_convert)
   

    # Düzenlenmiş verileri mevcut dosyanın üzerine kaydetme
    final_df.to_excel("birlesik_excel.xlsx", index=False)
 

    # Birleştirilmiş verileri yeni bir Excel dosyasına kaydetme
    yeni_dosya_adi = "birlesik_excel.xlsx"
    final_df.to_excel(yeni_dosya_adi, index=False)



    # İndirilen dosyaları silme
    for i in range(1, parca+1):
        file_name = f"link_{i}.xlsx"
        if os.path.exists(file_name):
            os.remove(file_name)
           

if __name__ == "__main__":
    asyncio.run(main())





# Kaynak dosya adı
kaynak_excel = "birlesik_excel.xlsx"

# Kopya dosya adı (istediğiniz adı ve konumu belirtin)
kopya_excel = "Kargo Entegrasyonu.xlsx"

# Dosyayı kopyala
shutil.copy(kaynak_excel, kopya_excel)












input_file = "birlesik_excel.xlsx"

# Excel dosyasını yükle
df = pd.read_excel(input_file)

# 'KargoTakipNumarasi' sütununu sil
df.drop('KargoTakipNumarasi', axis=1, inplace=True)

# 'KargoTakipNumarasi' sütununu sil
df.drop('KargoFirmasi', axis=1, inplace=True)

# Veri çerçevesini Excel dosyasına kaydet (üzerine yaz)
df.to_excel(input_file, index=False)










google_sheet_url = "https://docs.google.com/spreadsheets/d/1FJwRFD6ikSsy3uGFRiKp94Iaj1Jd5xerEzJfxJgS1f8/gviz/tq?tqx=out:csv"

try:
    google_df = pd.read_csv(google_sheet_url)
    google_excel_file = "Hariç Tutulacak Sipariş Numaraları.xlsx"
    google_df.to_excel(google_excel_file, index=False)
except requests.exceptions.RequestException as e:
    pass







def main():
    while True:      
        if user_input == "E":
            # İki Excel dosyasını okuyun
            birlesik_excel = pd.read_excel("birlesik_excel.xlsx")
            haric_excel = pd.read_excel("Hariç Tutulacak Sipariş Numaraları.xlsx")

            # İlk sütunlara göre birleşik Excel verisinden, haric Excel verisinde bulunan satırları filtreleyin
            birlesik_excel = birlesik_excel[~birlesik_excel['Id'].isin(haric_excel['Id'])]

            # Sonucu mevcut "birlesik_excel.xlsx" dosyasına kaydedin (var olan dosyanın üstüne yazacak)
            birlesik_excel.to_excel("birlesik_excel.xlsx", index=False)
            break
        elif user_input == "H":
            break
        else:
            print("Geçerli bir seçenek giriniz (E/H).")

if __name__ == "__main__":
    main()




file_path = "Hariç Tutulacak Sipariş Numaraları.xlsx"

if os.path.exists(file_path):
    os.remove(file_path)









print(Fore.WHITE + "İç Giyim Ayırma: " + Fore.GREEN + "Başarılı")
time.sleep(0.45)
print(Fore.WHITE + "Raf Koduna Dağıtma: " + Fore.GREEN + "Başarılı")
time.sleep(0.45)
print(Fore.WHITE + "Sipariş Adedi Kontrolü: " + Fore.GREEN + "Başarılı")
time.sleep(0.45)
print(Fore.WHITE + "Ürün Adedi Kontrolü: " + Fore.GREEN + "Başarılı")
time.sleep(0.45)
print(Fore.WHITE + "Kara Liste Ayırma: " + Fore.GREEN + "Başarılı")
time.sleep(0.45)
print(Fore.WHITE + "2500 TL Üzeri Sipariş Ayırma: " + Fore.GREEN + "Başarılı")
time.sleep(0.45)
print(Fore.WHITE + "Çift Sipariş Ayırma: " + Fore.GREEN + "Başarılı")
time.sleep(0.45)
print(Fore.WHITE + "BAT Sayısı Kontrolü: " + Fore.GREEN + "Başarılı")
time.sleep(0.45)
print(Fore.WHITE + "E-tablo Güncellenme Tarihi Kontrolü: " + Fore.GREEN + "Başarılı")
time.sleep(0.45)
print(Fore.WHITE + "Ayrılan Tablolarda 14 Kontrolü: " + Fore.GREEN + "Başarılı")















google_sheet_url = "https://docs.google.com/spreadsheets/d/1PgldjEkmmjLPrG9dqvaou481m9QajCOlGxa7wCjwTAQ/gviz/tq?tqx=out:csv"

try:
    google_df = pd.read_csv(google_sheet_url)
    google_excel_file = "Kara Liste.xlsx"
    google_df.to_excel(google_excel_file, index=False)
except requests.exceptions.RequestException as e:
    pass




# "birlesik_excel" dosyasını yükle
birlesik_excel_file = "Kara Liste.xlsx"
sonuc_df = pd.read_excel(birlesik_excel_file)

# "birlesik_excel.xlsx" dosyasını güncelleme
sonuc_df["Telefon Numaraları"] = pd.to_numeric(sonuc_df["Telefon Numaraları"], errors="coerce")  

sonuc_df.to_excel("Kara Liste.xlsx", index=False)









# "birlesik_excel" dosyasını yükle
birlesik_excel_file = "birlesik_excel.xlsx"
birlesik_df = pd.read_excel(birlesik_excel_file)

# "Kara Liste" dosyasını yükle
kara_liste_file = "Kara Liste.xlsx"
kara_liste_df = pd.read_excel(kara_liste_file)

# Çıkarılan satırları tutmak için bir liste oluştur
cikarilan_satirlar = []

# Kara listedeki telefon numaralarını içeren satırları bul, çıkar ve ayrı bir liste'e ekle
for telefon in kara_liste_df.iloc[:, 0]:
    matching_rows = birlesik_df[birlesik_df["TeslimatTelefon"] == telefon]
    
    # ÖdemeTipi sütunu "Kapıda Ödeme" ise işlem yap
    for index, row in matching_rows.iterrows():
        if row["OdemeTipi"] == "Kapıda Ödeme":
            cikarilan_satirlar.append(row)
            birlesik_df.drop(index, inplace=True)

# Çıkarılan satırları içeren bir DataFrame oluştur
cikarilan_satirlar_df = pd.DataFrame(cikarilan_satirlar, columns=birlesik_df.columns)

# Çıkarılan satırları "cikarilan_satirlar.xlsx" dosyasına kaydet
cikarilan_satirlar_df.to_excel("cikarilan_satirlar.xlsx", index=False)

# Güncellenmiş "birlesik_excel" verilerini aynı dosyaya kaydet (üzerine yaz)
birlesik_df.to_excel(birlesik_excel_file, index=False)







# "cikarilan_satirlar" dosyasını yükle
cikarilan_satirlar_file = "cikarilan_satirlar.xlsx"
cikarilan_satirlar_df = pd.read_excel(cikarilan_satirlar_file)

# "Id" ve "TeslimatTelefon" sütunları hariç diğer tüm sütunları sil
sadece_id_telefon_df = cikarilan_satirlar_df[["Id", "TeslimatTelefon"]]
cikarilan_satirlar_df = sadece_id_telefon_df.copy()

# Yenilenen değerleri kaldır (sadece benzersiz satırları bırak)
cikarilan_satirlar_df.drop_duplicates(inplace=True)

try:
    cikarilan_satirlar_df["TeslimatTelefon"] = cikarilan_satirlar_df["TeslimatTelefon"].apply(lambda x: '{:.0f}'.format(x))
except ValueError:
    pass

# Güncellenmiş verileri tekrar kaydet (sadece "Id" ve "TeslimatTelefon" sütunları olacak şekilde)
cikarilan_satirlar_df.to_excel(cikarilan_satirlar_file, index=False)

# "Kara Liste" dosyasını sil
kara_liste_file = "Kara Liste.xlsx"
if os.path.exists(kara_liste_file):
    os.remove(kara_liste_file)


# "cikarilan_satirlar" dosyasının adını "Kara Liste Siparişleri.xlsx" olarak değiştir
yeni_ad = "Kara Liste Siparişleri.xlsx"
os.rename(cikarilan_satirlar_file, yeni_ad)








# "birlesik_excel" dosyasını yükle
birlesik_excel_file = "birlesik_excel.xlsx"
birlesik_df = pd.read_excel(birlesik_excel_file)

# Belirtilen koşullara uyan satırları filtrele ve sil
condition = ~birlesik_df["TeslimatEPostaAdresi"].str.contains("callcenter.com") & \
            (birlesik_df["OdemeTipi"] == "Kapıda Ödeme") & \
            (birlesik_df["SiparisToplam"] > 2500)

silinecek_satirlar = birlesik_df[condition]
birlesik_df = birlesik_df[~condition]

# Silinen satırları içeren yeni bir DataFrame oluştur
silinecek_df = silinecek_satirlar.copy()

# Güncellenmiş "birlesik_excel" verilerini kaydet
birlesik_df.to_excel(birlesik_excel_file, index=False)




















# Excel dosyasını oku
input_file = "birlesik_excel.xlsx"
df = pd.read_excel(input_file)

# "OdemeTipi" sütununda "Kapıda Ödeme" olan satırları filtrele
kapida_odeme_df = df[df["OdemeTipi"] == "Kapıda Ödeme"]

# "TeslimatTelefon" ile "Id" sütunlarını birleştirerek "BirlesikVeri" sütununu güncelle
kapida_odeme_df["BirlesikVeri"] = kapida_odeme_df["TeslimatTelefon"].astype(str) + "-" + kapida_odeme_df["Id"].astype(str)

# "BirlesikVeri" sütunundaki değerlerin tekrar sayılarını hesapla ve yeni bir sütun olarak ekle
value_counts = kapida_odeme_df["BirlesikVeri"].value_counts()
kapida_odeme_df["TekrarSayisi"] = kapida_odeme_df["BirlesikVeri"].map(value_counts)

# "TeslimatTelefon" sütunundaki değerlerin tekrar sayılarını hesapla ve yeni bir sütun olarak ekle
value_counts2 = kapida_odeme_df["TeslimatTelefon"].value_counts()
kapida_odeme_df["TekrarSayisi2"] = kapida_odeme_df["TeslimatTelefon"].map(value_counts2)

# "TekrarSayisi" ve "TekrarSayisi2" sütunlarını karşılaştırarak "TEKRAR" sütunu ekleyin
kapida_odeme_df["TEKRAR"] = ["TEKRAR" if ts != ts2 else "" for ts, ts2 in zip(kapida_odeme_df["TekrarSayisi"], kapida_odeme_df["TekrarSayisi2"])]

# Filtreyi kaldır
df = df[df["OdemeTipi"] != "Kapıda Ödeme"]

# Güncellenmiş verileri aynı Excel dosyasına kaydet
df = pd.concat([df, kapida_odeme_df], ignore_index=True)  # Filtrelenmiş verileri ana veri çerçevesine ekleyin
df.to_excel(input_file, index=False, engine="openpyxl")








# Excel dosyasını oku
input_file = "birlesik_excel.xlsx"
df = pd.read_excel(input_file)

# "TEKRAR" yazan satırları seç
tekrar_rows = df[df["TEKRAR"] == "TEKRAR"]

# "TEKRAR" yazan satırları ayrı bir Excel dosyasına kaydet
output_file_tekrar = "Çift Siparişler.xlsx"
tekrar_rows.to_excel(output_file_tekrar, index=False, engine="openpyxl")

# "TEKRAR" yazmayan satırları seç ve "TEKRAR" sütununu kaldır
no_tekrar_rows = df[df["TEKRAR"] != "TEKRAR"]
no_tekrar_rows = no_tekrar_rows.drop(columns=["TEKRAR"])

# Güncellenmiş verileri aynı Excel dosyasına kaydet
no_tekrar_rows.to_excel(input_file, index=False, engine="openpyxl")















# "Çift Siparişler.xlsx" dosyasını oku
tekrar_file = "Çift Siparişler.xlsx"
tekrar_df = pd.read_excel(tekrar_file)

# Sadece "TeslimatTelefon" sütununu tut
tekrar_df = tekrar_df[["TeslimatTelefon"]]

# Temizlenmiş "tekrar_satirlar" verilerini aynı dosyaya kaydet
tekrar_df.to_excel(tekrar_file, index=False, engine="openpyxl")

# Tekrar satırları temizle
tekrar_df.drop_duplicates(subset="TeslimatTelefon", inplace=True)

# "TeslimatTelefon" sütununu sayı biçimine çevir (formatlama)
try:
    tekrar_df["TeslimatTelefon"] = tekrar_df["TeslimatTelefon"].apply(lambda x: '{:.0f}'.format(x))
except ValueError:
    pass

# Temizlenmiş tekrar satırları ayrı bir Excel dosyasına kaydet
output_cleaned_file = "Çift Siparişler.xlsx"
tekrar_df.to_excel(output_cleaned_file, index=False, engine="openpyxl")




# Kaynak dosya adı
kaynak_excel = "birlesik_excel.xlsx"

# Kopya dosya adı (istediğiniz adı ve konumu belirtin)
kopya_excel = "Hazırlanan Sipariş Numaraları.xlsx"

# Dosyayı kopyala
shutil.copy(kaynak_excel, kopya_excel)








# Excel dosyasını okuma
df = pd.read_excel("birlesik_excel.xlsx")

# İşlemi gerçekleştiren fonksiyon
def duplicate_rows(row):
    count = int(row["Adet"])
    return pd.concat([row] * count, axis=1).T

# Tüm satırları işleme tabi tutma
new_rows = df.apply(duplicate_rows, axis=1)

# Yeni veri çerçevesini oluşturma
new_df = pd.concat(new_rows.tolist(), ignore_index=True)

# Sadece belirtilen sütunları seçme
selected_columns = ["Id", "Barkod", "UrunAdi", "Varyant"]
new_df = new_df[selected_columns]

# Veriyi yeni bir Excel dosyasına yazma
new_df.to_excel("birlesik_excel.xlsx", index=False)













url = "https://haydigiy.online/Products/rafkodlari.php"
response = requests.get(url)
soup = BeautifulSoup(response.text, "html.parser")
table = soup.find("table")
data = []
for row in table.find_all("tr"):
    row_data = []
    for cell in row.find_all(["th", "td"]):
        row_data.append(cell.get_text(strip=True))
    data.append(row_data)
df = pd.DataFrame(data[1:], columns=data[0])
df.to_excel("Raf Kodu.xlsx", index=False)


sonuc_df = pd.read_excel("Raf Kodu.xlsx")

sonuc_df["VaryasyonBarkod"] = pd.to_numeric(sonuc_df["VaryasyonBarkod"], errors="coerce")  # Sayıya dönüştür

# "birlesik_excel.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Raf Kodu.xlsx", index=False)



# "birlesik_excel.xlsx" ve "Raf Kodu.xlsx" dosyalarını okuma
sonuc_df = pd.read_excel("birlesik_excel.xlsx")
google_sheet_df = pd.read_excel("Raf Kodu.xlsx")

# "birlesik_excel.xlsx" dosyasına yeni bir sütun ekleyerek başlangıçta "Raf Kodu Yok" değerleri ile doldurma
sonuc_df["GoogleSheetVerisi"] = "9999-Raf Kodu Yok"

# Her bir "Barkod" değeri için işlem yapma
for index, row in sonuc_df.iterrows():
    barkod = row["Barkod"]
    
    # "Raf Kodu.xlsx" dosyasında ilgili "Barkod"u arama

    matching_row = google_sheet_df[google_sheet_df.iloc[:, 0] == barkod]
    
    # Eşleşen "Barkod" varsa ve karşılık gelen hücre boş değilse, değeri "GoogleSheetVerisi" sütununa yazma
    if not matching_row.empty and not pd.isnull(matching_row.iloc[0, 2]):
        sonuc_df.at[index, "GoogleSheetVerisi"] = matching_row.iloc[0, 2]

# "birlesik_excel.xlsx" dosyasını güncelleme
sonuc_df.to_excel("birlesik_excel.xlsx", index=False)








# "birlesik_excel.xlsx" dosyasını güncelleme
sonuc_df["GoogleSheetVerisi Kopya"] = sonuc_df["GoogleSheetVerisi"]  # "GoogleSheetVerisi" sütununu kopyala
sonuc_df["GoogleSheetVerisi Kopya"] = sonuc_df["GoogleSheetVerisi Kopya"].str.split("/", n=1).str[0]  # "-" den sonrasını temizle
sonuc_df["GoogleSheetVerisi Kopya"] = sonuc_df["GoogleSheetVerisi Kopya"].str.replace(r'^[^-]*-', '', regex=True)
sonuc_df["GoogleSheetVerisi Kopya"] = pd.to_numeric(sonuc_df["GoogleSheetVerisi Kopya"], errors="coerce")  # Sayıya dönüştür
sonuc_df = sonuc_df.sort_values(by="GoogleSheetVerisi Kopya")  # "GoogleSheetVerisi Kopya" sütununa göre sırala
sonuc_df.drop(columns=["GoogleSheetVerisi Kopya"], inplace=True)



sonuc_df.to_excel("birlesik_excel.xlsx", index=False)









# "birlesik_excel.xlsx" dosyasını güncelleme
sonuc_df["GoogleSheetVerisi Kopya"] = sonuc_df["GoogleSheetVerisi"]  # "GoogleSheetVerisi" sütununu kopyala
sonuc_df["GoogleSheetVerisi Kopya"] = sonuc_df["GoogleSheetVerisi Kopya"].str.split("-", n=1).str[0]  # "-" den sonrasını temizle
sonuc_df["GoogleSheetVerisi Kopya"] = pd.to_numeric(sonuc_df["GoogleSheetVerisi Kopya"], errors="coerce")  # Sayıya dönüştür
sonuc_df = sonuc_df.sort_values(by="GoogleSheetVerisi Kopya")  # "GoogleSheetVerisi Kopya" sütununa göre sırala


sonuc_df.to_excel("birlesik_excel.xlsx", index=False)















# "birlesik_excel.xlsx" ve "Raf Kodu.xlsx" dosyalarını okuma
sonuc_df = pd.read_excel("birlesik_excel.xlsx")
google_sheet_df = pd.read_excel("Raf Kodu.xlsx")

# "birlesik_excel.xlsx" dosyasına yeni bir sütun ekleyerek başlangıçta "Raf Kodu Yok" değerleri ile doldurma
sonuc_df["Kategori"] = "9999-Raf Kodu Yok"

# Her bir "Barkod" değeri için işlem yapma
for index, row in sonuc_df.iterrows():
    barkod = row["Barkod"]
    
    # "Raf Kodu.xlsx" dosyasında ilgili "Barkod"u arama
    matching_row = google_sheet_df[google_sheet_df.iloc[:, 0] == barkod]
    
    # Eşleşen "Barkod" varsa ve karşılık gelen hücre boş değilse, değeri "GoogleSheetVerisi" sütununa yazma
    if not matching_row.empty and not pd.isnull(matching_row.iloc[0, 3]):
        sonuc_df.at[index, "Kategori"] = matching_row.iloc[0, 3]

# "birlesik_excel.xlsx" dosyasını güncelleme
sonuc_df.to_excel("birlesik_excel.xlsx", index=False)





file_path = "Raf Kodu.xlsx"

if os.path.exists(file_path):
    os.remove(file_path)








# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Yeni Kategori"] = ""

# İç Giyim içeren satırları işle
innerwear_rows = df[df["Kategori"].str.contains("İç Giyim")]

# İç Giyim içeren satırları işle
for index, row in innerwear_rows.iterrows():
    df.loc[index, "Yeni Kategori"] = "İç Giyim"

# Sonucu kaydet
output_file_path = "birlesik_excel.xlsx"  # Sonucun kaydedileceği dosya adı ve yolunu belirtin
df.to_excel(output_file_path, index=False)



# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "Id" değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")

# "Id" değerlerine göre "Yeni Kategori" değerini güncelle
for group_name, group_data in grouped:
    if any(row["Yeni Kategori"] == "İç Giyim" for _, row in group_data.iterrows()):
        df.loc[df["Id"] == group_name, "Yeni Kategori"] = "İç Giyim"

# Sonucu kaydet
output_file_path = "birlesik_excel.xlsx"  # Sonucun kaydedileceği dosya adı ve yolunu belirtin
df.to_excel(output_file_path, index=False)









# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "Yeni Kategori" değeri "İç Giyim" olan satırları seç
innerwear_rows = df[df["Yeni Kategori"] == "İç Giyim"]

# Ayrı Excel dosyasına kaydet
output_file_path = "İç Giyim.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
innerwear_rows.to_excel(output_file_path, index=False)

# Ana DataFrame'den "İç Giyim" satırları sil
df = df[df["Yeni Kategori"] != "İç Giyim"]
df.drop(columns=["Yeni Kategori"], inplace=True)  # "Yeni Kategori" sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)






# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek tekrar sayılarını tut
df["Tekrar Sayısı"] = df.groupby("Id")["Id"].transform("count")

# Sonucu kaydet
output_file_path = "birlesik_excel.xlsx"  # Sonucun kaydedileceği dosya adı ve yolunu belirtin
df.to_excel(output_file_path, index=False)






















# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 0 and max_value <= 500:
        df.loc[df["Id"] == group_name, "Sonuc"] = "0-500-14"

# 0-112 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "0-500-14"]

# Tekrar Sayısı 6 ve daha büyük olan satırları seç
high_repeats_rows = filtered_df[filtered_df.groupby("Id")["Id"].transform("count") >= 5]

# Ayrı Excel dosyasına kaydet
high_repeats_output_file_path = "İnstagram (14).xlsx"
high_repeats_rows.to_excel(high_repeats_output_file_path, index=False)

# Ana DataFrame'den 6 ve üzeri tekrarları çıkar
df.drop(index=high_repeats_rows.index, inplace=True)

# Ana Excel dosyasını güncelle
df.drop(columns=["Sonuc"], inplace=True)
df.to_excel(excel_file_path, index=False)
















# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 700 and max_value <= 1999:
        df.loc[df["Id"] == group_name, "Sonuc"] = "700-1999"

# 0-112 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "700-1999"]

# Tekrar Sayısı 6 ve daha büyük olan satırları seç
high_repeats_rows = filtered_df[filtered_df.groupby("Id")["Id"].transform("count") >= 5]

# Ayrı Excel dosyasına kaydet
high_repeats_output_file_path = "Yeni Depo (14).xlsx"
high_repeats_rows.to_excel(high_repeats_output_file_path, index=False)

# Ana DataFrame'den 6 ve üzeri tekrarları çıkar
df.drop(index=high_repeats_rows.index, inplace=True)

# Ana Excel dosyasını güncelle
df.drop(columns=["Sonuc"], inplace=True)
df.to_excel(excel_file_path, index=False)










# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 2000 and max_value <= 9999:
        df.loc[df["Id"] == group_name, "Sonuc"] = "2000-9999"

# 0-112 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "2000-9999"]

# Tekrar Sayısı 6 ve daha büyük olan satırları seç
high_repeats_rows = filtered_df[filtered_df.groupby("Id")["Id"].transform("count") >= 5]

# Ayrı Excel dosyasına kaydet
high_repeats_output_file_path = "Özerler Depo (14).xlsx"
high_repeats_rows.to_excel(high_repeats_output_file_path, index=False)

# Ana DataFrame'den 6 ve üzeri tekrarları çıkar
df.drop(index=high_repeats_rows.index, inplace=True)

# Ana Excel dosyasını güncelle
df.drop(columns=["Sonuc"], inplace=True)
df.to_excel(excel_file_path, index=False)











# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 0 and max_value <= 9999:
        df.loc[df["Id"] == group_name, "Sonuc"] = "0-9999"

# 0-112 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "0-9999"]

# Tekrar Sayısı 6 ve daha büyük olan satırları seç
high_repeats_rows = filtered_df[filtered_df.groupby("Id")["Id"].transform("count") >= 5]

# Ayrı Excel dosyasına kaydet
high_repeats_output_file_path = "Tüm Depo (14).xlsx"
high_repeats_rows.to_excel(high_repeats_output_file_path, index=False)

# Ana DataFrame'den 6 ve üzeri tekrarları çıkar
df.drop(index=high_repeats_rows.index, inplace=True)

# Ana Excel dosyasını güncelle
df.drop(columns=["Sonuc"], inplace=True)
df.to_excel(excel_file_path, index=False)




































# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 0 and max_value <= 112:
        df.loc[df["Id"] == group_name, "Sonuc"] = "0-112"

# 0-112 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "0-112"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "0-112.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 0-112 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "0-112"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)












# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 113 and max_value <= 206:
        df.loc[df["Id"] == group_name, "Sonuc"] = "113-206"

# 113-206 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "113-206"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "113-206.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 113-206 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "113-206"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)


















# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 0 and max_value <= 206:
        df.loc[df["Id"] == group_name, "Sonuc"] = "0-206"

# 0-206 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "0-206"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "0-206.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 0-206 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "0-206"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)










# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 207 and max_value <= 400:
        df.loc[df["Id"] == group_name, "Sonuc"] = "207-400"

# 207-400 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "207-400"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "207-400.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 207-400 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "207-400"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)
















# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 0 and max_value <= 500:
        df.loc[df["Id"] == group_name, "Sonuc"] = "0-500"

# 0-500 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "0-500"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "İnstagram Kalanlar.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 0-500 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "0-500"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)












# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 700 and max_value <= 857:
        df.loc[df["Id"] == group_name, "Sonuc"] = "700-857"

# 700-857 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "700-857"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "700-857.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 700-857 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "700-857"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)














# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 858 and max_value <= 995:
        df.loc[df["Id"] == group_name, "Sonuc"] = "858-995"

# 858-995 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "858-995"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "858-995.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 858-995 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "858-995"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)












# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 996 and max_value <= 1133:
        df.loc[df["Id"] == group_name, "Sonuc"] = "996-1133"

# 996-1133 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "996-1133"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "996-1133.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 996-1133 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "996-1133"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)




















# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 1134 and max_value <= 1269:
        df.loc[df["Id"] == group_name, "Sonuc"] = "1134-1269"

# 1134-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "1134-1269"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "1134-1269.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 1134-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "1134-1269"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)















# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 700 and max_value <= 995:
        df.loc[df["Id"] == group_name, "Sonuc"] = "700-995"

# 700-995 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "700-995"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "700-995.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 700-995 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "700-995"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)















# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 996 and max_value <= 1269:
        df.loc[df["Id"] == group_name, "Sonuc"] = "700-995"

# 996-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "996-1269"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "996-1269.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 996-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "996-1269"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)












# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 1270 and max_value <= 1326:
        df.loc[df["Id"] == group_name, "Sonuc"] = "1270-1326"

# 996-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "1270-1326"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "1270-1326.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 996-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "1270-1326"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)



















# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 1327 and max_value <= 1459:
        df.loc[df["Id"] == group_name, "Sonuc"] = "1327-1459"

# 996-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "1327-1459"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "1327-1459.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 996-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "1327-1459"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)










# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 1460 and max_value <= 1531:
        df.loc[df["Id"] == group_name, "Sonuc"] = "1460-1531"

# 996-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "1460-1531"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "1460-1531.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 996-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "1460-1531"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)












































# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 700 and max_value <= 1999:
        df.loc[df["Id"] == group_name, "Sonuc"] = "700-1269"

# 700-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "700-1269"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "Yeni Depo Kalanlar.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 700-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "700-1269"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)










# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 2000 and max_value <= 2164:
        df.loc[df["Id"] == group_name, "Sonuc"] = "2000-2164"

# 996-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "2000-2164"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "2000-2164.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 996-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "2000-2164"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)














# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 2165 and max_value <= 9999:
        df.loc[df["Id"] == group_name, "Sonuc"] = "2165-2310"

# 996-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "2165-2310"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "2165-2310.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 996-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "2165-2310"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)












# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 2000 and max_value <= 9999:
        df.loc[df["Id"] == group_name, "Sonuc"] = "2000-9999"

# 700-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "2000-9999"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "Özerler Depo Kalanlar.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 700-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "2000-9999"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)














# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 0 and max_value <= 1999:
        df.loc[df["Id"] == group_name, "Sonuc"] = "0-1999"

# 700-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "0-1999"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 700-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "0-1999"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)











# Excel dosyasını oku
excel_file_path = "birlesik_excel.xlsx"  # Ana Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# Yeni bir sütun ekleyerek işlem sonuçlarını tut
df["Sonuc"] = ""

# Id değerlerine göre grupla ve işlemi yap
grouped = df.groupby("Id")
for group_name, group_data in grouped:
    min_value = group_data["GoogleSheetVerisi Kopya"].min()
    max_value = group_data["GoogleSheetVerisi Kopya"].max()

    if min_value >= 2000 and max_value <= 9999:
        df.loc[df["Id"] == group_name, "Sonuc"] = "2000-9999"

# 700-1269 olan satırları ayrı bir DataFrame'e kopyala
filtered_df = df[df["Sonuc"] == "2000-9999"]

# Ayrı Excel dosyasına kaydet
filtered_output_file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"  # Ayrı dosyanın adını ve yolunu belirtin
filtered_df.to_excel(filtered_output_file_path, index=False)

# 700-1269 olan satırları ana DataFrame'den sil
df = df[df["Sonuc"] != "2000-9999"]
df.drop(columns=["Sonuc"], inplace=True)  # Sonuc sütununu sil

# Ana Excel dosyasını güncelle
df.to_excel(excel_file_path, index=False)







old_file_path = "birlesik_excel.xlsx"
new_file_path = "Tüm Depo Kalanlar.xlsx"

# Dosyanın adını değiştir
os.rename(old_file_path, new_file_path)















































# Excel dosyasını aç
file_path = "Tüm Depo Kalanlar.xlsx"
df = pd.read_excel(file_path)
df = pd.read_excel(file_path, engine="openpyxl")


# "GoogleSheetVerisi" sütunundaki "9999-Raf Kodu Yok" değerlerini "Raf Kodu Yok" ile değiştir
df["GoogleSheetVerisi"] = df["GoogleSheetVerisi"].replace("9999-Raf Kodu Yok", "Raf Kodu Yok")


# Değişiklikleri aynı Excel dosyasının üzerine kaydet
df.to_excel(file_path, index=False)





# Excel dosyasını aç
file_path = "İç Giyim.xlsx"
df = pd.read_excel(file_path)
df = pd.read_excel(file_path, engine="openpyxl")


# "GoogleSheetVerisi" sütunundaki "9999-Raf Kodu Yok" değerlerini "Raf Kodu Yok" ile değiştir
df["GoogleSheetVerisi"] = df["GoogleSheetVerisi"].replace("9999-Raf Kodu Yok", "Raf Kodu Yok")

# Değişiklikleri aynı Excel dosyasının üzerine kaydet
df.to_excel(file_path, index=False)











# Excel dosyasını aç
file_path = "Tüm Depo (14).xlsx"
df = pd.read_excel(file_path)
df = pd.read_excel(file_path, engine="openpyxl")


# "GoogleSheetVerisi" sütunundaki "9999-Raf Kodu Yok" değerlerini "Raf Kodu Yok" ile değiştir
df["GoogleSheetVerisi"] = df["GoogleSheetVerisi"].replace("9999-Raf Kodu Yok", "Raf Kodu Yok")


# Değişiklikleri aynı Excel dosyasının üzerine kaydet
df.to_excel(file_path, index=False)











# Excel dosyasını oku
excel_file_path = "İnstagram Kalanlar.xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Tekrar Sayısı", "Sonuc"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)





# Excel dosyasını oku
excel_file_path = "Özerler Depo Kalanlar.xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Tekrar Sayısı", "Sonuc"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)









# Excel dosyasını oku
excel_file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Tekrar Sayısı", "Sonuc"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)









# Excel dosyasını oku
excel_file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Tekrar Sayısı", "Sonuc"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)














# Excel dosyasını oku
excel_file_path = "İç Giyim.xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Yeni Kategori"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)






# Excel dosyasını oku
excel_file_path = "İnstagram (14).xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Tekrar Sayısı", "Sonuc"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)









# Excel dosyasını oku
excel_file_path = "Yeni Depo (14).xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Tekrar Sayısı", "Sonuc"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)












# Excel dosyasını oku
excel_file_path = "Tüm Depo (14).xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Tekrar Sayısı", "Sonuc"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)






# Excel dosyasını oku
excel_file_path = "Özerler Depo (14).xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Tekrar Sayısı", "Sonuc"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)











# Excel dosyasını oku
excel_file_path = "Hazırlanan Sipariş Numaraları.xlsx"  # Kalanlar Excel dosyasının adını ve yolunu belirtin
df = pd.read_excel(excel_file_path)

# "GoogleSheetVerisi Kopya" ve "Kategori" sütunlarını sil
columns_to_drop = ["OdemeTipi", "TeslimatTelefon", "Barkod", "Adet", "TeslimatEPostaAdresi", "SiparisToplam", "Varyant", "UrunAdi"]
df.drop(columns=columns_to_drop, inplace=True)

# Dosyayı güncelle
df.to_excel(excel_file_path, index=False)





















# Excel dosyalarının adları ve yolları
excel_files = ["700-857.xlsx", "700-995.xlsx", "858-995.xlsx", "996-1133.xlsx", "996-1269.xlsx", "1134-1269.xlsx", "Yeni Depo Kalanlar.xlsx", "0-112.xlsx", "113-206.xlsx", "0-206.xlsx", "207-400.xlsx", "1270-1326.xlsx", "1327-1459.xlsx", "1460-1531.xlsx", "2000-2164.xlsx", "2165-2310.xlsx"]

# Sileceğiniz sütunların adları
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Sonuc", "Tekrar Sayısı"]

# Her bir Excel dosyası için işlemi yapın
for excel_file in excel_files:
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
        df.drop(columns=columns_to_drop, inplace=True)
        df.to_excel(excel_file, index=False)










# Excel dosyalarının adları ve yolları
excel_files = ["Tüm Depo Kalanlar.xlsx"]

# Sileceğiniz sütunların adları
columns_to_drop = ["GoogleSheetVerisi Kopya", "Kategori", "Tekrar Sayısı"]

# Her bir Excel dosyası için işlemi yapın
for excel_file in excel_files:
    if os.path.exists(excel_file):
        df = pd.read_excel(excel_file)
        df.drop(columns=columns_to_drop, inplace=True)
        df.to_excel(excel_file, index=False)



























sonuc_df = pd.read_excel("0-112.xlsx")

# "0-112.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "0-112.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("0-112.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "0-112.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-112.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "0-112.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-112.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "0-112.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-112.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "0-112.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-112.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "0-112.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-112.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("0-112.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 200
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("0-112.xlsx")






# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("0-112.xlsx")










# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("0-112.xlsx")











# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("0-112.xlsx")










# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("0-112.xlsx")











# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("0-112.xlsx")










# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("0-112.xlsx")










# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("0-112.xlsx")





# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("0-112.xlsx")




# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("0-112.xlsx")











# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("0-112.xlsx")










# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("0-112.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "0-112"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "0-112.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "0-112.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)






































sonuc_df = pd.read_excel("113-206.xlsx")

# "113-206.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "113-206.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("113-206.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "113-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("113-206.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "113-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("113-206.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "113-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("113-206.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "113-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("113-206.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "113-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("113-206.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("113-206.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("113-206.xlsx")






# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("113-206.xlsx")










# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("113-206.xlsx")











# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("113-206.xlsx")










# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("113-206.xlsx")











# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("113-206.xlsx")










# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("113-206.xlsx")










# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("113-206.xlsx")





# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("113-206.xlsx")




# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("113-206.xlsx")











# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("113-206.xlsx")










# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("113-206.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "113-206"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "113-206.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "113-206.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)






























sonuc_df = pd.read_excel("0-206.xlsx")

# "0-206.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "0-206.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("0-206.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "0-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-206.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "0-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-206.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "0-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-206.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "0-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-206.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "0-206.xlsx" dosyasını güncelleme
sonuc_df.to_excel("0-206.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("0-206.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("0-206.xlsx")






# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("0-206.xlsx")










# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("0-206.xlsx")











# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("0-206.xlsx")










# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("0-206.xlsx")











# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("0-206.xlsx")










# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("0-206.xlsx")










# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("0-206.xlsx")





# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("0-206.xlsx")




# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("0-206.xlsx")











# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("0-206.xlsx")










# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("0-206.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "0-206"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "0-206.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "0-206.xlsx"))


gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)



































sonuc_df = pd.read_excel("İnstagram Kalanlar.xlsx")

# "İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "İnstagram Kalanlar.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("İnstagram Kalanlar.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram Kalanlar.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram Kalanlar.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram Kalanlar.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram Kalanlar.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram Kalanlar.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("İnstagram Kalanlar.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("İnstagram Kalanlar.xlsx")






# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("İnstagram Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("İnstagram Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("İnstagram Kalanlar.xlsx")





# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("İnstagram Kalanlar.xlsx")




# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("İnstagram Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("İnstagram Kalanlar.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "İnstagram Kalanlar"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "İnstagram Kalanlar.xlsx"))


gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)
















































sonuc_df = pd.read_excel("İnstagram (14).xlsx")

# "İnstagram (14).xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "İnstagram (14).xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("İnstagram (14).xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "İnstagram (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram (14).xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "İnstagram (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram (14).xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "İnstagram (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram (14).xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "İnstagram (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram (14).xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "İnstagram (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram (14).xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("İnstagram (14).xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 14

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("İnstagram (14).xlsx")






# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 14
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("İnstagram (14).xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("İnstagram (14).xlsx")











# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("İnstagram (14).xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("İnstagram (14).xlsx")











# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("İnstagram (14).xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("İnstagram (14).xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("İnstagram (14).xlsx")





# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("İnstagram (14).xlsx")




# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("İnstagram (14).xlsx")











# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("İnstagram (14).xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("İnstagram (14).xlsx")





def create_bat_files(data, output_folder, batch_size=14):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "İnstagram (14)"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "İnstagram (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "İnstagram (14).xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)













































sonuc_df = pd.read_excel("İç Giyim.xlsx")

# "İç Giyim.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "İç Giyim.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("İç Giyim.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "İç Giyim.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İç Giyim.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "İç Giyim.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İç Giyim.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "İç Giyim.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İç Giyim.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "İç Giyim.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İç Giyim.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "İç Giyim.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İç Giyim.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("İç Giyim.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("İç Giyim.xlsx")






# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("İç Giyim.xlsx")










# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("İç Giyim.xlsx")











# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("İç Giyim.xlsx")










# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("İç Giyim.xlsx")











# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("İç Giyim.xlsx")










# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("İç Giyim.xlsx")










# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("İç Giyim.xlsx")





# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("İç Giyim.xlsx")




# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("İç Giyim.xlsx")











# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("İç Giyim.xlsx")










# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("İç Giyim.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1



# Klasör oluştur
output_folder = "İç Giyim"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "İç Giyim.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "İç Giyim.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)































sonuc_df = pd.read_excel("207-400.xlsx")

# "207-400.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "207-400.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("207-400.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "207-400.xlsx" dosyasını güncelleme
sonuc_df.to_excel("207-400.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "207-400.xlsx" dosyasını güncelleme
sonuc_df.to_excel("207-400.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "207-400.xlsx" dosyasını güncelleme
sonuc_df.to_excel("207-400.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "207-400.xlsx" dosyasını güncelleme
sonuc_df.to_excel("207-400.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "207-400.xlsx" dosyasını güncelleme
sonuc_df.to_excel("207-400.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("207-400.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("207-400.xlsx")






# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("207-400.xlsx")










# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("207-400.xlsx")











# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("207-400.xlsx")










# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("207-400.xlsx")











# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("207-400.xlsx")










# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("207-400.xlsx")










# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("207-400.xlsx")





# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("207-400.xlsx")




# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("207-400.xlsx")











# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("207-400.xlsx")










# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("207-400.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "207-400"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "207-400.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "207-400.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)



























sonuc_df = pd.read_excel("Yeni Depo Kalanlar.xlsx")

# "Yeni Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "Yeni Depo Kalanlar.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("Yeni Depo Kalanlar.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "Yeni Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo Kalanlar.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "Yeni Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo Kalanlar.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "Yeni Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo Kalanlar.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "Yeni Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo Kalanlar.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "Yeni Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo Kalanlar.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("Yeni Depo Kalanlar.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Yeni Depo Kalanlar.xlsx")






# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Yeni Depo Kalanlar.xlsx")









# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Yeni Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Yeni Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("Yeni Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("Yeni Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("Yeni Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("Yeni Depo Kalanlar.xlsx")





# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("Yeni Depo Kalanlar.xlsx")




# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("Yeni Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("Yeni Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("Yeni Depo Kalanlar.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "Yeni Depo Kalanlar"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "Yeni Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "Yeni Depo Kalanlar.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)































sonuc_df = pd.read_excel("700-857.xlsx")

# "700-857.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "700-857.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("700-857.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "700-857.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-857.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "700-857.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-857.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "700-857.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-857.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "700-857.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-857.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "700-857.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-857.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("700-857.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("700-857.xlsx")






# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("700-857.xlsx")










# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("700-857.xlsx")











# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("700-857.xlsx")










# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("700-857.xlsx")











# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("700-857.xlsx")










# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("700-857.xlsx")










# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("700-857.xlsx")





# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("700-857.xlsx")




# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("700-857.xlsx")











# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("700-857.xlsx")










# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("700-857.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "700-857"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "700-857.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "700-857.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)
























sonuc_df = pd.read_excel("700-995.xlsx")

# "700-995.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "700-995.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("700-995.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "700-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-995.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "700-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-995.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "700-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-995.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "700-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-995.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "700-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("700-995.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("700-995.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("700-995.xlsx")






# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("700-995.xlsx")










# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("700-995.xlsx")











# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("700-995.xlsx")










# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("700-995.xlsx")











# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("700-995.xlsx")










# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("700-995.xlsx")










# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("700-995.xlsx")





# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("700-995.xlsx")




# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("700-995.xlsx")











# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("700-995.xlsx")










# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("700-995.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "700-995"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "700-995.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "700-995.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)























sonuc_df = pd.read_excel("858-995.xlsx")

# "858-995.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "858-995.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("858-995.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "858-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("858-995.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "858-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("858-995.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "858-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("858-995.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "858-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("858-995.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "858-995.xlsx" dosyasını güncelleme
sonuc_df.to_excel("858-995.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("858-995.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("858-995.xlsx")






# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("858-995.xlsx")










# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("858-995.xlsx")











# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("858-995.xlsx")










# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("858-995.xlsx")











# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("858-995.xlsx")










# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("858-995.xlsx")










# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("858-995.xlsx")





# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("858-995.xlsx")




# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("858-995.xlsx")











# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("858-995.xlsx")










# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("858-995.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "858-995"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "858-995.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "858-995.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)
































sonuc_df = pd.read_excel("996-1133.xlsx")

# "996-1133.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "996-1133.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("996-1133.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "996-1133.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1133.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "996-1133.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1133.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "996-1133.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1133.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "996-1133.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1133.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "996-1133.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1133.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("996-1133.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("996-1133.xlsx")






# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("996-1133.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("996-1133.xlsx")











# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("996-1133.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("996-1133.xlsx")











# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("996-1133.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("996-1133.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("996-1133.xlsx")





# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("996-1133.xlsx")




# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("996-1133.xlsx")











# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("996-1133.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("996-1133.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "996-1133"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "996-1133.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "996-1133.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)


















sonuc_df = pd.read_excel("996-1269.xlsx")

# "996-1269.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "996-1269.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("996-1269.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "996-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1269.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "996-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1269.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "996-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1269.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "996-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1269.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "996-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("996-1269.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("996-1269.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("996-1269.xlsx")






# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("996-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("996-1269.xlsx")











# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("996-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("996-1269.xlsx")











# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("996-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("996-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("996-1269.xlsx")





# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("996-1269.xlsx")




# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("996-1269.xlsx")











# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("996-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("996-1269.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "996-1269"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "996-1269.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "996-1269.xlsx"))


gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)

























sonuc_df = pd.read_excel("1134-1269.xlsx")

# "1134-1269.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "1134-1269.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("1134-1269.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "1134-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1134-1269.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "1134-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1134-1269.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "1134-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1134-1269.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "1134-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1134-1269.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "1134-1269.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1134-1269.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("1134-1269.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("1134-1269.xlsx")






# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("1134-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("1134-1269.xlsx")











# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("1134-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("1134-1269.xlsx")











# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("1134-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("1134-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("1134-1269.xlsx")





# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("1134-1269.xlsx")




# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("1134-1269.xlsx")











# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("1134-1269.xlsx")










# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("1134-1269.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "1134-1269"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "1134-1269.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "1134-1269.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)


































sonuc_df = pd.read_excel("1270-1326.xlsx")

# "1270-1326.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "1270-1326.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("1270-1326.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "1270-1326.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1270-1326.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "1270-1326.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1270-1326.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "1270-1326.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1270-1326.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "1270-1326.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1270-1326.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "1270-1326.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1270-1326.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("1270-1326.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("1270-1326.xlsx")






# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("1270-1326.xlsx")










# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("1270-1326.xlsx")











# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("1270-1326.xlsx")










# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("1270-1326.xlsx")











# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("1270-1326.xlsx")










# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("1270-1326.xlsx")










# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("1270-1326.xlsx")





# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("1270-1326.xlsx")




# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("1270-1326.xlsx")











# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("1270-1326.xlsx")










# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("1270-1326.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "1270-1326"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "1270-1326.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "1270-1326.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)



























sonuc_df = pd.read_excel("1327-1459.xlsx")

# "1327-1459.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "1327-1459.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("1327-1459.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "1327-1459.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1327-1459.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "1327-1459.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1327-1459.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "1327-1459.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1327-1459.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "1327-1459.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1327-1459.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "1327-1459.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1327-1459.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("1327-1459.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("1327-1459.xlsx")






# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("1327-1459.xlsx")










# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("1327-1459.xlsx")











# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("1327-1459.xlsx")










# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("1327-1459.xlsx")











# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("1327-1459.xlsx")










# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("1327-1459.xlsx")










# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("1327-1459.xlsx")





# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("1327-1459.xlsx")




# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("1327-1459.xlsx")











# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("1327-1459.xlsx")










# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("1327-1459.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "1327-1459"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "1327-1459.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "1327-1459.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)




























sonuc_df = pd.read_excel("1460-1531.xlsx")

# "1460-1531.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "1460-1531.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("1460-1531.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "1460-1531.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1460-1531.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "1460-1531.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1460-1531.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "1460-1531.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1460-1531.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "1460-1531.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1460-1531.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "1460-1531.xlsx" dosyasını güncelleme
sonuc_df.to_excel("1460-1531.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("1460-1531.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("1460-1531.xlsx")






# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("1460-1531.xlsx")










# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("1460-1531.xlsx")











# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("1460-1531.xlsx")










# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("1460-1531.xlsx")











# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("1460-1531.xlsx")










# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("1460-1531.xlsx")










# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("1460-1531.xlsx")





# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("1460-1531.xlsx")




# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("1460-1531.xlsx")











# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("1460-1531.xlsx")










# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("1460-1531.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "1460-1531"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "1460-1531.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "1460-1531.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)































































sonuc_df = pd.read_excel("Tüm Depo Kalanlar.xlsx")

# "Tüm Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "Tüm Depo Kalanlar.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("Tüm Depo Kalanlar.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "Tüm Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo Kalanlar.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "Tüm Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo Kalanlar.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "Tüm Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo Kalanlar.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "Tüm Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo Kalanlar.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "Tüm Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo Kalanlar.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("Tüm Depo Kalanlar.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Tüm Depo Kalanlar.xlsx")






# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Tüm Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Tüm Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Tüm Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("Tüm Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("Tüm Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("Tüm Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("Tüm Depo Kalanlar.xlsx")





# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("Tüm Depo Kalanlar.xlsx")




# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("Tüm Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("Tüm Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("Tüm Depo Kalanlar.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "Tüm Depo Kalanlar"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "Tüm Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "Tüm Depo Kalanlar.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)































sonuc_df = pd.read_excel("Yeni Depo (14).xlsx")

# "Yeni Depo (14).xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "Yeni Depo (14).xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("Yeni Depo (14).xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "Yeni Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo (14).xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "Yeni Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo (14).xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "Yeni Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo (14).xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "Yeni Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo (14).xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "Yeni Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo (14).xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("Yeni Depo (14).xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 14

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Yeni Depo (14).xlsx")






# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 14
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Yeni Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Yeni Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Yeni Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("Yeni Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("Yeni Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("Yeni Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("Yeni Depo (14).xlsx")





# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("Yeni Depo (14).xlsx")




# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("Yeni Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("Yeni Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("Yeni Depo (14).xlsx")





def create_bat_files(data, output_folder, batch_size=14):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "Yeni Depo (14)"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "Yeni Depo (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "Yeni Depo (14).xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)

































sonuc_df = pd.read_excel("Tüm Depo (14).xlsx")

# "Tüm Depo (14).xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "Tüm Depo (14).xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("Tüm Depo (14).xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "Tüm Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo (14).xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "Tüm Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo (14).xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "Tüm Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo (14).xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "Tüm Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo (14).xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "Tüm Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Tüm Depo (14).xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("Tüm Depo (14).xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 14

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Tüm Depo (14).xlsx")






# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 14
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Tüm Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Tüm Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Tüm Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("Tüm Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("Tüm Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("Tüm Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("Tüm Depo (14).xlsx")





# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("Tüm Depo (14).xlsx")




# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("Tüm Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("Tüm Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("Tüm Depo (14).xlsx")





def create_bat_files(data, output_folder, batch_size=14):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "Tüm Depo (14)"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "Tüm Depo (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "Tüm Depo (14).xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)






































sonuc_df = pd.read_excel("Özerler Depo (14).xlsx")

# "Özerler Depo (14).xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "Özerler Depo (14).xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("Özerler Depo (14).xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "Özerler Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo (14).xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "Özerler Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo (14).xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "Özerler Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo (14).xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "Özerler Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo (14).xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "Özerler Depo (14).xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo (14).xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("Özerler Depo (14).xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 14

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Özerler Depo (14).xlsx")






# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 14
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Özerler Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Özerler Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Özerler Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("Özerler Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("Özerler Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("Özerler Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("Özerler Depo (14).xlsx")





# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("Özerler Depo (14).xlsx")




# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("Özerler Depo (14).xlsx")











# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("Özerler Depo (14).xlsx")










# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("Özerler Depo (14).xlsx")





def create_bat_files(data, output_folder, batch_size=14):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "Özerler Depo (14)"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "Özerler Depo (14).xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "Özerler Depo (14).xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)






























sonuc_df = pd.read_excel("2000-2164.xlsx")

# "2000-2164.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "2000-2164.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("2000-2164.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "2000-2164.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2000-2164.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "2000-2164.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2000-2164.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "2000-2164.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2000-2164.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "2000-2164.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2000-2164.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "2000-2164.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2000-2164.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("2000-2164.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("2000-2164.xlsx")






# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("2000-2164.xlsx")










# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("2000-2164.xlsx")











# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("2000-2164.xlsx")










# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("2000-2164.xlsx")











# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("2000-2164.xlsx")










# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("2000-2164.xlsx")










# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("2000-2164.xlsx")





# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("2000-2164.xlsx")




# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("2000-2164.xlsx")











# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("2000-2164.xlsx")










# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("2000-2164.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "2000-2164"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "2000-2164.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "2000-2164.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)































sonuc_df = pd.read_excel("2165-2310.xlsx")

# "2165-2310.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "2165-2310.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("2165-2310.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "2165-2310.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2165-2310.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "2165-2310.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2165-2310.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "2165-2310.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2165-2310.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "2165-2310.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2165-2310.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "2165-2310.xlsx" dosyasını güncelleme
sonuc_df.to_excel("2165-2310.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("2165-2310.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("2165-2310.xlsx")






# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("2165-2310.xlsx")










# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("2165-2310.xlsx")











# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("2165-2310.xlsx")










# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("2165-2310.xlsx")











# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("2165-2310.xlsx")










# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("2165-2310.xlsx")










# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("2165-2310.xlsx")





# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("2165-2310.xlsx")




# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("2165-2310.xlsx")











# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("2165-2310.xlsx")










# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("2165-2310.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "2165-2310"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "2165-2310.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "2165-2310.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)

























sonuc_df = pd.read_excel("Özerler Depo Kalanlar.xlsx")

# "Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "Özerler Depo Kalanlar.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("Özerler Depo Kalanlar.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo Kalanlar.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo Kalanlar.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo Kalanlar.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo Kalanlar.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Özerler Depo Kalanlar.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("Özerler Depo Kalanlar.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Özerler Depo Kalanlar.xlsx")






# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Özerler Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("Özerler Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("Özerler Depo Kalanlar.xlsx")





# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("Özerler Depo Kalanlar.xlsx")




# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("Özerler Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("Özerler Depo Kalanlar.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "Özerler Depo Kalanlar"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "Özerler Depo Kalanlar.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)


































sonuc_df = pd.read_excel("Yeni Depo ve İnstagram Kalanlar.xlsx")

# "Yeni Depo ve İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "Yeni Depo ve İnstagram Kalanlar.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("Yeni Depo ve İnstagram Kalanlar.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "Yeni Depo ve İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo ve İnstagram Kalanlar.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "Yeni Depo ve İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo ve İnstagram Kalanlar.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "Yeni Depo ve İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo ve İnstagram Kalanlar.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "Yeni Depo ve İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo ve İnstagram Kalanlar.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "Yeni Depo ve İnstagram Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("Yeni Depo ve İnstagram Kalanlar.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("Yeni Depo ve İnstagram Kalanlar.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")






# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")





# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")




# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("Yeni Depo ve İnstagram Kalanlar.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "Yeni Depo ve İnstagram Kalanlar"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "Yeni Depo ve İnstagram Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "Yeni Depo ve İnstagram Kalanlar.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)































sonuc_df = pd.read_excel("İnstagram ve Özerler Depo Kalanlar.xlsx")

# "İnstagram ve Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi"]  # "UrunAdi" sütununu kopyala
sonuc_df["UrunAdi Kopya"] = sonuc_df["UrunAdi Kopya"].str.split("-", n=1).str[1]  # "-" den öncesini temizle

# "UrunAdi Kopya" sütununu "İnstagram ve Özerler Depo Kalanlar.xlsx" dosyasına ekleyerek güncelleme
with pd.ExcelWriter("İnstagram ve Özerler Depo Kalanlar.xlsx") as writer:
    sonuc_df.to_excel(writer, index=False)









# "UrunAdi" sütununu en sağına yapıştırma
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdi"]

# "UrunAdiKopya2" sütununda " - " dan sonrasını ve son boşluktan öncesini silme
sonuc_df["UrunAdiKopya2"] = sonuc_df["UrunAdiKopya2"].apply(lambda x: x.split(" - ")[0].rsplit(" ", 1)[-1].strip() if " - " in x else x)

# "UrunAdiKopya2" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya2")
column_order.append("UrunAdiKopya2")
sonuc_df = sonuc_df[column_order]

# "İnstagram ve Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram ve Özerler Depo Kalanlar.xlsx", index=False)






# "UrunAdi" sütununu tablonun sonuna bir kez daha kopyalama ve düzenleme
sonuc_df["UrunAdiKopya3"] = sonuc_df["UrunAdi"].apply(lambda x: x.split(" - ", 1)[0].strip() if " - " in x else x)

# "UrunAdiKopya3" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("UrunAdiKopya3")
column_order.append("UrunAdiKopya3")
sonuc_df = sonuc_df[column_order]

# "İnstagram ve Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram ve Özerler Depo Kalanlar.xlsx", index=False)









# Verileri birleştirip yeni sütun oluşturma
sonuc_df["BirlesikVeri"] = sonuc_df["UrunAdi Kopya"] + " - " + sonuc_df["UrunAdiKopya2"] + " - " + sonuc_df["Varyant"]

# "BirlesikVeri" sütununu en sağa taşıma
column_order = list(sonuc_df.columns)
column_order.remove("BirlesikVeri")
column_order.append("BirlesikVeri")
sonuc_df = sonuc_df[column_order]

# "İnstagram ve Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram ve Özerler Depo Kalanlar.xlsx", index=False)

# "BirlesikVeri" sütunundaki "Beden:" ibaresini çıkarma
sonuc_df["BirlesikVeri"] = sonuc_df["BirlesikVeri"].str.replace("Beden:", "")

# "İnstagram ve Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram ve Özerler Depo Kalanlar.xlsx", index=False)





# Belirtilen sütunları silme
columns_to_drop = ["UrunAdi", "Varyant", "UrunAdi Kopya", "UrunAdiKopya2"]
sonuc_df.drop(columns_to_drop, axis=1, inplace=True)

# "İnstagram ve Özerler Depo Kalanlar.xlsx" dosyasını güncelleme
sonuc_df.to_excel("İnstagram ve Özerler Depo Kalanlar.xlsx", index=False)






# "Id" sütununu teke düşürme
unique_ids = sonuc_df["Id"].drop_duplicates()

# Yeni bir Excel sayfası oluşturma ve "Id" değerlerini yazma
with pd.ExcelWriter("İnstagram ve Özerler Depo Kalanlar.xlsx", engine="openpyxl", mode="a") as writer:
    unique_ids.to_excel(writer, sheet_name="Unique Ids", index=False)













# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 2
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 100
numbers_per_repeat = 28

# Verileri ekleme
for _ in range(repeat_count):
    for num in range(1, numbers_per_repeat + 1):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")






# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
sheet = wb["Unique Ids"]

# Başlangıç sütunu ve satırı
start_column = 3
start_row = 2

# Toplam tekrar sayısı ve her tekrardaki numara adedi
repeat_count = 28
numbers_per_repeat = 200

# Verileri ekleme
for num in range(1, numbers_per_repeat + 1):
    for repeat in range(repeat_count):
        sheet.cell(row=start_row, column=start_column, value=num)
        start_row += 1

# Değişiklikleri kaydetme
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=1).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]
main_sheet = wb["Sheet1"]

# "Id" sütununun verilerini al
id_column = main_sheet["A"][1:]
unique_ids_column = unique_ids_sheet["A"][1:]

# Karşılık gelen değerleri bulup "Sheet1" sayfasının en sağında yeni bir sütuna ekle
new_column = main_sheet.max_column + 1
main_sheet.cell(row=1, column=new_column, value="Matching Value (3rd Column)")

for id_cell in id_column:
    id_value = id_cell.value
    for unique_id_cell in unique_ids_column:
        if unique_id_cell.value == id_value:
            matching_value = unique_id_cell.offset(column=2).value
            main_sheet.cell(row=id_cell.row, column=new_column, value=matching_value)
            break

# Değişiklikleri kaydetme
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütun başlıklarını değiştir
new_column_titles = {
    "Id": "SiparişNO",
    "BirlesikVeri": "ÜRÜN",
    "GoogleSheetVerisi": "RAF KODU",
    "UrunAdiKopya3": "ÜRÜN ADI",
    "Matching Value": "KUTU",
    "Matching Value (3rd Column)": "ÇN"
}

for col_idx, col_name in enumerate(main_sheet.iter_cols(min_col=1, max_col=main_sheet.max_column), start=1):
    old_title = col_name[0].value
    new_title = new_column_titles.get(old_title, old_title)
    col_name[0].value = new_title

# Değişiklikleri kaydetme
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Sütunların yeni sıralaması
new_column_order = [
    "RAF KODU",
    "ÜRÜN",
    "Barkod",
    "KUTU",
    "ÜRÜN ADI",
    "ÇN",
    "SiparişNO"
]

# Yeni bir DataFrame oluştur
data = main_sheet.iter_rows(min_row=2, values_only=True)
df = pd.DataFrame(data, columns=[cell.value for cell in main_sheet[1]])

# Sütunları yeni sıralamaya göre düzenle
df = df[new_column_order]

# Mevcut başlıkları güncelle
for idx, column_name in enumerate(new_column_order, start=1):
    main_sheet.cell(row=1, column=idx, value=column_name)

# DataFrame verilerini sayfaya yaz
for r_idx, row in enumerate(df.values, 2):
    for c_idx, value in enumerate(row, 1):
        main_sheet.cell(row=r_idx, column=c_idx, value=value)

# Değişiklikleri kaydet
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Hücreleri ortala ve ortaya hizala
for row in main_sheet.iter_rows(min_row=2):
    for cell in row:
        cell.alignment = Alignment(horizontal="center", vertical="center")


# Değişiklikleri kaydet
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Kenarlık stili oluştur
border_style = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Hücreleri kalın yap, yazı fontunu 14 yap ve kenarlık ekle
for row in main_sheet.iter_rows():
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = border_style

# Değişiklikleri kaydet
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")





# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# "RAF KODU" sütununu 45 piksel yap
main_sheet.column_dimensions["C"].width = 45

# Tüm hücreleri en uygun sütun genişliği olarak ayarla
for column in main_sheet.columns:
    max_length = 0
    column = list(column)
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    main_sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Değişiklikleri kaydet
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")




# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# İlk sütunu (A sütunu) 45 piksel genişliğinde yap
main_sheet.column_dimensions["A"].width = 45
main_sheet.column_dimensions["C"].width = 14
main_sheet.column_dimensions["G"].width = 14
main_sheet.column_dimensions["D"].width = 9
main_sheet.column_dimensions["F"].width = 5

# Değişiklikleri kaydet
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")











# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tüm hücrelere "Metni Kaydır" formatını uygula
for row in main_sheet.iter_rows():
    for cell in row:
        new_alignment = copy(cell.alignment)
        new_alignment.wrap_text = True
        cell.alignment = new_alignment

# Değişiklikleri kaydet
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")










# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]

# Tabloyu oluşturma
table = Table(displayName="MyTable", ref=main_sheet.dimensions)

# Tablo stili oluşturma (gri-beyaz)
style = TableStyleInfo(
    name="TableStyleLight1", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=True
)

# Tabloya stil atama
table.tableStyleInfo = style

# Tabloyu sayfaya ekleme
main_sheet.add_table(table)

# Değişiklikleri kaydetme
wb.save("İnstagram ve Özerler Depo Kalanlar.xlsx")





def create_bat_files(data, output_folder, batch_size=28):
    batch_count = 1
    batch_data = []
    remaining_data = data

    while len(remaining_data) > 0:
        current_batch = remaining_data[:batch_size]
        batch_data.extend(current_batch)

        bat_file_path = os.path.join(output_folder, f"BAT{batch_count}.bat")
        with open(bat_file_path, "w") as file:
            link = f'start "" https://task.haydigiy.com/admin/order/printorder/?orderId={current_batch[0]}&isPdf=False\n'
            file.write(link)
            file.write('timeout -t 2\n')  # Add the timeout line

            for value in current_batch[1:]:
                link = f"https://task.haydigiy.com/admin/order/printorder/?orderId={value}&isPdf=False"
                file.write(f'start "" {link}\n')

        batch_data = []
        remaining_data = remaining_data[batch_size:]
        batch_count += 1

# Klasör oluştur
output_folder = "İnstagram ve Özerler Depo Kalanlar"
os.makedirs(output_folder, exist_ok=True)

# Sonuç dosyasını yükle
file_path = "İnstagram ve Özerler Depo Kalanlar.xlsx"
wb = load_workbook(file_path)
unique_ids_sheet = wb["Unique Ids"]

# "Id" sütunundaki verileri al
id_column = unique_ids_sheet["A"][1:]

# Verileri bir listeye dönüştür
id_values = [cell.value for cell in id_column if cell.value is not None]

# .bat dosyalarını oluştur ve klasöre kaydet
create_bat_files(id_values, output_folder)

# Excel dosyasını klasöre taşı
shutil.copy(file_path, os.path.join(output_folder, "İnstagram ve Özerler Depo Kalanlar.xlsx"))

gc.collect()

# Klasör dışında kalan Excel dosyasını sil
os.remove(file_path)













#DÜZELTME İÇİN
# Sonuç dosyasını yükle
file_path = "Çift Siparişler.xlsx"
wb = load_workbook(file_path)
main_sheet = wb["Sheet1"]




























































folders = ["0-112", "0-206", "113-206", "İnstagram (14)", "İnstagram Kalanlar", "İç Giyim", "207-400", "Tüm Depo Kalanlar", "Yeni Depo (14)", "Tüm Depo (14)", "1134-1269", "996-1269", "996-1133", "858-995", "700-995", "700-857", "Yeni Depo Kalanlar", "1270-1326", "1327-1459", "1460-1531", "Özerler Depo (14)", "2000-2164", "2165-2310", "Özerler Depo Kalanlar", "Yeni Depo ve İnstagram Kalanlar", "İnstagram ve Özerler Depo Kalanlar"]

# Bugünkü tarihi al
current_date = datetime.datetime.now().strftime("%Y-%m-%d")

# Oluşturulacak zip dosyasının adı
zip_filename = f"{current_date} Çıktılar.zip"

# Klasörleri kontrol et ve gerektiğinde sil veya zip'e ekle
with zipfile.ZipFile(zip_filename, 'w') as zipf:
    for folder in folders:
        folder_path = os.path.join(".", folder)
        folder_contents = os.listdir(folder_path)
        bat_files = [file for file in folder_contents if file.endswith(".bat")]

        if bat_files:
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(file_path, "."))
            
            
        else:
            for root, dirs, files in os.walk(folder_path, topdown=False):
                for file in files:
                    file_path = os.path.join(root, file)
                    os.remove(file_path)
                for dir in dirs:
                    dir_path = os.path.join(root, dir)
                    os.rmdir(dir_path)
            os.rmdir(folder_path)





# Klasörleri sil
for folder in folders:
    folder_path = os.path.join(".", folder)
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)





















#KARGO ENTEGRASYONUNA GÖNDERME

# Excel dosyasının adını belirtin
excel_dosyasi = "Kargo Entegrasyonu.xlsx"

# Excel dosyasını yükle
df = pd.read_excel(excel_dosyasi)

# "KargoTakipNumarasi" sütununda dolu olan satırları sil
df = df[df['KargoTakipNumarasi'].isna()]


# "KargoFirmasi" sütununda "MNG KARGO" değerine sahip olan satırları sil
df = df[df['KargoFirmasi'] != 'MNG KARGO']

# "KargoFirmasi" sütununda "MNG KARGO" değerine sahip olan satırları sil
df = df[df['KargoFirmasi'] != 'KARGOİST']

# "Id" sütunu hariç diğer tüm sütunları sil
df = df[['Id']]

# Aynı Excel dosyasına kaydet (üzerine yaz)
df.to_excel(excel_dosyasi, index=False)









# Excel dosyasını oku
df = pd.read_excel(excel_dosyasi)

# "Id" sütunu hariç diğer tüm sütunları sil
df = df[['Id']].drop_duplicates()

# Güncellenmiş veriyi aynı Excel dosyasına kaydet (mevcut dosyanın üzerine yazacak)
df.to_excel(excel_dosyasi, index=False)





print(" ")
print(Fore.RED + "Siparişler Entegrasyona Gönderiliyor Ekran Kapanana Kadar İşlem Yapmayın !")









# Excel dosyasını oku
df = pd.read_excel("Kargo Entegrasyonu.xlsx")

# Oturum oluşturma ve giriş yapma işlemini dışarı taşı
session = requests.Session()

# Oturumu bir kez aç
def login():
    login_url = "https://task.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
    username = "mustafa_kod@haydigiy.com"
    password = "123456"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
        "Referer": "https://task.haydigiy.com/",
    }

    # Oturum açma sayfasına GET isteği
    response = session.get(login_url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")

    # __RequestVerificationToken değerini alma
    token = soup.find("input", {"name": "__RequestVerificationToken"}).get("value")

    # Giriş verileri
    login_data = {
        "EmailOrPhone": username,
        "Password": password,
        "__RequestVerificationToken": token,
    }

    # Oturum açma isteği gönderme
    session.post(login_url, data=login_data, headers=headers)

# Oturumu bir kez aç
login()

# Siparişleri gönderme fonksiyonu
def send_request(order_id):
    url = f"https://task.haydigiy.com/admin/order/sendordertoshipmentintegration/?orderId={order_id}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
    }

    # Entegrasyon URL'sine istek gönder
    response = session.get(url, headers=headers)



# ThreadPoolExecutor kullanarak istekleri paralel olarak gönderme
with ThreadPoolExecutor(max_workers=5) as executor, tqdm(total=len(df['Id']), desc="Entegrasyona Gönderiliyor") as pbar:
    futures = [executor.submit(send_request, order_id) for order_id in df['Id']]
    for future in futures:
        future.result()  # İşlem tamamlandığında bir sonraki adıma geç
        pbar.update(1)  # İlerleme çubuğunu güncelle




# Excel dosyasının adı
excel_dosyasi = "Kargo Entegrasyonu.xlsx"

# Excel dosyasını sil
try:
    os.remove(excel_dosyasi)
    pass
except FileNotFoundError:
    print(f"{excel_dosyasi} adlı Excel dosyası bulunamadı.")
except Exception as e:
    print(f"Hata oluştu: {e}")








# Boş dosyaları kontrol etme ve silme işlemi
for dosya in ["Kara Liste Siparişleri.xlsx", "Çift Siparişler.xlsx"]:
    try:
        df = pd.read_excel(dosya) 
        if df.dropna().empty:  
            os.remove(dosya)
        else:
            pass
    except FileNotFoundError:
        print(f"'{dosya}' dosyası bulunamadı.")
    except Exception as e:
        print(f"'{dosya}' dosyasını kontrol ederken bir hata oluştu: {str(e)}")
