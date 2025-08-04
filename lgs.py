from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import time
import pandas as pd
import openpyxl
from io import StringIO
import pandas as pd
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox, QLabel, QLineEdit, QVBoxLayout, QPushButton, QDialog
import sys
from selenium.webdriver.support.ui import Select
import logging
import os

chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)

logging.basicConfig(filename='error_log.txt', level=logging.ERROR)
def save_progress(index):
    with open('progress.txt', 'w') as f:
        f.write(str(index))
def load_progress():
    try:
        with open('progress.txt', 'r') as f:
            return int(f.read())
    except FileNotFoundError:
        return 0
def delete_progress():
    try:
        os.remove('progress.txt')
    except FileNotFoundError:
        pass
def dosya_ac():
    global write_file, excel_file, read_sheet_name, write_sheet_name
    ogrenci_dosyası_secme_mesajı()
    options = QFileDialog.Options()
    excel_file, _ = QFileDialog.getOpenFileName(None, "Dosya Seç", "", "All Files (*)", options=options)
    read_workbook = openpyxl.load_workbook(excel_file)
    read_sheet_name = read_workbook.sheetnames[0]
    sonuc_dosyası_secme_mesajı()
    write_file, _ = QFileDialog.getOpenFileName(None, "Dosya Seç", "", "All Files (*)", options=options)
    write_workbook = openpyxl.load_workbook(write_file)
    write_sheet_name = write_workbook.sheetnames[0]

def tercih_yapmayan_ogrenci_veri():
    birthday = str(bDay) + "." + str(bMonth) + "." + str(bYear)
    data = {
        '0': ["T.C. Kimlik No", "Adı Soyadı", "Doğum Tarihi", "Sonuç"],
        '1': [str(id), name, birthday, 'T.C. Kimlik Numaranızı/Doğum Tarihinizi Yanlış Girdiniz veya Tercih Başvurunuz Bulunmamaktadır!']
    }
    data_frame = pd.DataFrame(data)
    tablo_verileri.append(data_frame)

def tablo_verilerini_al():
    tablo_elementleri = driver.find_elements(By.XPATH, "//table")
    for table in tablo_elementleri:
        tablo_icerigi = table.get_attribute('outerHTML')
        data_frame = pd.read_html(StringIO(tablo_icerigi))[0]
        tablo_verileri.append(data_frame)
    return pd.concat(tablo_verileri).drop_duplicates()


def tablo_verilerini_yaz():
    concatenated_df = pd.concat(tablo_verileri)
    write_dolu_satir_sayisi = read_excel_row_count(write_file, write_sheet_name)
    # Eğer write_dolu_satir_sayisi 0 ise, startrow 0 olur, aksi takdirde startrow write_dolu_satir_sayisi olur
    startrow = write_dolu_satir_sayisi if write_dolu_satir_sayisi == 0 else write_dolu_satir_sayisi + 1
    print(startrow)
    
    with pd.ExcelWriter(write_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        concatenated_df.to_excel(writer, sheet_name=write_sheet_name, index=False, startrow=startrow, header=write_dolu_satir_sayisi==0)

    # Clear the list after writing to the file
    tablo_verileri.clear()    

def yeni_sorgu():
    yeni_sorgu_selector = (By.PARTIAL_LINK_TEXT, "Yeni Sorgu")
    if is_element_present(driver, *yeni_sorgu_selector):
        yeni_sorgu = driver.find_element(By.PARTIAL_LINK_TEXT, "Yeni Sorgu")
        yeni_sorgu.click()
    else:
        driver.back()
        time.sleep(1)


def find_input_by_attribute(driver, attribute, value):
    return driver.find_element(By.CSS_SELECTOR, f'input[{attribute}="{value}"]')

def find_element_with_fallbacks(driver, selectors_list, element_name):
    """
    Birden fazla selector deneyerek elementi bulmaya çalışır
    """
    for selector_type, selector_value in selectors_list:
        try:
            element = driver.find_element(selector_type, selector_value)
            print(f"{element_name} bulundu: {selector_type} = '{selector_value}'")
            return element
        except NoSuchElementException:
            continue
    
    # Hiçbiri bulunamazsa hata fırlat
    print(f"HATA: {element_name} bulunamadı. Denenen selectors:")
    for selector_type, selector_value in selectors_list:
        print(f"  - {selector_type}: '{selector_value}'")
    
    # Sayfanın kaynak kodunu kontrol etmek için
    print("Sayfa kaynağında TC kimlik ile ilgili inputlar:")
    page_source = driver.page_source.lower()
    if 'kimlik' in page_source:
        print("✓ 'kimlik' kelimesi sayfada bulundu")
    if 'tc' in page_source:
        print("✓ 'tc' kelimesi sayfada bulundu")
    
    raise NoSuchElementException(f"{element_name} için hiçbir selector çalışmadı")

def fill_date_dropdowns():
    """
    Tarih dropdown'larını güvenli şekilde doldur
    """
    dropdown_selectors = [
        ("GUN", str(bDay), "Gün"),
        ("AY", str(bMonth), "Ay"), 
        ("YIL", str(bYear), "Yıl")
    ]
    
    for name_value, select_value, field_name in dropdown_selectors:
        try:
            # Farklı name değerleri dene
            possible_names = [name_value, name_value.lower(), name_value.upper()]
            dropdown_found = False
            
            for possible_name in possible_names:
                try:
                    dropdown = Select(driver.find_element(By.NAME, possible_name))
                    dropdown.select_by_value(select_value)
                    print(f"{field_name} seçildi: {select_value}")
                    dropdown_found = True
                    break
                except NoSuchElementException:
                    continue
            
            if not dropdown_found:
                # ID ile de dene
                try:
                    dropdown = Select(driver.find_element(By.ID, name_value))
                    dropdown.select_by_value(select_value)
                    print(f"{field_name} (ID ile) seçildi: {select_value}")
                except NoSuchElementException:
                    print(f"UYARI: {field_name} dropdown'u bulunamadı")
                    
        except Exception as e:
            print(f"UYARI: {field_name} seçilirken hata: {e}")

def debug_page_elements():
    """
    Sayfadaki tüm input ve select elementlerini listele
    """
    print("\n=== SAYFA ELEMENT ANALIZI ===")
    
    # Tüm input elementleri
    inputs = driver.find_elements(By.TAG_NAME, "input")
    print(f"Toplam input sayısı: {len(inputs)}")
    
    for i, inp in enumerate(inputs):
        try:
            inp_type = inp.get_attribute("type") or "text"
            inp_id = inp.get_attribute("id") or "N/A"
            inp_name = inp.get_attribute("name") or "N/A"
            inp_placeholder = inp.get_attribute("placeholder") or "N/A"
            inp_class = inp.get_attribute("class") or "N/A"
            
            print(f"Input {i+1}: type={inp_type}, id={inp_id}, name={inp_name}, placeholder={inp_placeholder}, class={inp_class}")
        except:
            print(f"Input {i+1}: Element analiz edilemedi")
    
    # Tüm select elementleri
    selects = driver.find_elements(By.TAG_NAME, "select")
    print(f"\nToplam select sayısı: {len(selects)}")
    
    for i, sel in enumerate(selects):
        try:
            sel_id = sel.get_attribute("id") or "N/A"
            sel_name = sel.get_attribute("name") or "N/A"
            sel_class = sel.get_attribute("class") or "N/A"
            
            print(f"Select {i+1}: id={sel_id}, name={sel_name}, class={sel_class}")
        except:
            print(f"Select {i+1}: Element analiz edilemedi")
    
    print("=== ANALIZ TAMAMLANDI ===\n")

def giris_elemanları():
    try:
        # TC Kimlik No için farklı selector alternatifleri
        tc_selectors = [
            (By.XPATH, '//input[contains(@placeholder, "T.C. Kimlik")]'),
            (By.XPATH, '//input[contains(@placeholder, "Kimlik")]'),
            (By.XPATH, '//input[contains(@placeholder, "T.C.")]'),
            (By.XPATH, '//input[@id="TC_KIMLIK_NO"]'),
            (By.XPATH, '//input[@id="TCNO"]'),
            (By.XPATH, '//input[@id="ADAY_NO"]'),
            (By.XPATH, '//input[@name="TC_KIMLIK_NO"]'),
            (By.XPATH, '//input[@name="TCNO"]'),
            (By.XPATH, '//input[contains(@class, "tc")]'),
            (By.CSS_SELECTOR, 'input[placeholder*="Kimlik"]'),
            (By.CSS_SELECTOR, 'input[placeholder*="T.C"]'),
        ]
        
        # Okul No için farklı selector alternatifleri
        okul_selectors = [
            (By.XPATH, '//input[contains(@placeholder, "Okul")]'),
            (By.XPATH, '//input[contains(@placeholder, "OKUL")]'),
            (By.XPATH, '//input[@id="OKULNO"]'),
            (By.XPATH, '//input[@id="GUVENLIKNUMARASI"]'),
            (By.XPATH, '//input[@name="OKULNO"]'),
            (By.XPATH, '//input[@name="OKUL_NO"]'),
            (By.XPATH, '//input[contains(@class, "okul")]'),
            (By.CSS_SELECTOR, 'input[placeholder*="Okul"]'),
            (By.CSS_SELECTOR, 'input[placeholder*="OKUL"]'),
        ]
        
        # Elementleri bul
        tcInput = find_element_with_fallbacks(driver, tc_selectors, "TC Kimlik Input")
        okulNoInput = find_element_with_fallbacks(driver, okul_selectors, "Okul No Input")
        
        # Verileri temizle ve gir
        tcInput.clear()
        tcInput.send_keys(str(id))
        print(f"TC Kimlik girildi: {id}")
        
        okulNoInput.clear()
        okulNoInput.send_keys(str(okulNo))
        print(f"Okul No girildi: {okulNo}")
        
        # Dropdown'lar için güvenli seçim
        fill_date_dropdowns()
        
    except Exception as e:
        print(f"HATA - Giriş elemanları doldurulurken: {e}")
        print(f"Mevcut URL: {driver.current_url}")
        print(f"Sayfa başlığı: {driver.title}")
        
        # Hata durumunda sayfanın HTML'ini dosyaya kaydet
        with open('hata_sayfa.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        print("Sayfa kaynağı 'hata_sayfa.html' dosyasına kaydedildi")
        
        raise
def uyarı_mesaj_guvenlik_metin():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText("Girdiğiniz güvenlik kodunda bir hata oluştu! sadece güvenlik kodunu elle giriniz ve sonra uyarı mesajındaki tamama basınız")
    msg.setWindowTitle("Hata Mesajı")
    msg.exec_()

def bilgi_mesajı():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText("Güvenlik kodunu giriniz ve sonra uyarı mesajındaki tamama basınız")
    msg.setWindowTitle("Dikkat")
    msg.exec_()

def sınava_girmeyen_ogrenci_mesajı():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText(f"{name} adlı öğrenci sınava girmemiştir veya tercih başvurusunda bulunmamıştır. Tamam basarak sistemin işlemesine devam ediniz")
    msg.setWindowTitle("Dikkat")
    msg.exec_()

def bitis_mesajı():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText("Öğrencilerin sonuçlarını alma işlemi tamamlanmıştır.")
    msg.setWindowTitle("Dikkat")
    msg.exec_()

def ogrenci_dosyası_secme_mesajı():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText("Lgs sonuçları çekilecek olan öğrencilerin excel dosyasını açılan ekrandan seçiniz")
    msg.setWindowTitle("Dikkat")
    msg.exec_()

def sonuc_dosyası_secme_mesajı():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText("Lgs sonuçlarının kaydedileceği excel dosyasını açılan ekrandan seçiniz")
    msg.setWindowTitle("Dikkat")
    msg.exec_()

def is_element_present(driver, by, selector):
    try:
        driver.find_element(by, selector)
        return True
    except NoSuchElementException:
        return False

def read_excel_row(file_path, sheet_name, row_index):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    row_data = df.iloc[row_index].values.tolist()
    return row_data

def read_excel_row_count(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    dolu_satir_sayisi = df.shape[0]
    return dolu_satir_sayisi

def başarılı_sorgu_sonuc():
    tablo_verilerini_al()
    yeni_sorgu()

def get_input():
    dialog = QDialog()
    dialog.setWindowTitle('LGS SONUÇ URL')

    label = QLabel('Sonuçların alınacağı url giriniz:')
    input_field = QLineEdit()

    def on_ok_clicked():
        global input_value
        input_value = input_field.text()
        dialog.accept()

    ok_button = QPushButton('Tamam')
    ok_button.clicked.connect(on_ok_clicked)

    layout = QVBoxLayout()
    layout.addWidget(label)
    layout.addWidget(input_field)
    layout.addWidget(ok_button)

    dialog.setLayout(layout)
    dialog.setFixedSize(300, 150)

    dialog.exec_()
    return input_value

def sonucları_al():
    global name, bDay, bMonth, bYear, id, okulNo
    start_index = load_progress()
    
    for i in range(start_index, dolu_satir_sayisi):
        try:
            row_index = i  
            row_data = read_excel_row(excel_file, read_sheet_name, row_index)
            name, bDay, bMonth, bYear, id, okulNo = row_data[:6]
            
            if int(bDay) < 10:
                bDay = "0" + str(bDay)
            if int(bMonth) < 10:
                bMonth = "0" + str(bMonth)
            
            try:
                giris_elemanları()
            except NoSuchElementException:
                print("Element bulunamadı - Debug yapılıyor...")
                debug_page_elements()
                raise
            guvenlık_kodu_selector = (By.XPATH, '//input[(@id="GUVENLIKKODU") or (@id="gkodu")]')
            hata_kodu_selector = (By.ID, "hata")
            yeni_sayfada_hata_kodu_selector = (By.XPATH, "//p[@align='center']")
            
            tamam_button=driver.find_element(By.NAME, "Submit")
            
            if is_element_present(driver, *guvenlık_kodu_selector):
                bilgi_mesajı()
                tamam_button.click()
                time.sleep(2)
                
                if is_element_present(driver, *hata_kodu_selector):
                    hatakodutext = driver.find_element(By.XPATH, "//*[@id='hata']")
                    if hatakodutext.text == "Güvenlik Kodunu yanlış girdiniz!":
                        uyarı_mesaj_guvenlik_metin()
                        try:
                            giris_elemanları()
                        except NoSuchElementException:
                            print("Element bulunamadı - Debug yapılıyor...")
                            debug_page_elements()
                            raise
                        time.sleep(1)
                        tamam_button=driver.find_element(By.NAME, "Submit")
                        tamam_button.click()
                        if is_element_present(driver, *hata_kodu_selector):
                            sınava_girmeyen_ogrenci_mesajı()
                            tercih_yapmayan_ogrenci_veri()
                            continue
                    else:
                        sınava_girmeyen_ogrenci_mesajı()
                        tercih_yapmayan_ogrenci_veri()
                        continue
                
                if is_element_present(driver, *yeni_sayfada_hata_kodu_selector):
                    yeni_hatakodutext = driver.find_element(By.XPATH, '//p[1]')
                    if yeni_hatakodutext.text in ["T.C. Kimlik Numaranızı/Doğum Tarihinizi Yanlış Girdiniz veya Tercih Başvurunuz Bulunmamaktadır!", "T.C. Kimlik Numaranızı veya Doğum Tarihinizi Yanlış Girdiniz!"]:
                        sınava_girmeyen_ogrenci_mesajı()
                        tercih_yapmayan_ogrenci_veri()
                        yeni_sorgu()
                        continue
                    else:
                        yeni_sorgu()
                        try:
                            giris_elemanları()
                        except NoSuchElementException:
                            print("Element bulunamadı - Debug yapılıyor...")
                            debug_page_elements()
                            raise
                        uyarı_mesaj_guvenlik_metin()
                        tamam_button=driver.find_element(By.NAME, "Submit")
                        tamam_button.click()
                        if is_element_present(driver, *yeni_sayfada_hata_kodu_selector):
                            sınava_girmeyen_ogrenci_mesajı()
                            tercih_yapmayan_ogrenci_veri()
                            yeni_sorgu()
                            continue
                
                başarılı_sorgu_sonuc()
            
            else:
                tamam_button.click()
                time.sleep(2)
                if is_element_present(driver, *hata_kodu_selector):
                    sınava_girmeyen_ogrenci_mesajı()
                    tercih_yapmayan_ogrenci_veri()
                    continue
                if is_element_present(driver, *yeni_sayfada_hata_kodu_selector):
                    yeni_sorgu()
                    sınava_girmeyen_ogrenci_mesajı()
                    tercih_yapmayan_ogrenci_veri()
                    continue
                başarılı_sorgu_sonuc()
            
            save_progress(i + 1)  # İlerlemeyi her adımda kaydet
            
            if (i + 1) % 5 == 0:
                tablo_verilerini_yaz()
                
        except Exception as e:
            logging.error(f"Hata oluştu (öğrenci indeksi {i}): {str(e)}")
            save_progress(i)  # Hata durumunda son başarılı konumu kaydet
            tablo_verilerini_yaz()  # Hata durumunda mevcut verileri kaydet
            raise  # Hatayı yeniden fırlat
    
    tablo_verilerini_yaz()  # Son kez kaydet

   
if __name__ == "__main__":
    app = QApplication(sys.argv)
    get_input()
    driver = webdriver.Chrome(options=chrome_options)  # options ekleyin
    driver.get(input_value)
    #driver.maximize_window()
    driver.implicitly_wait(1)
    dosya_ac()
    dolu_satir_sayisi = read_excel_row_count(excel_file, read_sheet_name)
    tablo_verileri = []
    try:
        
        sonucları_al()
        delete_progress()
    except Exception as e:
        logging.error(f"Hata oluştu: {str(e)}")
    bitis_mesajı()
    driver.quit()
    app.quit()
    sys.exit(app.exec_())
   