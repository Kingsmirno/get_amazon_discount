import sys
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

def open_excel(excel_file):
    try:
        workbook = openpyxl.load_workbook(excel_file)
        # Arbeitsblatt auswählen (z.B. das erste Arbeitsblatt)
        #return workbook
    except FileNotFoundError:
        print(f"Die Datei {excel_file} wurde nicht gefunden.")
        sys.exit(1)
    sheet = workbook.active
    return workbook,sheet

def check_amazon(ASIN):
    Amazon_link = 'https://www.amazon.de/gp/product/' + ASIN + '/ref=ppx_yo_dt_b_asin_title_o00_s00?ie=UTF8&th=1&psc=1'
    # Erstellen der Optionen für den Chrome-Browser
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36")
    
    # Starten des Chrome-Browsers mit den angepassten Optionen
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(Amazon_link)
    
    try:
        # Warten, bis das Element geladen ist
        driver.implicitly_wait(2)
        
        # Extrahieren des ganzen Teils und des Dezimalteils des Preises
        whole_price = driver.find_element(By.CLASS_NAME, 'a-price-whole').text
        decimal_price = driver.find_element(By.CLASS_NAME, 'a-price-fraction').text
        price =float( whole_price.strip() +"."+ decimal_price.strip().replace(',', '.'))
        # print (f"\n\t Ganzzahl:\t {whole_price} \n\t Dezimal:\t{decimal_price}")

        print(f"{ASIN}\taktueller Preis: {price}")  

        # extrahiere Coupon, falls vorhanden
        try:
            coupon_element =    driver.find_element(By.XPATH, "//*[contains(@id, 'couponText')]")
            coupon =            coupon_element.text
            if "%" in coupon:
                discount =          int(coupon[:3].strip(" %-")) if coupon else 0
                discounted_price =  price * (1 - discount / 100)
            elif "€" in coupon:
                discount =          int(coupon[:3].strip(" €-")) if coupon else 0
                discounted_price =  price - discount
            print(f"\tdiscount: **{discount}% **")
            print(f"\tdiscounted price = : {discounted_price}")  
  # Debugging-Ausgabe
        except:
            discount = 0
            discounted_price = price
            print("kein Coupon gefunden")

        return price, discount,discounted_price
    except Exception as e:
        print(f"Fehler beim Abrufen des Preises: {e}")
        return 5000, 0, discounted_price
    finally:
        driver.quit()

def get_data_from_excel(sheet, row):#->ASIN, Product_name, consideration_price, old_smallest__price, old_discount, old_Amazon_Price, Amazon_link:
    ASIN =                  sheet[f'B{row}'].value
    Product_name =          sheet[f'C{row}'].value if sheet[f'C{row}'].value else "Produktname nicht gefunden"
    consideration_price =   (sheet[f'H{row}'].value) if sheet[f'H{row}'].value else 0
    old_smallest__price =   (sheet[f'I{row}'].value) if sheet[f'I{row}'].value else 0
    old_discount =          (sheet[f'J{row}'].value) if sheet[f'J{row}'].value else 0
    old_Amazon_Price =      (sheet[f'K{row}'].value) if sheet[f'K{row}'].value else 0
    Last_update         =   (sheet[f'L{row}'].value)

    #print(f"ASIN: {type(ASIN)}{ASIN}\nProduct_name: {Product_name}\nconsideration_price:{type(consideration_price)}, {consideration_price}\nold_smallest__price:{type(old_smallest__price)} {old_smallest__price}\nold_discount:{type(old_discount)}, {old_discount}\nold_Amazon_Price:{type(old_Amazon_Price)}, {old_Amazon_Price}")
    return ASIN, Product_name, consideration_price, old_smallest__price, old_discount, old_Amazon_Price, Last_update

def break_after_x(x):
    if count % x == 0:
        print("Count ist durch {x} teilbar. Stoppen.")
        return True

def main():
    #Öffnet Excel-Datei
    excel_file =          '/Users/ulrich/Documents/Dropbox/Amazon/Tesprodukte_aktuell.xlsx'
    workbook,sheet =      open_excel(excel_file)

    count = 0
    
    for row in range(2, sheet.max_row + 1):
        ASIN, Product_name, consideration_price, old_smallest__price, old_discount, old_Amazon_Price, Last_update= get_data_from_excel(sheet, row)
        if ASIN and Last_update is None: #falls Spalte ASIN hat aber noch nicht gecheckt wurde
            price, discount, discouted_price = check_amazon(ASIN)
            sheet[f'I{row}']    = x =  min([old_smallest__price if old_smallest__price >0 else 5000, consideration_price,discouted_price, price])
            sheet[f'J{row}']    = max([old_discount, discount])
            sheet[f'K{row}']    = min([old_Amazon_Price, price])
            sheet[f'L{row}']    = datetime.now()
            sheet[f'M{row}']    = discouted_price
            workbook.save(excel_file)
            count += 1
            print(x)
            print ("Count: ", count,"\n")
            repeat = 4
            if count % repeat == 0:
                print(f"Count ist durch {repeat} teilbar. Stoppen.")
                break

    workbook.save(excel_file)
   
if __name__ == "__main__":
    main()