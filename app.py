from time import sleep
import pandas

from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


excel_data = pandas.read_excel('data.xlsx', sheet_name='data')

count = 0

driver = webdriver.Chrome(r'chromedriver.exe')
driver.get('https://web.whatsapp.com')
input('Scan QR Code Kemudian Tekan Enter.')
for column in excel_data['Kontak'].tolist():
    try:
        url = 'https://web.whatsapp.com/send?phone=' + str(excel_data['Kontak'][count]) + '&text=' + excel_data['Pesan'][0]
        sent = False
        driver.get(url)
        xpath_val = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p'
        wait = WebDriverWait(driver, 10)
        sleep(2)
        wait.until(EC.presence_of_element_located((By.XPATH, xpath_val))).send_keys(Keys.ENTER)
        sent = True
        sleep(5)
        print('Pesan Terkirim: ' + str(excel_data['Kontak'][count]))
        count = count + 1
    except Exception as e:
        print('Pesan Tidak Terkirim: ' + str(excel_data['Kontak'][count]) + str(e))
driver.quit()
print('Sampai Jumpa Kembali.')
