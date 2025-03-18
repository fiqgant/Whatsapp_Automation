import re
import pandas as pd
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# === Fungsi untuk Membersihkan Nomor Telepon ===
def clean_phone_number(phone_number):
    """
    Membersihkan nomor telepon:
    - Menghapus spasi, tanda +, tanda hubung (-), dan tanda kurung ()
    - Jika nomor diawali '0', ubah menjadi format internasional (misal: 0812 -> 62812)
    """
    phone_number = re.sub(r"[^\d]", "", phone_number)  # Hanya biarkan angka (0-9)
    
    if phone_number.startswith("0"):  
        phone_number = "62" + phone_number[1:]  # Ubah 0812 -> 62812

    return phone_number

# === Load Data dari Excel ===
file_path = "data.xlsx"  # Sesuaikan dengan path file
sheet_name = "data"

try:
    excel_data = pd.read_excel(file_path, sheet_name=sheet_name)
    print(f"Berhasil membaca file Excel. Jumlah kontak: {len(excel_data)}")
except Exception as e:
    print(f"Error membaca file Excel: {str(e)}")
    exit()

# === Setup WebDriver ===
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")  # Hindari deteksi bot
options.add_argument("--user-data-dir=./User_Data")  # Gunakan data login agar tidak scan QR ulang
options.add_argument("--profile-directory=Default")  # Gunakan profil default Chrome

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# === Buka WhatsApp Web ===
driver.get("https://web.whatsapp.com")
input("Scan QR Code di WhatsApp Web jika belum login, lalu tekan Enter untuk lanjut...")

# === Kirim Pesan ke Setiap Nomor ===
wait = WebDriverWait(driver, 20)

for index, row in excel_data.iterrows():
    phone_number = clean_phone_number(str(row["Kontak"]).strip())  # Bersihkan nomor telepon
    message = str(row["Pesan"]).strip().replace("\n", "%0A")  # Ganti \n dengan %0A untuk enter di WA

    url = f"https://web.whatsapp.com/send?phone={phone_number}&text={message}"
    driver.get(url)

    try:
        # Tunggu input pesan muncul
        xpath_input = '//div[@contenteditable="true"][@data-tab="10"]'
        input_box = wait.until(EC.presence_of_element_located((By.XPATH, xpath_input)))

        sleep(2)  # Beri waktu agar elemen stabil
        input_box.send_keys(Keys.ENTER)  # Tekan Enter untuk mengirim pesan

        print(f"✅ Pesan terkirim ke: {phone_number}")
        sleep(5)  # Tunggu beberapa detik sebelum lanjut ke nomor berikutnya

    except Exception as e:
        print(f"❌ Pesan gagal dikirim ke {phone_number}. Error: {str(e)}")

# === Selesai, Tutup Browser ===
driver.quit()
print("✅ Selesai mengirim semua pesan.")
