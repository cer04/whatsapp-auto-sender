# -*- coding: utf-8 -*-
import pandas as pd
import time
import pyperclip
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# === CONFIGURATION ===
excel_path = r"C:\Users\Cyber\Downloads\29-4_9 (1).xlsx" 
start_row = 1   # zero-based index
edgedriver_path = r"C:\Users\Cyber\Desktop\محاسبة\edgedriver_win64\msedgedriver.exe"  # <-- update

# === Load Excel Data ===
df = pd.read_excel(excel_path)
df = df.iloc[start_row:].reset_index(drop=True)

print(f"Loaded {len(df)} rows from Excel (starting at row index {start_row}).")

# === Setup Edge WebDriver ===
options = webdriver.EdgeOptions()
options.add_argument("--user-data-dir=C:\\temp\\whatsapp_edge_profile")  # reuse existing session
options.add_argument("--disable-features=IsolateOrigins,site-per-process")
options.add_experimental_option('excludeSwitches', ['enable-logging'])

service = webdriver.edge.service.Service(executable_path=edgedriver_path)
driver = webdriver.Edge(service=service, options=options)

# === Open WhatsApp Web ===
driver.get("https://web.whatsapp.com/")
print("🔓 Opening WhatsApp Web in Edge. If this is the first run, scan the QR code in the browser profile window.")

try:
    WebDriverWait(driver, 80).until(
        EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
    )
    print("✅ WhatsApp web loaded.")
except TimeoutException:
    print("❌ Timeout waiting for WhatsApp to load. Exiting.")
    driver.quit()
    raise SystemExit


# === Helpers ===
def clear_search_box(el):
    """Clear WhatsApp search box safely."""
    try:
        el.click()
        el.send_keys(Keys.CONTROL + 'a')
        el.send_keys(Keys.BACKSPACE)
    except Exception:
        try:
            el.send_keys(Keys.BACKSPACE * 20)
        except Exception:
            pass

def human_delay(a=0.5, b=2.5):
    """Randomized sleep to simulate human behavior."""
    time.sleep(random.uniform(a, b))

def send_message(name, message_text):
    try:
        # Find search box
        search_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
        )
        clear_search_box(search_box)
        human_delay(0.7, 2.0)

        # Type contact name
        search_box.send_keys(name)
        human_delay(1.5, 3.5)

        # Click contact
        try:
            contact = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.XPATH, f'//span[@title="{name}"]'))
            )
            contact.click()
            human_delay(1.0, 2.5)
        except TimeoutException:
            print(f"❌ Contact '{name}' not found.")
            return

        # Find message box
        message_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )

        # Copy + Paste message
        pyperclip.copy(message_text)
        message_box.click()
        human_delay(0.4, 1.0)

        ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

        # Random pause before sending (like rereading the message)
        human_delay(1.2, 4.5)

        # Send
        message_box.send_keys(Keys.ENTER)
        print(f"✅ Sent to {name}")

        # Sometimes take a longer break
        if random.random() < 0.25:
            pause = random.randint(8, 20)
            print(f"⏳ Taking a human-like pause of {pause} seconds...")
            time.sleep(pause)

    except Exception as e:
        print(f"⚠️ Error sending to {name}: {e}")


# === Iterate rows and send messages ===
for idx, row in df.iterrows():
    name_raw = row.iloc[0]
    produced_raw = row.iloc[1] if len(row) > 1 else None
    delivered_raw = row.iloc[2] if len(row) > 2 else None
    required_raw = row.iloc[4] if len(row) > 4 else None

    print(f"\nRow {idx + start_row + 1} -> name: {name_raw}, produced: {produced_raw}, delivered: {delivered_raw}, required: {required_raw}")

    if pd.isna(name_raw) or pd.isna(produced_raw) or pd.isna(delivered_raw) or pd.isna(required_raw):
        print(f"⏭️ Skipping row {idx + start_row + 1} due to missing data.")
        continue

    try:
        produced_int = int(round(float(produced_raw)))
        delivered_int = int(round(float(delivered_raw)))
        required_float = float(required_raw)
    except Exception as e:
        print(f"⚠️ Number conversion error for row {idx + start_row + 1}: {e}. Skipping.")
        continue

    name = str(name_raw).strip()

    # === Choose message template ===
    if required_float < 0:
        message_text = f"""على قروب ( المملكه )

انتجت      {produced_int}                   
استهلكت      {delivered_int}               
بترتب عليك ( {abs(required_float)} 🌹 ) 💴💷💵💶

نقتطع اول (2) اوردرات في خانه الانتاج 
وبعد ال 20 اول (2) 
وبعد ال 60 اول (2)  ......الخ

نقتطع أول شيء ما لنا من إنتاجك ثم نخصم إستهلاكك 💯

بستنى تحولهم خلال الـ 24 ساعه 
آخر مهله يوم السبت قبل الساعه 7 مساءً 🚫 
او الاضافه راحت بعد 48 ساعه 🚫

بتقدر تحولنا على .......
كليك بنك الاتحاد :  
SYF82  

أو  
0779060760  
اورنج موني  

بأسم ( صفاء يوسف نصر )

وابعتلي صوووورة عن الحواله مع وقت ارسالها 📥 🤍❤ 🤝
لا نعترف ب صوره المتأخره ❌"""
    else:
        message_text = f"""على قروب ( المملكه )

انتجت      {produced_int}                   
استهلكت      {delivered_int}               
بترتب لك ( {required_float} 🌹 ) 💴💷💵💶

نقتطع اول (2) اوردرات في خانه الانتاج 
وبعد ال 20 اول (2) 
وبعد ال 60 اول (2)  ......الخ

نقتطع أول شيء ما لنا من إنتاجك ثم نخصم إستهلاكك 💯

بستنى تحولهم خلال الـ 24 ساعه 
آخر مهله يوم السبت قبل الساعه 7 مساءً 🚫 
او الاضافه راحت بعد 48 ساعه 🚫

بتقدر تحولنا على .......
كليك بنك الاتحاد :  
SYF82  

أو  
0779060760  
اورنج موني  

بأسم ( صفاء يوسف نصر )

وابعتلي صوووورة عن الحواله مع وقت ارسالها 📥 🤍❤ 🤝
لا نعترف ب صوره المتأخره ❌"""

    # === Send message ===
    send_message(name, message_text)

# Done
print("\nAll done. Closing browser.")
driver.quit()