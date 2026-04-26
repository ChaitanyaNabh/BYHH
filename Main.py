from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import os
import glob
import pandas as pd
import shutil
import urllib.request
import io

# ====== CONFIG ======
KINSER_URL = "https://kinnser.net/login.cfm"
USERNAME = "reporting@carebridgehh.com"
PASSWORD = "Carebridge@2026"
DOWNLOAD_DIR = r"G:\My Drive\BetterYou\Billing\Billed\Billed_Report"
DOWNLOADS_DIR1 = os.path.join(os.path.expanduser("~"), "Downloads")
DEST_DIR = r"G:\My Drive\BetterYou\Billing\Billed\Billed_Auto_Report"

# Clean folders
for folder in [DOWNLOAD_DIR, DEST_DIR]:
    if os.path.exists(folder):
        for file in os.listdir(folder):
            try:
                os.remove(os.path.join(folder, file))
            except:
                pass

# ====== HEADLESS OPTIONS ======
options = Options()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

# ====== DRIVER ======
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

driver.execute_cdp_cmd(
    "Page.setDownloadBehavior",
    {"behavior": "allow", "downloadPath": DOWNLOAD_DIR}
)

driver.execute_script("document.body.classList.remove('cdk-overlay-open');")

wait = WebDriverWait(driver, 20)

# ====== LOGIN ======
driver.get(KINSER_URL)

wait.until(EC.visibility_of_element_located((By.ID, "username"))).send_keys(USERNAME)
wait.until(EC.visibility_of_element_located((By.ID, "password"))).send_keys(PASSWORD)
wait.until(EC.element_to_be_clickable((By.ID, "login_btn"))).click()

try:
    WebDriverWait(driver, 5).until(EC.alert_is_present())
    driver.switch_to.alert.accept()
except:
    pass

wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

# ====== NAV ======
wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Go To"))).click()
wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Financials"))).click()

wait.until(EC.element_to_be_clickable((By.XPATH, "//div[normalize-space()='A/R']"))).click()

# ====== BRANCH ======
branch_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-1")))
driver.execute_script("arguments[0].click();", branch_dropdown)

branch_option = wait.until(EC.element_to_be_clickable((
    By.XPATH,
    "//mat-option//span[normalize-space(text())='Better You Home Health LLC']"
)))
driver.execute_script("arguments[0].click();", branch_option)

driver.find_element(By.TAG_NAME, "body").click()
time.sleep(0.5)

# ====== INSURANCE TYPE ======
insurance_type_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "mat-select-2")))

driver.find_element(By.TAG_NAME, "body").click()
time.sleep(0.5)
driver.execute_script("arguments[0].click();", insurance_type_dropdown)

first_checkbox_input = wait.until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="mat-mdc-checkbox-6-input"]'))
)
driver.execute_script("arguments[0].click();", first_checkbox_input)

driver.find_element(By.TAG_NAME, "body").click()
time.sleep(0.5)

# ====== INSURANCE NAME ======
insurance_name_dropdown = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//multiple-selection-select-all[@id='insuranceKeys']//mat-select"))
)
driver.execute_script("arguments[0].click();", insurance_name_dropdown)

insurance_name_checkbox = wait.until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="mat-mdc-checkbox-7-input"]'))
)
driver.execute_script("arguments[0].click();", insurance_name_checkbox)

driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
time.sleep(1)

# ====== UNBILLED ======
checkbox_label = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//label[@for='unbilledCheckBox-input']"))
)
driver.execute_script("arguments[0].click();", checkbox_label)

# ====== BILLED DATE ======
billed_date_radio = wait.until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "input[value='Billed Date']"))
)
driver.execute_script("arguments[0].click();", billed_date_radio)

# ====== APPLY ======
apply_filters_btn = wait.until(EC.element_to_be_clickable((By.ID, "applyButton")))
driver.execute_script("arguments[0].click();", apply_filters_btn)

# ====== EXPORT ======
export_btn = WebDriverWait(driver, 60).until(
    lambda d: (elem := d.find_element(By.ID, "accountsReceivable_btnexport"))
    and elem.get_attribute("href") not in [None, ""]
    and elem
)

driver.execute_script("arguments[0].click();", export_btn)

time.sleep(5)

driver.quit()

download_folder = r"G:\My Drive\BetterYou\Billing\Billed\Billed_Report"

timeout = 60
start_time = time.time()
latest_file = None

while time.time() - start_time < timeout:
    list_of_files = glob.glob(os.path.join(download_folder, "*.xlsx"))
    if list_of_files:
        latest_file = max(list_of_files, key=os.path.getctime)
        if not latest_file.endswith(".crdownload"):
            break
    time.sleep(1)

df = pd.read_excel(latest_file, skiprows=4)

df = df.sort_values(by="Billed Date", ascending=True).reset_index(drop=True)

df = df[[
    "Billed Date",
    "Payer",
    "Patient Last Name",
    "Patient First Name",
    "MRN",
    "Billing Period Start Date",
    "Billing Period End Date",
    "Balance",
    "Claim Number"
]]

df["Billing Period Start Date"] = df["Billing Period Start Date"].dt.strftime("%m-%d-%Y")
df["Billing Period End Date"] = df["Billing Period End Date"].dt.strftime("%m-%d-%Y")
df["Billed Date"] = df["Billed Date"].dt.strftime("%m-%d-%Y")

df3 = df.copy()

sheet_id = "1C0PbKe0f-CaHqPVN9yzUXle3Js1vQJlrBx6pemTEP9E"
tab_gid = "842092887"

url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={tab_gid}"

response = urllib.request.urlopen(url)
data = response.read().decode("utf-8")

Billed = pd.read_csv(io.StringIO(data))

Billed = Billed.rename(columns={
    'Last Name': 'Patient Last Name',
    'First Name': 'Patient First Name'
})

df3["MRN"] = df3["MRN"].astype(str)
Billed = Billed[Billed["MRN"].notna()]
Billed["MRN"] = Billed["MRN"].astype(int).astype(str)

df3["Claim Number"] = pd.to_numeric(df3["Claim Number"], errors="coerce").astype("Int64")
Billed["Claim Number"] = pd.to_numeric(Billed["Claim Number"], errors="coerce").astype("Int64")

mask_not_in_df3 = (~Billed["Claim Number"].isin(df3["Claim Number"]))
Billed.loc[mask_not_in_df3, 'Status'] = "Resolved"

rows_to_add = df3.loc[~df3["Claim Number"].isin(Billed["Claim Number"])]
Billed = pd.concat([Billed, rows_to_add], ignore_index=True)

save_folder = r"G:\My Drive\BetterYou\Billing\Billed\Billed_Final"
os.makedirs(save_folder, exist_ok=True)

from datetime import datetime

output_file = os.path.join(
    save_folder,
    f"Final_Billed_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
)

Billed.to_excel(output_file, index=False)

print(f"💾 Final file saved at: {output_file}")