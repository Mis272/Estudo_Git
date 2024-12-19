from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))

driver.get("https://aki.blueservice.app/index.php")

try:
    username = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='username']"))
    )
    username.send_keys("KARLA.ADRIELLY")
    username.send_keys(Keys.TAB)
    password = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='passwd']"))
    )
    password.send_keys("Banco@35")
    password.send_keys(Keys.RETURN)
    
except Exception as e:
    print(f"Erro ao realizar login: {e}")
    driver.quit()
try:
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[1]/td/div/div[2]/ul/li[1]/a"))).click()
    time.sleep(5)
except:
    pass

menu_1 = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='div_menu_esquerdo_principal']/div[1]/span[1]"))
)
menu_1.click()

menu_2 = driver.find_element(By.XPATH, "//*[@id='div_menu_esquerdo_principal']/div[3]")
menu_2.click()

menu_3 = driver.find_element(By.XPATH, "/html/body/div[5]/div[3]/div[2]/a/span")
menu_3.click()

dropdown = driver.find_element(By.XPATH, "//*[@id='frmSelectEsteira']/select")
dropdown.click()

first_option = dropdown.find_element(By.XPATH, ".//option[1]")
first_option.click()

table_xpath = "/html/body/div[6]/div[2]/div[5]/table[2]"
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, table_xpath)))

table = driver.find_element(By.XPATH, table_xpath)
rows = table.find_elements(By.TAG_NAME, "tr")

data = []

for row in rows:
    cols = row.find_elements(By.TAG_NAME, "td")
    cols_text = [col.text for col in cols]
    data.append(cols_text)

df = pd.DataFrame(data)
df.to_excel("tabela_extraida.xlsx", index=False)

driver.quit()

print("Automação concluída com sucesso! A tabela foi salva como 'tabela_extraida.xlsx'.")
