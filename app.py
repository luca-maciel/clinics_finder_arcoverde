from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import pyautogui as pg
import openpyxl as opxl

nav = webdriver.Chrome(keep_alive=True)

nav.get("https://www.google.com.br/maps/search/clinica/@-8.4226488,-37.0765385,14z?entry=ttu&g_ep=EgoyMDI1MDQyMy4wIKXMDSoASAFQAw%3D%3D")

nav.maximize_window()

pg.moveTo(x=200, y=250)

# print(nav.find_element(By.XPATH, '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[1]/div[1]/div[2]').location)

# nav.find_element(By.XPATH, '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[1]/div[1]/div[2]').click()
sleep(2)

for i in range(20):
    pg.scroll(-5000)
    sleep(2)

sleep(5)

clinicas = nav.find_elements(By.XPATH, "//a[@class='hfpxzc']")

# print(clinicas)

workbook = opxl.Workbook()
workbook.create_sheet(title="clinicas")
# workbook.save(filename="clinicas")

dados = []

for clinica in clinicas:
    nome = clinica.accessible_name 
    link = clinica.get_property("href")
    dados.append([nome, link])
    print(f"Nome: {nome} | link: {link}\n")

for dado in dados:
    workbook["clinicas"].append(dado)

workbook.save("clinicasTeste.xlsx")
