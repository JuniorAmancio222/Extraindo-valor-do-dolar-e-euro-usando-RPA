# Importa o WebDriver do Selenium, renomeado como "opcoes_selenium_aula"
from selenium import webdriver as opcoes_selenium_aula

# Importa a classe Keys para simular pressionamentos de teclas
from selenium.webdriver.common.keys import Keys

# Importa a biblioteca pyautogui, renomeada como "tempoPausaComputador", para automação de ações no computador
import pyautogui as tempoPausaComputador


import pyautogui as teclasAtalho

# Importa a classe By para localizar elementos na página
from selenium.webdriver.common.by import By

#Passamos autorização ao acesso as configurações do Chrome
meuNavegador = opcoes_selenium_aula.Chrome()
meuNavegador.get("https://www.google.com.br/")

tempoPausaComputador.sleep(4)

#Procurando pelo elemento NAME e quando encontrar vou escrever Dolar hoje
meuNavegador.find_element(By.NAME, "q").send_keys("Dolar cotação")

tempoPausaComputador.sleep(4)

#Faz a busca do valor que está digitado no campo NAME q
meuNavegador.find_element(By.NAME, "q").send_keys(Keys.RETURN)

tempoPausaComputador.sleep(4)

#No resultado da pesquias pegamo o XPATH e no meio pegamos o primeiro elemento da lista
valorDolarPeloGoogle = meuNavegador.find_element(By.XPATH, '//*[@id="rso"]/div[1]/div/div[5]/block-component/div/div[1]/div/div/div/div/div[1]/div/div/div/div/div[1]/div/div[1]/div').text
                                        
print(valorDolarPeloGoogle)

#---------------------------------------------

tempoPausaComputador.sleep(2)

meuNavegador.find_element(By.NAME, "q").send_keys("")

#Estamos usando o pyautgui para apertar a tecla TAB
teclasAtalho.press('tab')

tempoPausaComputador.sleep(4)

teclasAtalho.press('enter')

meuNavegador.find_element(By.NAME, "q").send_keys("euro cotação hojee")

tempoPausaComputador.sleep(4)

meuNavegador.find_element(By.NAME, "q").send_keys(Keys.RETURN)

valorEuroPeloGoogle = meuNavegador.find_element(By.XPATH, '//*[@id="rso"]/div[1]/div/block-component/div/div[1]/div/div/div/div/div[1]/div/div/div[2]/div/div[1]/div/div[1]/div').text

print(valorEuroPeloGoogle)

#--------------------------------------------------------

import xlsxwriter
import os

nomeCaminhoArquivo = "Imprime Dolar e Euro Google.xlsx"  # Salva na mesma pasta que o script
planilhaCriada = xlsxwriter.Workbook(nomeCaminhoArquivo)
sheet1 = planilhaCriada.add_worksheet()

sheet1.write("A1", "Dolar")
sheet1.write("B1", "Euro")
sheet1.write("A2", valorDolarPeloGoogle)
sheet1.write("B2", valorEuroPeloGoogle)

#Substituir a vírgula por ponto para o valor converter para float
valorDolarPeloGoogle = valorDolarPeloGoogle.replace(',','.')
valorEuroPeloGoogle = valorEuroPeloGoogle.replace(',','.')

#Convertendo o valor do dolar e euro de String para Float
valor_Dolar_Tipo_Float = float(valorDolarPeloGoogle)
valor_Euro_Tipo_Float = float(valorEuroPeloGoogle)

sheet1.write("A3", valor_Dolar_Tipo_Float)
sheet1.write("B3", valor_Euro_Tipo_Float)

planilhaCriada.close()

os.startfile(nomeCaminhoArquivo)