import time
import urllib
import pandas as pd
from tkinter import Tk as tk
from tkinter import filedialog as fd
from tkinter import messagebox as msgbox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

navegador = webdriver.Chrome()
whats_url = "https://web.whatsapp.com"

janela = tk()
janela.withdraw()
imagem = fd.askopenfilename(title = "Selecione a Imagem!").replace("/", "\\")
arquivo_clientes = fd.askopenfilename(title = "Selecione o Excel com as informações dos clientes!")

clientes_df = pd.read_excel(arquivo_clientes)

navegador.get(whats_url)
while len(navegador.find_elements_by_id("side")) < 1:
  time.sleep(1)

for index, linha in clientes_df.iterrows():
  texto = urllib.parse.quote(f"Olá {linha['Nome']}!\n{linha['Mensagem']}")   
  navegador.get(f"{whats_url}/send?phone={linha['Número']}&text={texto}")

  while len(navegador.find_elements_by_id("side")) < 1:
    time.sleep(1)

  navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/div/span').click()
  navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/span/div[1]/div/ul/li[1]/button/input').send_keys(imagem)
  time.sleep(2)

  navegador.find_element_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/span/div/div/span').click()
  time.sleep(30)

navegador.close()
msgbox.showinfo("Programa Executado.", "As promoções foram compartilhadas!")
janela.destroy()