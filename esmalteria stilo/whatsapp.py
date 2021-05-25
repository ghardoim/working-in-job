import time
import urllib
import pandas as pd
from random import randint
from tkinter import Tk
from tkinter import filedialog as fd
from tkinter import messagebox as msgbox
from tkinter import PhotoImage as newIMG
from tkinter import StringVar as varSTR
from tkinter import Button as newBTN
from tkinter import Label as newLBL
from tkinter import Text as newINP
from tkinter.font import Font as newFNT
from datetime import datetime as dt
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

whats_url = "https://web.whatsapp.com"
footer = '//*[@id="main"]/footer/div[1]/div[1]/div[2]/div/'

def sendMessage(browser, withImage):
  if withImage:
    browser.find_element_by_xpath('//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/span/div/div/span').click()
  else: browser.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button/span').click()

def lastSend(number):
  open("lastNumber.txt", "w").write(str(number))

def wait_for_it(browser):
  while len(browser.find_elements_by_id("side")) < 1: time.sleep(1)

def getPathImg():
  imagem.set(fd.askopenfilename(title = "Selecione a Imagem!").replace("/", "\\"))
    
def getPathExcel():
  excel.set(fd.askopenfilename(title = "Selecione o Excel com as informações dos clientes!"))

def enviar():
  if not excel.get():
    msgbox.showinfo("Esqueceu alguma coisa!", "Selecione o arquivo com os contados dos clientes!")
    return

  clientesDF = pd.read_excel(excel.get())
  mensagem.set(inputTXT.get(1.0, "end"))

  browser = webdriver.Chrome()
  browser.get(whats_url)

  wait_for_it(browser)

  lastNumber = int(open("lastNumber.txt", "r").read())
  for i, linha in clientesDF.iterrows():
    if i < lastNumber: continue

    texto = urllib.parse.quote(f"Olá {linha['Nome']}!\n{mensagem.get()}")
    browser.get(f"{whats_url}/send?phone={linha['Número']}&text={texto}")

    wait_for_it(browser)
    try:
      if imagem.get():
        browser.find_element_by_xpath(f'{footer}div/span').click()
        browser.find_element_by_xpath(f'{footer}span/div[1]/div/ul/li[1]/button/input').send_keys(imagem.get())
        time.sleep(1.2)

      sendMessage(browser, imagem.get())
      time.sleep(randint(9, 15))

    except NoSuchElementException as error:
      log = f"{'-' * 50}\n{dt.now().strftime('%d/%m/%Y - %H:%M')}[{i}]\tNome: {linha['Nome']}\tNúmero: {linha['Número']}\n{'-' * 50}\n"
      open("errorfilelog.txt", "a").write(log)

    lastSend(i + 1)
    if len(clientesDF) - 1 == i: lastSend(0)

  browser.close()

janela = Tk()

janela.title("HardsBot")
janela.config(bg = "lightblue")
janela.iconphoto(False, newIMG(file = "hardsbot.png"))
janela.rowconfigure([0, 1, 2, 3], weight = 1)
janela.columnconfigure([0, 1], weight = 1)
janela.geometry("550x350")
janela.minsize(550, 350)
janela.maxsize(550, 350)

newLBL(text = "Esmalteria Stilo", bg = "lightblue", font = newFNT(size = 20, slant = "italic")).grid(row = 0, column = 0, columnspan = 2)

inputTXT = newINP(width = 40, height = 15)
inputTXT.insert(1.0, "Desperte sua beleza!\n\nConheça as promoções da nossa semana! 😍")
inputTXT.grid(row = 1, column = 0, rowspan = 3)

excel = varSTR()
imagem = varSTR()
mensagem = varSTR()

btnFont = newFNT(size = 12)
newBTN(text = "Buscar Imagem", command = getPathImg, bg = "pink", width = 15, height = 2, font = btnFont).grid(row = 1, column = 1)
newBTN(text = "Buscar Clientes", command = getPathExcel, bg = "pink", width = 15, height = 2, font = btnFont).grid(row = 2, column = 1)
newBTN(text = "Enviar", command = enviar, bg = "pink", width = 15, height = 2, font = btnFont).grid(row = 3, column = 1)

janela.mainloop()