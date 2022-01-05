from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from tkinter.filedialog import askopenfilename
from selenium.webdriver.common.by import By
from tkinter import StringVar as strVAR
from tkinter import Button as newBTN
from tkinter import IntVar as intVAR
from tkinter import Label as newLBL
from tkinter import Entry as newINP
from selenium import webdriver
from tkinter import Tk
import pandas as pd

def cria_jogos(valor_palpite):
    if not valor_palpite: return
    if not filename.get(): filename.set(askopenfilename(title = "ESCOLHA A PLANILHA COM OS JOGOS").lower())

    tipos = { "l": "lotinha", "q": "quininha", "s": "seninha", "qb": "quina-brasil"}
    for key in tipos:
        if tipos[key].replace("-", " ") in filename.get():
            tipo_jogo = tipos[key]
            break
    if not tipo_jogo: return

    jogosdf = pd.read_excel(filename.get()).dropna(how = "all").iloc[3:].reset_index().fillna(0)
    options = webdriver.ChromeOptions()
    options.add_argument('ignore-certificate-errors')
    browser = webdriver.Chrome(options = options)
    browser.maximize_window()

    wait = WebDriverWait(browser, 20)
    browser.get(f"https://www.sorteplay.com/{tipo_jogo}")
    wait.until(EC.element_to_be_clickable((By.ID, "conteinerNumeros")))

    for row, jogo in jogosdf.iterrows():
        if row < last_game.get(): continue
        print(last_game.get())

        for dezena in range(1, len(jogosdf.columns) - 1):
            n = str(int(jogo[f'Unnamed: {dezena}'])).zfill(2)
            if int(row) and int(n):
                browser.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[@digito='{n}']"))))

        valor = wait.until(EC.element_to_be_clickable((By.ID, "valor")))
        browser.execute_script(f"arguments[0].value = '{valor_palpite}';", valor)
        browser.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.ID, "addPalpite"))))

        last_game.set(last_game.get() + 1)
        if last_game.get() % 24 == 0: break
        for numero in browser.find_elements(By.CLASS_NAME, "num active"):
            browser.execute_script("arguments[0].classList.remove('active');", numero)

    browser.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text() = 'Avançar']"))))
    browser.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text() = 'PAGAR']"))))

janela = Tk()

janela.title("Desk.Robot")
janela.config(bg = "lightgray")
janela.rowconfigure([0, 3], weight = 1)
janela.columnconfigure([0], weight = 1)
janela.minsize(350, 250)
janela.maxsize(350, 250)

filename = strVAR(janela, "")
last_game = intVAR(janela, 0)

newLBL(text = "valor por jogo:", bg = "lightgray").grid(row = 1)
valorTXT = newINP(janela, width = 30)
valorTXT.grid(row = 2)

newBTN(text = "abrir o navegador e criar os jogos", command = lambda: cria_jogos(valorTXT.get()),
        bg = "lightblue", width = 35, height = 1).grid(row = 3)
janela.mainloop()