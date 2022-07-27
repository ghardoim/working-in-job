from playwright.sync_api import sync_playwright
from tkinter.filedialog import askopenfilename
from main import _entry, _label, _window
from tkinter import Button
import pandas as pd

def loto_giro(email:str, senha:str, value:str) -> None:
    if email and senha and value:

        url_base = "https://lotogiro.com.br/lotogiro/backoffice/"
        url_game = f"{url_base}games/lotofacil#"
        url_login = f"{url_base}login"

        browser = sync_playwright().start().chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(url_login)
        page.wait_for_load_state("load")

        page.locator('//*[@id="login"]/div[2]/div/div/div[2]/input').fill(email)
        page.locator(".password").fill(senha)

        while page.url == url_login:
            page.wait_for_timeout(1000)

        page.goto(url_game)
        page.wait_for_load_state("load")
        page.locator("#valor1").fill(value)

        df = pd.read_excel(askopenfilename(title="ESCOLHA O ARQUIVO COM AS DEZENAS")).dropna(how="all").iloc[3:].reset_index().fillna(0)
        for _, game in df.iterrows():
            for ten in range(1, len(df.columns) - 1):
                page.locator(".just", has_text=str(int(game[f'Unnamed: {ten}']))).nth(0).click()

            page.locator("span", has_text="Validar Aposta").click()

            page.goto(url_game)
            page.wait_for_load_state("load")

def _button(text:str, script, row:int=0, column:int=0, rowspan:int=1, colspan:int=1) -> None:
    Button(text=text, command=lambda: script(email.get(), senha.get(), value.get()), font=("Arial", 15), bg="lightblue", width=20) \
        .grid(row=row, column=column, rowspan=rowspan, columnspan=colspan)

if "__main__" == __name__:
    window = _window()

    _label()
    _label("email:", 1, 1)
    _label("senha:", 2, 1)
    _label("valor por jogo:", 3, 1)

    _label(row=4)
    email = _entry(1, 2, width=20)
    senha = _entry(2, 2, show="*", width=20)
    value = _entry(3, 2)

    _label(row=4, column=3)
    _button("realizar jogo", loto_giro, 5, column=1, colspan=2)

    _label(row=6)
    window.mainloop()