from playwright._impl._api_types import TimeoutError
from playwright.sync_api._generated import Page
from playwright.sync_api import sync_playwright
from tkinter.filedialog import askopenfilename
import pandas as pd

def loading(page:Page) -> None:
    for state in ["visible", "hidden"]:
        try: page.wait_for_selector("#popupLoading", state=state, timeout=2500)
        except TimeoutError: pass

def wait_for(page:Page, selector:str) -> None:
    try: page.wait_for_selector(selector, state="visible", timeout=2500).click()
    except TimeoutError: pass

def sorte_(what:str, guess:str="") -> None:
    if not guess: return
    file_name = askopenfilename(title="ESCOLHA A PLANILHA COM OS JOGOS").lower()
    
    _types = { "l": "lotinha", "q": "quininha", "s": "seninha", "qb": "quina-brasil"}
    game_type = [ _types[key] for key in _types if _types[key].replace("-", " ") in file_name]
    try: game_type = game_type[0]
    except: return

    df = pd.read_excel(file_name).dropna(how="all").iloc[3:].reset_index().fillna(0)
    browser = sync_playwright().start().chromium.launch(headless=False)

    page, last_game = browser.new_page(), 0
    page.goto(f"https://www.{'sorteplay.com' if 'play' == what else 'sortenet.bet'}/{game_type}")

    if "net" == what: loading(page)
    wait_for(page, ".btn-fechar-notificacao")

    for row, game in df.iterrows():
        if row < last_game: continue
        for ten in range(1, len(df.columns) - 1):
            n = str(int(game[f'Unnamed: {ten}'])).zfill(2)
            if int(row) and int(n): page.locator("#numeros").locator("[digito='%s']" % n).click()

        page.wait_for_timeout(500)
        page.locator("#valor").fill(guess)
        page.locator("#addPalpite").click()
        wait_for(page, f"text=\"{'Fechar' if 'play' == what else 'Voltar'}\"")

        last_game += 1
        if last_game % (20 if "play" == what else 12) == 0: break

        page.wait_for_timeout(500)
        page.locator(".btn-apagar").click()
    page.locator("#btnPagar" if "play" == what else ".bx-cicle.icon-carrinho").click()