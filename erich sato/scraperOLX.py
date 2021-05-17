import requests
import pandas as pd
from bs4 import BeautifulSoup
from twilio.rest import Client
from datetime import timedelta, datetime as dt

header = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.106 Safari/537.36'}

search_terms = []

account_sid = ""
auth_token = ""
client = Client(account_sid, auth_token)

def request_get(term, n_page):
  return requests.get(f"http://www.olx.com.br/brasil?o={n_page}&q={term}", headers = header)

def has_next_button(response):
  return BeautifulSoup(response.text, 'html.parser').find(name = "a", attrs = { "data-lurker-detail": "forward_button" })

def extracted_date(day):
  if "Ontem" == day:
    day = (dt.now() - timedelta(days = 1))
  elif "Hoje" == day:
    day = dt.now()
  else:
    day = dt.strptime(f"{day}/{dt.now().year}", r"%d/%m/%Y")
  return day.strftime(r"%Y-%m-%d")

def format_price(price):
  return int(price.get_text().split("R$ ")[1].replace(".", ""))

def extract_infos(search_terms):
  infos = []
  for term in search_terms:
    n_page = 1    
    while True:
      print(f'{term} / pagina {n_page}')

      response = request_get(term, n_page)
      soup_data = BeautifulSoup(response.text, 'html.parser')

      for item in soup_data.find_all(name = "a", attrs = { "data-lurker-detail": "list_id" }):
        olx_info = {}

        olx_info["id"] = item["data-lurker_list_id"]
        olx_info["post_url"] = item["href"]
        olx_info["url_image"] = item.find(name = "img")["src"]
        olx_info["title"] = item["title"]
        olx_info["search_term"] = term
        olx_info["channel"] = "olx"
        olx_info["scrape_date"] = dt.now().strftime(r"%Y-%m-%d")
        olx_info["scrape_timestamp"] = dt.now().strftime(r"%H:%M:%S")
        olx_info["full_info"] = item.parent

        full_date = item.div.find(name = "p", class_ = "sc-1iuc9a2-4 hDBjae sc-ifAKCX fWUyFm")
        if full_date is not None:
          full_date = full_date.text.split(" ")
          olx_info["upload_date"] = extracted_date(full_date[0])
          olx_info["upload_timestamp"] = full_date[2]

        price = item.find(name = "p", class_ = "sc-ifAKCX eoKYee")
        olx_info["price"] = format_price(price) if price is not None else 0
        
        infos.append(olx_info)
      n_page += 1
      if not has_next_button(request_get(term, n_page)):
        break
  return infos

scraped_data = extract_infos(search_terms)

old_data = pd.read_excel("data/scraped_data_olx.xlsx")
old_ids = set([ int(_id) for _id in old_data["id"]])

ids = [ int(item["id"]) for item in scraped_data ]
new_ids = [ _id for _id in ids if _id not in old_ids ]

print(new_ids)
if new_ids:
  scraped_data = [ data for data in scraped_data if int(data["id"]) in new_ids ]

  new_data = old_data.append(scraped_data)
  new_data.to_excel("data/scraped_data_olx.xlsx", index = False)
  
  for new_item in scraped_data:
    msg_text = f"Tem item novo na OLX!\n{new_item['title']}\n{new_item['post_url']}"
    
    message = client.messages.create(
      body = msg_text,
      from_ = 'whatsapp:+',
      to = 'whatsapp:+',
      media_url = f"{new_item['url_image']}", 
    )