import requests as req
from bs4 import BeautifulSoup
from datetime import datetime as dt
from twilio.rest import Client
import pandas as pd

account_sid = ""
auth_token = ""
client = Client(account_sid, auth_token)

lbr_mainURL = "https://www.leiloesbr.com.br/"
search_terms = []

def extract_infos(search_terms):
  infos = []
  for term in search_terms:

    nPage = 1
    while True:
      html = BeautifulSoup(req.get(f"{lbr_mainURL}busca.asp?v=126&op=2&pesquisa={term}&pag={nPage}").text, 'html.parser')

      for i, item in enumerate(html.find_all("div", class_ = "col-xs-12 col-sm-4 col-md-4")):
        lbr_info = {}

        tagA = item.find("a", attrs = { "data-ref": i})
        img_url = tagA["data-zoom-image"]
        lbr_info["url_image"] = img_url

        lbr_info["id"] = img_url.split('/')[6].split('.')[0]
        lbr_info["post_url"] = lbr_mainURL + tagA["href"]

        lbr_info["title"] = item.find("div", class_ = "item-title").a["data-original-title"]

        lbr_info["search_term"] = term
        lbr_info["channel"] = "leilõesBR"
        lbr_info["scrape_date"] = dt.now().strftime(r"%Y-%m-%d")
        lbr_info["scrape_timestamp"] = dt.now().strftime(r"%H:%M:%S")
        lbr_info["full_info"] = item

        # item.find("span", class_ = "pesq-uf").parent.text.split(" - ")[0] # prazo? vencimento?
        lbr_info["upload_date"] = ""
        lbr_info["upload_timestamp"] = ""

        lbr_info["price"] = int(item.find("div", class_ = "item-price").h4.text.replace("R$ ", "").replace(",00", "").replace(".", ""))
        infos.append(lbr_info)

      nPage += 1

      rButton = html.find(class_ = "fa fa-chevron-right")
      hasNext = False if "a" != rButton.parent.name or rButton.has_attr("aria-hidden") else True
      if not hasNext:
        break

  return infos

scraped_data = extract_infos(search_terms)

old_data = pd.read_excel("data/scraped_data_leiloesBR.xlsx")
old_ids = set([ int(_id) for _id in old_data["id"]])

ids = [ int(item["id"]) for item in scraped_data ]
new_ids = [ _id for _id in ids if _id not in old_ids ]

if new_ids:
  scraped_data = [ data for data in scraped_data if int(data["id"]) in new_ids ]

  new_data = old_data.append(scraped_data)
  new_data.to_excel("data/scraped_data_leiloesBR.xlsx", index = False)

  for new_item in scraped_data:
    msg_text = f"Tem item novo no LeilõesBR!\n{new_item['title']}\n{new_item['post_url']}"
    message = client.messages.create(
      body = msg_text,
      from_ = 'whatsapp:+',
      to = 'whatsapp:+',
      media_url = f"{new_item['url_image']}")