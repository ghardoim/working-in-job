from json import dumps
import requests as rq
import pandas as pd

neriage_pd = pd.DataFrame()
for i in range(1000):
    body = dumps({
        "call": "ListarProdutos",
        "app_key": "",
        "app_secret": "",
        "param":[ {
            "pagina": i,
            "registros_por_pagina": 100,
            "apenas_importado_api": "N",
            "filtrar_apenas_omiepdv": "N"
        }]
    })
    header = {
        "Content-type": "application/json"
    }
    resp = rq.post("https://app.omie.com.br/api/v1/geral/produtos/", headers = header, data = body)
    if 200 != resp.status_code: break
    neriage_pd = neriage_pd.append(pd.DataFrame(resp.json()["produto_servico_cadastro"]), ignore_index = True)

neriage_pd.to_excel("neriage.xlsx")