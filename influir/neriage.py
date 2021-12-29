from requests import api
from credenciais import *
from json import dumps
import requests as rq
import pandas as pd

def neriage_api(api_call, endpoint, resp_field, api_df = pd.DataFrame()):
    i = 0
    while True:
        body = dumps({
            "call": api_call,
            "app_key": neriage_app_key,
            "app_secret": neriage_app_secret,
            "param":[ {
                "pagina": i,
                "registros_por_pagina": 100,
                "apenas_importado_api": "N",
                **({"filtrar_apenas_omiepdv": "N"} if "Produtos" in api_call else {})
            }]
        })
        header = {
            "Content-type": "application/json"
        }

        resp = rq.post(f"https://app.omie.com.br/api/v1/{endpoint}", headers = header, data = body)
        if 200 != resp.status_code: break

        api_df = api_df.append(pd.DataFrame(resp.json()[resp_field]), ignore_index = True)
        i += 1

    for column in api_df.columns:
        if type(api_df[column][0]) is dict:
            api_df = api_df.join(api_df[column].apply(pd.Series))
            del api_df[column]

    api_df.to_excel(f"{api_call}.xlsx", index = False)

neriage_api('ListarProdutos', 'geral/produtos/', 'produto_servico_cadastro')
neriage_api('ListarPedidos', 'produtos/pedido/', 'pedido_venda_produto')