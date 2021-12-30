from pandas import ExcelWriter
from credenciais import *
from json import dumps
import requests as rq
import pandas as pd

def handle_dict_of_dict(df, field_name, fields_list):    
    for field in fields_list:
        df[field] = df["det"].apply(lambda det: det[field_name][field])
    return df

def expand_dicts(df):
    for column in df.columns:
        if type(df[column][0]) is dict and "det" not in column:
            df = df.join(df[column].apply(pd.Series))
            del df[column]
    return df

def keep_columns(df, columns_to_keep):
    for column in df.columns:
        if column not in columns_to_keep: del df[column]
    return df

def neriage_api(api_call, endpoint, resp_field, df = pd.DataFrame(), i = 0):
    while True:
        body = dumps({
            "call": api_call,
            "app_key": neriage_app_key,
            "app_secret": neriage_app_secret,
            "param":[{
                "pagina": i,
                "registros_por_pagina": 100,
                "apenas_importado_api": "N",
                **({"filtrar_apenas_omiepdv": "N"} if "Produtos" in api_call else {})
            }]
        })
        header = { "Content-type": "application/json" }
        resp = rq.post(f"https://app.omie.com.br/api/v1/{endpoint}", headers = header, data = body)
        # if 200 != resp.status_code: break

        df = df.append(pd.DataFrame(resp.json()[resp_field]), ignore_index = True)
        i += 1
        break
    return df

file_name = "_bases_/NERIAGE_VENDA_E_ESTOQUE"
with ExcelWriter(f"{file_name}.xlsx", "xlsxwriter", date_format = "DD/MM/YYYY") as writer:
    writer.book.filename = f'{file_name}.xlsm'
    writer.book.add_vba_project('vbaProject.bin')

    df = neriage_api('ListarProdutos', 'geral/produtos/', 'produto_servico_cadastro')
    keep_columns(df, ["codigo", "codigo_familia", "codigo_produto", "descr_detalhada", "descricao",
                        "descricao_familia", "modelo", "quantidade_estoque", "valor_unitario"]
    ).to_excel(writer, "BASE_PRODUTOS", index = False)

    df = neriage_api('ListarPedidos', 'produtos/pedido/', 'pedido_venda_produto')
    df = expand_dicts(keep_columns(df, ["cabecalho", "det", "informacoes_adicionais"]).explode("det"))
    df = handle_dict_of_dict(df, "ide", ["codigo_item", "codigo_item_integracao"])
    df = handle_dict_of_dict(df, "inf_adic", ["item_pedido_compra", "numero_pedido_compra"])
    df = handle_dict_of_dict(df, "produto", ["codigo", "codigo_produto", "descricao", "percentual_desconto", "quantidade",
                                    "tipo_desconto", "valor_desconto", "valor_mercadoria", "valor_total", "valor_unitario"])
    del df["det"]
    df.to_excel(writer, "BASE_VENDAS", index = False)