import requests as rq
import json

class Consulta:
    def __init__(self, cnpj) -> None:
        self.__cnpj = cnpj.replace(".", "").replace("-", "").replace("/", "").strip()
        self.__endpoint = f"https://receitaws.com.br/v1/cnpj/{self.__cnpj}"

    def run(self):
        return rq.get(self.__endpoint).json()

print(json.dumps(Consulta("07.526.557/0001-00").run(), indent=4))