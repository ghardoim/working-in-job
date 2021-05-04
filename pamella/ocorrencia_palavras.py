nome_arquivo = input("Digite o nome do arquivo: ")
palavras = ()
qnt_palavras = ()

while nome_arquivo != "":
  with open(nome_arquivo, "r") as arquivo:
    print(set(arquivo))
    arquivo.seek(0)
    for i, linha in enumerate(arquivo):
      for palavra in tuple(linha.split(" ")):
        palavra = palavra.rstrip("\n")
        palavras += (palavra,)

  nome_arquivo = input("Digite o nome do arquivo: ")

ocorrencias = 0
m_palavra = ""

for palavra in palavras:
  qnt = palavras.count(palavra)

  if qnt >= ocorrencias:
    ocorrencias = qnt
    m_palavra = palavra

print(f"A palavra '{m_palavra}' apareceu {ocorrencias} vezes") 