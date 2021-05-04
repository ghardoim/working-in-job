def eh_bomba(novo_campo, x, y):
  if '-' != novo_campo[x][y] and '*' != novo_campo[x][y]:
    novo_campo[x][y] += 1
  else:
    novo_campo[x][y] = 1

  return novo_campo

linhas = int(input("Informe o número de linhas: "))
colunas = int(input("Informe o número de colunas: "))

print("Informe as posições!\nUse '*' (asterisco) para indicar as minas!\nUse '-' (hífen) para indicar as casas vazias!")
campo_minado = []
for linha in range(1, linhas + 1):
  str_linha = str(input(f"Informe as {colunas} casas da {linha}° linha: "))
  campo_minado.append(str_linha)

novo_campo = [ list(linha) for linha in campo_minado ]
for i_linha, linha in enumerate(campo_minado):
  if "*" in linha:

    bomba = linha.index("*")
    antes_bomba = bomba - 1
    depois_bomba = bomba + 1

    if 0 < antes_bomba:
      novo_campo = eh_bomba(novo_campo, i_linha, antes_bomba)

    if depois_bomba < len(novo_campo[i_linha]):
      novo_campo = eh_bomba(novo_campo, i_linha, depois_bomba)

    if 0 < i_linha:
      novo_campo = eh_bomba(novo_campo, i_linha - 1, bomba)

    if i_linha < len(campo_minado) - 1:
      novo_campo = eh_bomba(novo_campo, i_linha + 1, bomba)

[ print(*linha) for linha in novo_campo ]