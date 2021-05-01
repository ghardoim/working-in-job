from os.path import exists
from json import loads

txt_estoque = "estoque.txt"
EstoqueLivros = []
if exists(txt_estoque):
  with open(txt_estoque, "r") as arquivo_estoque:
    EstoqueLivros = [ loads(livro.strip("\n").replace("'","\"")) for livro in arquivo_estoque.readlines() ]

txt_saldo = "saldo.txt"
Saldo = 0
if exists(txt_saldo):
  with open(txt_saldo, "r") as arquivo_saldo:
    Saldo = float(arquivo_saldo.readline())

def cadastra_livro(titulo, valor, isbn, quantidade_estoque):
  for livro in EstoqueLivros:
    if isbn == livro["ISBN"]:
      livro["QuantidadeEstoque"] += quantidade_estoque
      return
  novo_livro = {}       
  novo_livro["Titulo"] = titulo
  novo_livro["ISBN"] = isbn
  novo_livro["Valor"] = valor
  novo_livro["QuantidadeEstoque"] = quantidade_estoque
  EstoqueLivros.append(novo_livro)

def consulta_estoque_titulo(titulo):
  for livro in EstoqueLivros:
    if titulo == livro["Titulo"]:
      print(r"O livro foi encontrado em estoque!")
      print(f"Titulo: {livro['Titulo']}\nISBN: {livro['ISBN']}\nValor: {livro['Valor']}\nQuantidade em Estoque: {livro['QuantidadeEstoque']}")
      return
  print(f"O livro {titulo} não foi encontrado em estoque")

def consulta_estoque_ISBN(isbn):
  for livro in EstoqueLivros:
    if isbn == livro["ISBN"]:
      print(r"O livro foi encontrado em estoque!")
      print(f"Titulo: {livro['Titulo']}\nISBN: {livro['ISBN']}\nValor: {livro['Valor']}\nQuantidade em Estoque: {livro['QuantidadeEstoque']}")
      return
  print(f"O livro de ISBN {isbn} não foi encontrado em estoque")

def vender_livro(isbn, quantidade):
  total = 0.0
  for livro in EstoqueLivros:
    if isbn == livro["ISBN"]:
      if livro["QuantidadeEstoque"] - quantidade >= 0:
        livro["QuantidadeEstoque"] -= quantidade
        total += livro["Valor"] * quantidade
        print(f"Foram vendidos {quantidade} livros do IBSN {isbn}")
        
      else:
        print(f"O livro de IBSN {isbn} está em falta no estoque para a quantidade desejada!")
      return total

  print(f"O livro de ISBN {isbn} não foi encontrado para ser vendido!")

def consular_saldo():
  print(f"O saldo atual é de R${Saldo}")

def salvar_dados():
  with open(txt_estoque, "w") as arquivo_estoque:
    for livro in EstoqueLivros:
      arquivo_estoque.write(f"{livro}\n")
  
  with open(txt_saldo, "w") as arquivo_saldo:
      arquivo_saldo.write(f"{Saldo}")

if "__main__" == __name__:
  while True:
    try:
      opcao = int(input("""
        MENU PRINCIPAL
        1. Cadastrar Livro
        2. Consultar Estoque (Buscar Por Título)
        3. Consultar Estoque (Buscar Por ISBN)
        4. Vender Livro
        5. Consultar Saldo da Loja
        6. Salvar Dados
        
        9. Sair
        
      : """))
    except ValueError:
      print("Informe uma opção válida!")
      continue
    
    if 1 == opcao:
      titulo = input("Digite o título do livro: ")
      isbn = int(input("Digite o ISBN do livro: "))
      valor = float(input("Digite o valor do livro: "))
      qnt_estoque = int(input("Digite a quantidade de livros: "))
      cadastra_livro(titulo, valor, isbn, qnt_estoque)

    elif 2 == opcao:
      titulo = input("Digite o título do livro que deseja procurar: ")
      consulta_estoque_titulo(titulo)
      
    elif 3 == opcao:
      isbn = int(input("Digite o ISBN do livro que deseja procurar: "))
      consulta_estoque_ISBN(isbn)

    elif 4 == opcao:
      isbn = int(input("Digite o ISBN do livro que deseja vender: "))
      quantidade = int(input("Digite a quantidade de livros que deseja vender: "))
      Saldo += vender_livro(isbn, quantidade)
    
    elif 5 == opcao:
      consular_saldo()

    elif 6 == opcao:
      salvar_dados()

    elif 9 == opcao:
      print("Encerrando!")
      break