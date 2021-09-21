agenda = []

def procurar(qual, info, agenda):
  clientes_encontrados = [ cliente for cliente in agenda if info in cliente[qual] ]
  for cliente in clientes_encontrados:
    print(f"\nnome: {cliente['nome']}\nidade: {cliente['idade']} anos\nnacionalidade: {cliente['nacionalidade']}")
  
  if not clientes_encontrados:
    print("\nCliente não encontrado :/")

def cadastrar():
  nome = str(input("Informe o nome do cliente: ")).capitalize()
  idade = int(input("Informe a idade do cliente: "))
  nacionalidade = str(input("Informe a nacionalidade do cliente: ")).capitalize()

  if len(agenda) < 10:
    agenda.append({ "nome": nome, "idade": idade, "nacionalidade": nacionalidade })
  else:
    print("Agenda cheia")

def consultar(agenda):
  procurar("nome", str(input("Informe o nome do cliente: ")).capitalize(), agenda)

def listar_mesma_nacioanalidade(agenda):
  procurar("nacionalidade", str(input("Informe nacionalidade do cliente: ")).capitalize(), agenda)

def menu():
    while True:
      print("\n" * 2)
      print("-" * 50)
      print("1 - Cadastrar novo cliente")
      print("2 - Consultar cliente")
      print("3 - Listar clientes de uma mesma nacionalidade")
      print("4 - Fim")

      try:
        opcao = int(input("Escolha uma opção: "))
      except ValueError:
        print("\nDigite uma opção válida!")
        continue

      if 1 == opcao:
        cadastrar()
      elif 2 == opcao:
        consultar(agenda)
      elif 3 == opcao:
        listar_mesma_nacioanalidade(agenda)
      elif 4 == opcao:
        break
      else:
        print("Não temos essa opção no menu :/")
menu()