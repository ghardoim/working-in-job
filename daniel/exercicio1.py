pesos = (3, 5, 2)
alunos = {}

for i in range(10):
  nome = input(f"Digite o nome do {i + 1}° aluno: ")

  soma_pesos = 0
  notas = []
  total = 0

  for j in range(3):
    notas.append(float(input(f"Informe a {j + 1}° nota: ")))

    total += notas[j] * pesos[j]
    soma_pesos += pesos[j]

  alunos[nome] = total / soma_pesos

for n in range(5):
  nome = input("Informe qual aluno deseja visualizar a média: ")

  if nome not in alunos.keys():
    print(f"O aluno {nome} não foi encontrado!")
  else:
    print(f"A média do aluno {nome} é {alunos[nome]}")