pesos = (3, 5, 2)
alunos = {}

for i in range(3):
  nome = input(f"Digite o nome do {i + 1}° aluno: ")

  notas = []
  for j in range(3):
    notas.append(float(input(f"Informe a {j + 1}° nota: ")))

    total = 0
    soma_pesos = 0
    for n in range(3):
      total += notas[n] * pesos[n]
      soma_pesos += pesos[n]

    alunos[nome] = total / soma_pesos

for i in range(5):
  nome = input("Informe qual aluno deseja visualizar a média: ")

  if nome not in alunos.keys():
    print(f"O aluno {nome} não foi encontrado!")
  else:
    print(f"A média do aluno {nome} é {alunos[nome]}")