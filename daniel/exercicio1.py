pesos = (3, 5, 2)
alunos = {}

for i in range(10):
  nome = input(f"Digite o nome do {i + 1}° aluno: ")

  notas = []
  for j in range(3):
    notas.append(float(input(f"Informe a {j + 1}° nota: ")))

  total = sum([ notas[n] * pesos[n] for n in range(3) ])
  alunos[nome] = total / sum(pesos)

for i in range(5):
  nome = input("Informe qual aluno deseja visualizar a média: ")

  if nome not in alunos.keys():
    print(f"O aluno {nome} não foi encontrado!")
  else:
    print(f"A média do aluno {nome} é {alunos[nome]}")