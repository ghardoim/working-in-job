import time

def valor(polinomio, valor):
  vPolinomio = 0
  graus = [ grau for grau in range(len(polinomio) - 1, 0, -1) ]

  for grau in graus:
    coeficiente = int(polinomio[graus.index(grau)])

    if 0 != coeficiente:
      vPolinomio += coeficiente * (valor ** grau)

  vPolinomio += int(polinomio[len(polinomio) - 1])
  return vPolinomio

def derivada(polinomio):
  derivada = []   
  graus = [ grau for grau in range(len(polinomio) - 1, 0, -1) ]

  for grau in graus:
    coeficiente = int(polinomio[graus.index(grau)])
    derivada.append(coeficiente * grau)

  return derivada

def somapoli(lista):
  soma = dict.fromkeys([ grau for grau in range(len(max(lista)) - 1, 0, -1) ], 0)

  for polinomio in lista:
    graus = [ grau for grau in range(len(polinomio) - 1, 0, -1) ]

    for grau in graus:
      soma[grau] += int(polinomio[graus.index(grau)])

  novo_polinomio = []
  for key, value in soma.items():
    novo_polinomio.append(value)

  return novo_polinomio

print(valor(list(input("Informe os coeficientes separados por vírgula: ").split(",")), int(input("Informe o valor de X: "))))
print(derivada(list(input("Informe os coeficientes separados por vírgula: ").split(","))))

coeficientes = []
for i in range(int(input("Informe quantos polinomios deseja somar: "))):
  coeficientes.append(list(input(f"Informe os coeficientes do {i + 1}° polinomio separados por vírgula: ").split(",")))

print(somapoli(coeficientes))

time.sleep(5)