from math import sqrt

lista = [4, 9, 16, 25, 36, 49, 64, 81, 100, 121, 144, 169, 196]
nova_lista = [ int(sqrt(n)) for n in lista ]

print(f"A média da raiz quadrada da lista é {sum(nova_lista) / len(nova_lista)}")