lista = [4, 9, 16, 25, 36, 49, 64, 81, 100, 121, 144, 169, 196]
nova_lista = [ int(n ** 0.5) for n in lista ]

soma = 0
for n in nova_lista:
  soma += n

print(f"A média da raiz quadrada da lista é {soma / len(nova_lista)}")