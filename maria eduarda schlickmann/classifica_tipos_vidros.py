tabela = [
    [ 1, 152.101, 13.64, 4.49, 1.10, 71.78, 0.06, 8.75, 0.00, 0.00, 1 ],
    [ 2, 151.761, 13.89, 3.60, 1.36, 72.73, 0.48, 7.83, 0.00, 0.00, 1 ],
    [ 3, 151.618, 13.53, 3.55, 1.54, 72.99, 0.39, 7.78, 0.00, 0.00, 1 ],
    [ 4, 151.766, 13.21, 3.69, 1.29, 72.61, 0.57, 8.22, 0.00, 0.00, 1 ],
    [ 92, 151.605, 12.90, 3.44, 1.45, 73.06, 0.44, 8.27, 0.00, 0.00, 2 ],
    [ 93, 151.588, 13.12, 3.41, 1.58, 73.26, 0.07, 8.39, 0.00, 0.19, 2 ],
    [ 94, 151.590, 13.24, 3.34, 1.47, 73.10, 0.39, 8.22, 0.00, 0.00, 2 ],
    [ 150, 151.643, 12.16, 3.52, 1.35, 72.89, 0.57, 8.53, 0.00, 0.00, 3 ],
    [ 151, 151.665, 13.14, 3.45, 1.76, 72.48, 0.60, 8.38, 0.00, 0.17, 3 ],
    [ 152, 152.127, 14.32, 3.90, 0.83, 71.50, 0.00, 9.49, 0.00, 0.00, 3 ],
    [ 153, 151.779, 13.64, 3.65, 0.65, 73.00, 0.06, 8.93, 0.00, 0.00, 3 ],
    [ 154, 151.610, 13.42, 3.40, 1.22, 72.69, 0.59, 8.32, 0.00, 0.00, 3 ],
    [ 167, 152.151, 11.03, 1.71, 1.56, 73.44, 0.58, 11.62, 0.00, 0.00, 5 ],
    [ 182, 151.888, 14.99, 0.78, 1.74, 72.50, 0.00, 9.95, 0.00, 0.00, 6 ],
    [ 183, 151.916, 14.15, 0.00, 2.09, 72.74, 0.00, 10.88, 0.00, 0.00, 6 ],
    [ 184, 151.969, 14.56, 0.00, 0.56, 73.48, 0.00, 11.22, 0.00, 0.00, 6 ],
    [ 185, 151.115, 17.38, 0.00, 0.34, 75.41, 0.00, 6.65, 0.00, 0.00, 6 ],
    [ 186, 151.131, 13.69, 3.20, 1.81, 72.81, 1.76, 5.43, 1.19, 0.00, 7 ],
    [ 193, 151.623, 14.20, 0.00, 2.79, 73.46, 0.04, 9.04, 0.40, 0.09, 7 ],
    [ 194, 151.719, 14.75, 0.00, 2.00, 73.02, 0.00, 8.53, 1.59, 0.08, 7 ],
    [ 195, 151.683, 14.56, 0.00, 1.98, 73.29, 0.00, 8.52, 1.57, 0.07, 7 ],
    [ 196, 151.545, 14.14, 0.00, 2.68, 73.39, 0.08, 9.07, 0.61, 0.05, 7 ]
]

def maior_valor(indice_buscar, indice_comparar):
    maior, valor_buscado = 0, 0
    for linha in tabela:
        if linha[indice_comparar] > maior:
            maior = linha[indice_comparar]
            valor_buscado = linha[indice_buscar]
    return valor_buscado

print(f"O número de indentificação do maior valor de cálcio é: {maior_valor(0, 7)}")
print(f"O valor do alumínio do vidro com maior índice de refração {maior_valor(4, 1)}")

qnt_e_soma = [ [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0] ]
for linha in tabela:
    for qnt in range(1, len(qnt_e_soma) + 1):
        if qnt == linha[-1]:
            qnt_e_soma[qnt - 1][0] += 1
            qnt_e_soma[qnt - 1][1] += linha[5]

for media in range(1, len(qnt_e_soma) + 1):
    if 0 < qnt_e_soma[media - 1][0]:
        print(f"A média de silício do tipo de vidro {media} é {qnt_e_soma[media - 1][1] / qnt_e_soma[media - 1][0]}")
    else:
        print(f"A média de silício do tipo de vidro {media} é 0")