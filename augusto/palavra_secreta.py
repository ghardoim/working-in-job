#usado para descobrir a posição da letra
alfabeto_latino = "abcdefghijklmnopqrstuvwxyz"
criptografia = str(input("Entre com uma sequencia de 26 caracteres para a criptografia: \n-> ")).lower()

if 26 != len(criptografia) or not criptografia.isalpha():
    print("A criptografia precisa conter apenas 26 letras!")

else:
    palavra = input("Informe a palavra codificada: ")
    print(*[ criptografia[alfabeto_latino.index(letra)] for letra in palavra ], sep = "")