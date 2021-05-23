from random import choice
from random import sample
import time

with open("palavras.txt", "r") as arquivo:
  palavra = str(choice([ linha.rstrip() for linha in arquivo.readlines() ]))
  print(f"Você tem 6 tentativas para adivinhar a palavra: {''.join(sample(palavra, len(palavra)))}")

  acertou = False
  for tentativa in range(6):

    if str(input(f"Seu {tentativa + 1}° chute: ")) == palavra:
      acertou = True
      break

  if acertou:
    print(f"Parabéns! Você acertou a palavra: {palavra}")
  else:
    print(f"Você errou a palavra: {palavra}")
time.sleep(5)