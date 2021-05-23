import time

conta = { 
  "saldo": 0,
  "transacoes": 0,
  "media": 0
}
def compra(conta, valor):

  conta["saldo"] -= valor
  conta["transacoes"] += 1
  conta["media"] = abs(conta["saldo"] / conta["transacoes"])

  return conta

for n in range(int(input("Quantas compras deseja fazer? "))):
  conta = compra(conta, float(input(f"Qual o valor da {n + 1}° compra?")))

print(f'O seu saldo agora é de R${conta["saldo"]} pois foram realizadas {conta["transacoes"]} transações numa média de R${conta["media"]}')
time.sleep(5)