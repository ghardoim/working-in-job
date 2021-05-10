lancamentos = (1, 1, 2, 6, 6, 6, 4, 1, 3, 3, 1, 2, 2, 6, 5, 5, 1, 3, 5, 4, 2, 1, 3, 2, 1, 1, 2, 3, 3, 3, 4, 4, 5, 6, 2, 4, 2, 3, 1, 2, 4, 5, 2, 6, 4, 1, 3, 2, 2, 4)

dado = {}
for i in range(1, 7):

  qnt_lancamentos = 0
  for lancamento in lancamentos:

    if lancamento == i:
      qnt_lancamentos += 1

  dado[str(i)] = f"{int(qnt_lancamentos / len(lancamentos) * 100)}%"

print(dado)