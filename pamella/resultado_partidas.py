def get_time_gols(linha):
  linha = tuple(linha.split("#"))
  return linha[0].strip(), int(linha[1])

nome_arquivo = input("Digite o nome do arquivo de jogos ocorridos: ")
resultados = {}

with open(nome_arquivo, "r") as arquivo:
  for i, linha in enumerate(arquivo):        
    print(linha.rstrip("\n"))
    placar_c, placar_v = tuple(linha.split("X"))

    time_v, gols_v = get_time_gols(placar_v)
    time_c, gols_c = get_time_gols(placar_c)

    resultados[time_c] = { "Jgs": 0, "Vits": 0, "Emps": 0, "Ders": 0, "Pts": 0, "GsP": 0, "GsC": 0, "SGols": 0}
    resultados[time_v] = { "Jgs": 0, "Vits": 0, "Emps": 0, "Ders": 0, "Pts": 0, "GsP": 0, "GsC": 0, "SGols": 0}

  print("\n---------\n")
  arquivo.seek(0)

  for i, linha in enumerate(arquivo):
    placar_c, placar_v = tuple(linha.split("X"))
    
    time_v, gols_v = get_time_gols(placar_v)
    time_c, gols_c = get_time_gols(placar_c)

    resultados[time_c]["Jgs"] += 1
    resultados[time_v]["Jgs"] += 1
    resultados[time_c]["GsP"] += gols_c
    resultados[time_v]["GsP"] += gols_v
    resultados[time_c]["GsC"] += gols_c
    resultados[time_v]["GsC"] += gols_v
    resultados[time_c]["SGols"] += abs(gols_c - gols_v)
    resultados[time_v]["SGols"] += abs(gols_c - gols_v)

    if gols_c > gols_v:
      resultados[time_c]["Vits"] += 1
      resultados[time_c]["Pts"] += 3
      resultados[time_v]["Ders"] += 1

    elif gols_v > gols_c:
      resultados[time_v]["Vits"] += 1
      resultados[time_v]["Pts"] += 3
      resultados[time_c]["Ders"] += 1

    else:
      resultados[time_v]["Emps"] += 1
      resultados[time_v]["Pts"] += 1
      resultados[time_c]["Emps"] += 1
      resultados[time_c]["Pts"] += 1

for resultado in resultados:
  print(f"{resultado}: {resultados[resultado]}")