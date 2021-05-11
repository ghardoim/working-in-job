def converte(bytes_usados):
  return round(bytes_usados / 1048576, 2)

def percentual_de_uso(espaco_usado, total):
  return round((espaco_usado / total) * 100, 2)

funcionarios = {}
with open("funcionarios.txt", "r") as arquivo:    
  for linha in arquivo.readlines():
    if linha != "\n":

      linha = linha.rstrip().split(" ")
      funcionarios[linha[0]] = int(linha[1])

total_espaco_usado = converte(sum(funcionarios.values()))
media_espaco_usado = round(total_espaco_usado / len(funcionarios.keys()), 2)

with open("relatorio.txt", "wb") as arquivo:
  arquivo.write(f"EmpresaX{' ' * 19}Uso do espaço em disco pelos usuários\n\n".encode("utf8"))
  arquivo.write(f"{'-' * 70}\n\n".encode("utf8"))
  arquivo.write(f"Usuário{' ' * 20}Espaço utilizado{' ' * 20}% do uso\n\n".encode("utf8"))

  for nome, espaco_utilizado in funcionarios.items():
    espaco_em_mb = converte(espaco_utilizado)
    percentual = percentual_de_uso(espaco_em_mb, total_espaco_usado)
    
    arquivo.write(f"{nome}{' ' * 22}{espaco_em_mb} MB{' ' * 22}{percentual}%\n\n".encode("utf8"))

  arquivo.write(f"Espaço total ocupado: {total_espaco_usado} MB\n\n".encode("utf8"))
  arquivo.write(f"Espaço médio ocupado: {media_espaco_usado} MB".encode("utf8"))