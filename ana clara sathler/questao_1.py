from datetime import datetime as dt

hora_consulta = str(input("Informe o horário da sua consulta: (HH:MM)"))
hora_chegada = str(input("Informe o horário da chegada: (HH:MM)"))

hora_consulta = dt.strptime(hora_consulta, '%H:%M')
hora_chegada = dt.strptime(hora_chegada, '%H:%M')

minutos = abs((hora_chegada - hora_consulta).total_seconds() / 60.0)

if hora_chegada < hora_consulta:
  print("Você chegou a tempo da consulta!")
  print(f"Está {minutos} minutos adiantado!")

elif hora_chegada > hora_consulta:
  print("Você chegou atrasado para a consulta!")
  print(f"Está {minutos} minutos atrasado!")

else:
  print("Você chegou na hora exata para a consulta!")