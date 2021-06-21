from urllib import request as rq
import pandas as pd

covid_filename = "owid-covid-data.csv"

rs = rq.urlopen("https://covid.ourworldindata.org/data/owid-covid-data.csv")
open(covid_filename, "w").write("")

with open(covid_filename, "a") as covidfile:
  for line in rs:
    covidfile.write(line.decode())

covid_df = pd.read_csv(covid_filename)
covid_df = covid_df.fillna(0)

def ehValido(numero):
  return numero >= 0 and numero < len(paises)

def getIntervalo(df, inicio, fim):
  df = df.loc[ df["date"] >= inicio, : ] if inicio else df
  df = df.loc[ df["date"] <= fim, : ] if inicio else df

  return df

def evolucaoDeMortes(pais, inicio = "", fim = ""):
  pais_df = covid_df.loc[covid_df["location"] == pais, ["date", "total_deaths"]]
  pais_df = getIntervalo(pais_df, inicio, fim)

  pais_df.plot(figsize = (15, 5), title = f"{pais} - Evolução de Mortes", x = "date", xlabel = "Data", ylabel = "N° de Mortes")
  pais_df.to_csv(f"{pais} - Evolução de mortes.csv", index = False)

def ranking(continente, pais, inicio = "", fim = ""):
  continente_df = covid_df.loc[covid_df["continent"] == continente, ["total_deaths", "location", "date"]].groupby("location", as_index = False).sum().sort_values("total_deaths", ascending = False).head(10).append(covid_df.loc[covid_df["location"] == pais, ["total_deaths", "location"]].groupby("location", as_index = False).sum())

  continente_df = getIntervalo(continente_df, inicio, fim)
  titulo = f"{pais} - Em relação ao Top10 da {continente}"

  continente_df.plot(figsize = (15, 5), kind = "bar", x = "location", xlabel = "Paises", y = "total_deaths", ylabel = "N° de Mortes", title = titulo)

  continente_df.to_csv(f"{pais} - Top10 da {continente}.csv", index = False)

def correlacao2021(pais, inicio = "", fim = ""):
  pais_df = covid_df.loc[covid_df["location"] == pais, ["total_vaccinations", "total_tests", "date"]]
  pais_df = getIntervalo(pais_df, inicio, fim)

  pais_df.plot(figsize = (15, 5), x = "date", xlabel = "Data", ylabel = "Total", title = f"{pais} - Correlação entre testes e vacinas")
  pais_df.to_csv(f"{pais} - Correlação entre testes e vacinas.csv", index = False)    

while True:

  continentes = list(set(covid_df["continent"]))[1:]
  for i, continente in enumerate(continentes):
    print(i, continente, sep = " - ")
  c = int(input("Escolha, pelo número, um continente: "))

  paises = list(set(covid_df.loc[covid_df["continent"] == continentes[c], "location"]))
  for i, pais in enumerate(paises):
    print(i, pais, sep = " - ")
  p = int(input(f"Escolha, pelo número, um país da {continentes[c]}: "))

  if ehValido(p):
    print("Qual intervalo de tempo deseja ver?", "Deixe em branco para ver todo o período!", sep = "\n")

    inicio = str(input("Data de Início (AAAA-MM-DD): "))
    fim = str(input("Data de Fim (AAAA-MM-DD): "))

    evolucaoDeMortes(paises[p], inicio, fim)
    ranking(continentes[c], paises[p], inicio, fim)
    correlacao2021(paises[p], inicio, fim)

  if "S" == str(input("Deseja sair? (S/N)")).upper() :
    break