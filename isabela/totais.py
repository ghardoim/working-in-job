import pandas as pd

cases_df = pd.read_csv("Brazil.txt", sep = "\t")

cases_df.plot(figsize = (15, 5), title = "Total de Casos", x = "date", xlabel = "Data", y = "total_cases", ylabel = "Total Acumulado", color = "orange")
cases_df.plot(figsize = (15, 5), title = "Total de Mortes", x = "date", xlabel = "Data", y = "total_deaths", ylabel = "Total Acumulado", color = "red")