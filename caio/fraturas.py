import pandas as pd

def handle_sexo(sexo):
    sexo = sexo.upper()
    return "Masculino" if "H" == sexo else "Feminino" if "M" == sexo else sexo.capitalize()

def handle_diagnostico(diagnostico):
    return diagnostico.lower().strip()

def handle_lado(diagnostico):
    return "Direito" if "dir" in diagnostico else "Esquerdo" if "esq" in diagnostico else "Não Identificado"

def handle_tipo(diagnostico):
    return "Fratura" if "fratura" in diagnostico else "Luxação" if "luxacao" in diagnostico else ""

def handle_local(diagnostico):
    local = ""
    if "clavicula" in diagnostico:
        local += " / " if local else ""
        local = "Esternoclavicular" if "esterno" in diagnostico else "Acromioclavicular" if "acromio" in diagnostico else "Clavícula"
    
    if "escapula" in diagnostico:
        local += " / " if local else ""
        local += "Escapula"
    
    if "umero" in diagnostico:
        local += " / " if local else ""
        local += "Diáfise de " if "diafis" in diagnostico else "Cabeça do " if "cabeca" in diagnostico else ""
        local += "Úmero"
        local += " Distal" if "distal" in diagnostico else " Proximal" if "proximal" in diagnostico else ""

    if "radio" in diagnostico:
        local += " / " if local else ""
        local += "Diáfise de " if "diafis" in diagnostico else "Cabeça do " if "cabeca" in diagnostico else ""
        local += "Rádio"
        local += " Distal" if "distal" in diagnostico else " Proximal" if "proximal" in diagnostico else ""
        
    if "femur" in diagnostico:
        local += " / " if local else ""
        local += "Diáfise de " if "diafis" in diagnostico else "Cabeça do " if "cabeca" in diagnostico else ""
        local += "Fêmur"
        local += " Distal" if "distal" in diagnostico else " Proximal" if "proximal" in diagnostico else ""
    
    if "coronoide" in diagnostico:
        local += " / " if local else ""
        local += "Processo Coronoide"

    if "olecrano" in diagnostico:
        local += " / " if local else ""
        local += "Olécrano"

    if "cotovelo" in diagnostico:
        local += " / " if local else ""
        local += "Cotovelo"

    if "antebraco" in diagnostico:
        local += " / " if local else ""
        local += "Ossos do Antebraço"
    
    if ("carpo" in diagnostico or "escafoide" in diagnostico):
        local += " / " if local else ""
        local += "Ossos do Carpo ou Escafoide"

    if "falange" in diagnostico:
        local += " / " if local else ""
        local += "Falange"
    
    if "coluna" in diagnostico:
        local += " / " if local else ""
        local += "Coluna"
    
    if "sacro" in diagnostico:
        local += " / " if local else ""
        local += "Sacro"
    
    if "iliaco" in diagnostico or "isquio" in diagnostico or "pubis" in diagnostico or "pelve" in diagnostico:
        local += " / " if local else ""
        local += "Ossos da Pelve"

    if "patela" in diagnostico:
        local += " / " if local else ""
        local += "Patela"

    if "tibia" in diagnostico:
        if "plato" in diagnostico:
            local += " / " if local else ""
            local += "Platô Tibial"
        
        if "pilao" in diagnostico:

            local += " / " if local else ""
            local += "Pilão Tibial"
        
        if "plato" not in diagnostico and "pilao" not in diagnostico:
            local += " / " if local else ""
            local += "Ossos da Perna"    

    if "tornozelo" in diagnostico:
        local += " / " if local else ""
        local += "Tornozelo"

    if "calcaneo" in diagnostico:
        local += " / " if local else ""
        local += "Calcâneo"

    if "talus" in diagnostico:
        local += " / " if local else ""
        local += "Talus"

    if "navicular" in diagnostico:
        local += " / " if local else ""
        local += "Navicular"

    if "metatarso" in diagnostico:
        local += " / " if local else ""
        local += "Metatarso"
    
    return local if local else "Outros" if diagnostico else "Não Identificado"

df = pd.read_excel("CAIO.xlsx")
df.rename(columns = {
    "Unnamed: 10": "Lado Fratura",
    "diagnóstico": "Diagnóstico",
    "data internação": "Data Internação",
    "data da cirurgia": "Data Cirurgia",
    "data alta": "Data Alta"
}, inplace = True)

df2 = pd.read_excel("Caio.xls")
df2.rename(columns = {
    "sexo": "Sexo",
    "Nip": "NIP",
    "data da internação": "Data Internação",
    "data da alta": "Data Alta"
}, inplace = True)

df = df.append(df2, ignore_index = True)
df = df.dropna(axis = 1, how = "all").dropna(axis = 0, how = "all")
df.drop(df.tail(1).index, inplace = True)
df = df.fillna("")

df["Sexo"] = df["Sexo"].apply(lambda item: handle_sexo(item))
df["Lado Fratura"] = df["Lado Fratura"].apply(lambda item: handle_lado(item))
df["Diagnóstico"] = df["Diagnóstico"].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
df["Diagnóstico"] = df["Diagnóstico"].apply(lambda item: handle_diagnostico(item))
df["Tipo Fratura"] = df["Diagnóstico"].apply(lambda item: handle_tipo(item))
df["Local Fratura"] = df["Diagnóstico"].apply(lambda item: handle_local(item))
df["Local Fratura"] = df["Local Fratura"].str.split(" / ")
df = df.explode(["Local Fratura"])

df = df.reset_index()
df.to_excel("tratado.xlsx", index = False)