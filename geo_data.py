import pandas as pd
def get_sigla(provincia):
    df = pd.read_csv('gi_province.csv', sep= ';')
    return df.loc[df['denominazione_provincia'] == provincia]['sigla_provincia'].item()

def get_comuni(sigla_provincia):
    df = pd.read_csv('gi_comuni.csv', sep=';')
    return [i for i in df.loc[df['sigla_provincia'] == sigla_provincia]['denominazione_ita']]
        

if __name__ == '__main__':
    print(get_comuni(get_sigla('Cagliari')))