import pandas as pd

df = pd.read_excel('impacto analises.xlsx')
clientes_afetados = df.iloc[0:, 10].unique()

contatos = pd.read_excel('Contactos CIP Remedy.xlsx')
linhas_correspondentes = contatos.loc[contatos['LAST_NAME'].isin(clientes_afetados)]

print(linhas_correspondentes)