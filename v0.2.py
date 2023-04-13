import pandas as pd

df = pd.read_excel('impacto analises.xlsx')
clientes_afetados = df.iloc[0:, 10].unique()

contatos = pd.read_excel('Contactos CIP Remedy.xlsx')
linhas_correspondentes = contatos.loc[contatos['LAST_NAME'].isin(clientes_afetados)]
linhas_correspondentes = linhas_correspondentes.sort_values(by=['LAST_NAME'])

emails = linhas_correspondentes['INTERNET_E_MAIL'].tolist()
dados = {'Cliente afetado': clientes_afetados, 'E-mails de contato': emails}
resultado = pd.DataFrame(dados)

print(resultado)