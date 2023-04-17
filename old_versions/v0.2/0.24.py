import pandas as pd

df = pd.read_excel('Marcus_testev2.xlsx')
clientes_afetados = df.iloc[0:, 1].unique()
circuito = df.iloc[0:, 0].unique()
morada_cliente = df.iloc[0:, 2].unique()

contatos = pd.read_excel('Contactos CIP Remedy.xlsx')
grupos_correspondentes = contatos.loc[contatos['LAST_NAME'].isin(clientes_afetados)].groupby('LAST_NAME')

emails = []
for cliente_afetado in clientes_afetados:
    grupo = grupos_correspondentes.get_group(cliente_afetado)
    email_grupo = grupo['INTERNET_E_MAIL'].unique()
    email = email_grupo[0] if len(email_grupo) > 0 else ''
    emails.append(email)

dados = {'Cliente afetado': clientes_afetados, 'Circuito': circuito, 'Morada': morada_cliente, 'E-mails de contato': emails}
resultado = pd.DataFrame(dados)

resultado.to_excel('Resultado.xlsx', index=False)