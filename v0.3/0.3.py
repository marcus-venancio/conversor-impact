import pandas as pd
import numpy as np

df = pd.read_excel('Marcus_testev2.xlsx')
clientes_afetados = df.iloc[0:, 1].values
circuito = df.iloc[0:, 0].unique()
morada_cliente = df.iloc[0:, 2].unique()

# Make all arrays the same length
max_length = max(len(clientes_afetados), len(circuito), len(morada_cliente))
clientes_afetados = np.pad(
    clientes_afetados, (0, max_length - len(clientes_afetados)), mode='constant')
circuito = np.pad(circuito, (0, max_length - len(circuito)), mode='constant')
morada_cliente = np.pad(
    morada_cliente, (0, max_length - len(morada_cliente)), mode='constant')

contatos = pd.read_excel('Contactos CIP Remedy.xlsx')
grupos_correspondentes = contatos.loc[contatos['LAST_NAME'].isin(
    clientes_afetados)].groupby('LAST_NAME')

emails = []
for cliente_afetado in clientes_afetados:
    try:
        grupo = grupos_correspondentes.get_group(cliente_afetado)
        email_grupo = grupo['INTERNET_E_MAIL'].unique()
        email = email_grupo[0] if len(email_grupo) > 0 else ''
    except KeyError:
        email = ''
    emails.append(email)

dados = {'Clientes afetados': clientes_afetados, 'Circuito': circuito,
         'Morada do Cliente': morada_cliente, 'E-mails de contato': emails}
resultado = pd.DataFrame(dados)

resultado.to_excel('Resultado2.xlsx', index=False)
print(resultado)
