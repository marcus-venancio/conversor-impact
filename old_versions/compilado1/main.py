import pandas as pd
import numpy as np

df = pd.read_excel('Impact Analysis.xlsx')
clientes_afetados = df.iloc[:, 10].values
circuito = df.iloc[:, 9].values
morada_cliente = df.iloc[:, 12].values

max_length = max(len(clientes_afetados), len(circuito), len(morada_cliente))
clientes_afetados = np.pad(clientes_afetados, (0, max_length - len(clientes_afetados)), mode='constant')
circuito = np.pad(circuito, (0, max_length - len(circuito)), mode='constant')
morada_cliente = np.pad(morada_cliente, (0, max_length - len(morada_cliente)), mode='constant')

contatos = pd.read_excel('Contactos CIP Remedy.xlsx')
grupos_correspondentes = contatos.loc[contatos['LAST_NAME'].isin(clientes_afetados)].groupby('LAST_NAME')
emails = [grupo['INTERNET_E_MAIL'].iloc[0] if len(grupo) > 0 else '' for _, grupo in grupos_correspondentes]

emails = []
for cliente_afetado in clientes_afetados:
    try:
        grupo = grupos_correspondentes.get_group(cliente_afetado)
        email_grupo = grupo['INTERNET_E_MAIL'].values
        email = email_grupo[0] if len(email_grupo) > 0 else ''
    except KeyError:
        email = 'Sem contacto na base de dados do CIP.'
    emails.append(email)

dados = {
    'Circuito NCO': circuito, 
    'Cliente afetado': clientes_afetados, 
    'Morada do cliente': morada_cliente, 
    'E-mails de contato': emails}
resultado = pd.DataFrame(dados)

resultado.to_excel('Resultado.xlsx', index=False)