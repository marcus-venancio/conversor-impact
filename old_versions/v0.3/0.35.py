import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz, process

df = pd.read_excel('IMPACTANALYSISS.xlsx')
clientes_afetados = pd.Series(df.iloc[:, 10].astype(str).values)
circuito = df['NCO'].values
morada_cliente = df.iloc[:, 12].values

max_length = max(len(clientes_afetados), len(circuito), len(morada_cliente))
clientes_afetados = np.pad(clientes_afetados, (0, max_length - len(clientes_afetados)), mode='constant')
circuito = np.pad(circuito, (0, max_length - len(circuito)), mode='constant')
morada_cliente = np.pad(morada_cliente, (0, max_length - len(morada_cliente)), mode='constant')

contatos = pd.read_excel('Contactos CIP Remedy.xlsx')
emails = []
for cliente_afetado in clientes_afetados:
    try:
        cliente_afetado_cip = cliente_afetado.upper()
        matches = process.extractOne(cliente_afetado_cip, contatos['LAST_NAME'].dropna(), scorer=fuzz.partial_ratio)
        if matches[1] >= 85:
            grupo = contatos[contatos['LAST_NAME'] == matches[0]]
            email_grupo = grupo['INTERNET_E_MAIL'].values
            email = email_grupo[0] if len(email_grupo) > 0 else ''
        else:
            email = 'Sem correspondência no CIP.'
    except KeyError:
        email = 'Sem contato no CIP.'
    emails.append(email)

dados = {
    'Circuito NCO': circuito, 
    'Cliente afetado': clientes_afetados, 
    'Morada do cliente': morada_cliente, 
    'E-mails de contato': emails}

resultado = pd.DataFrame(dados)
resultado = resultado.dropna(subset=['Circuito NCO'])
resultado = resultado.drop_duplicates()

resultado.to_excel('Resultado CIP.xlsx', index=False)
print(resultado)