import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz

df = pd.read_excel('IMPACTANALYSISS.xlsx')
clientes_afetados = pd.Series(df.iloc[:, 10].astype(str).values)
clientes_afetados = clientes_afetados.str.replace(r'^PT\d+:', '', regex=True)
circuito = df['NCO'].values
morada_cliente = df.iloc[:, 12].values

max_length = max(len(clientes_afetados), len(circuito), len(morada_cliente))
clientes_afetados = np.pad(clientes_afetados, (0, max_length - len(clientes_afetados)), mode='constant')
circuito = np.pad(circuito, (0, max_length - len(circuito)), mode='constant')
morada_cliente = np.pad(morada_cliente, (0, max_length - len(morada_cliente)), mode='constant')

threshold = 90 # mínimo match de 90%
contatos = pd.read_excel('Contactos CIP Remedy.xlsx')
emails = []
for cliente_afetado in clientes_afetados:
    match = None
    for last_name in contatos['LAST_NAME']:
        ratio = fuzz.token_set_ratio(cliente_afetado, last_name)
        if ratio >= threshold:
            match = contatos[contatos['LAST_NAME'] == last_name]['INTERNET_E_MAIL'].iloc[0]
            break
    if match is None:
        match = 'Sem contato na base de dados do CIP.'
    emails.append(match)

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
