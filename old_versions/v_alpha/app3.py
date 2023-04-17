import pandas as pd
import openpyxl

# Lendo o arquivo 'impacto_analises.xlsx'
df_impacto = pd.read_excel('impacto analises.xlsx', sheet_name='IMPACTANALYSISS_cadao_168123043', header=1)

# Filtrando empresas da coluna 11 a partir da linha 2
empresas_filtradas = df_impacto.iloc[0:, 10].unique()

# Lendo o arquivo 'Contactos CIP Remedy.xlsx'
workbook = openpyxl.load_workbook('Contactos CIP Remedy.xlsx')
sheet = workbook.active

# Filtrando contato na linha 5
contato_filtrado = sheet.cell(row=2, column=5).value

# Criando uma nova tabela com as duas informações filtradas
df_nova_tabela = pd.DataFrame({'Empresa': empresas_filtradas, 'Contato': contato_filtrado})

# Exibindo a nova tabela
print(df_nova_tabela)