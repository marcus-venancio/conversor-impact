import pandas as pd
import openpyxl

book = openpyxl.load_workbook('impacto analises.xlsx')
empresas_page = book['IMPACTANALYSISS_cadao_168123043']
for cols in empresas_page.iter_cols(min_col=11, max_col=11):
    for cell in cols:
        print(cell.value)