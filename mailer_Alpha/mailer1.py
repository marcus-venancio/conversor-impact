import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Dados de entrada
data_manutencao = '01/05/2023'
nome_empresa = 'Oni'

# Ler dados dos contatos
df_contatos = pd.read_excel('Contactos CIP Remedy.xlsx')

# Selecionar grupos correspondentes
clientes_afetados = df.iloc[0:, 10].unique()
grupos_correspondentes = df_contatos.loc[df_contatos['LAST_NAME'].isin(clientes_afetados)].groupby('LAST_NAME')

# Iterar sobre grupos e enviar e-mails
for cliente_afetado in clientes_afetados:
    grupo = grupos_correspondentes.get_group(cliente_afetado)
    emails_destino = grupo['INTERNET_E_MAIL'].unique()
    if len(emails_destino) > 0:
        mensagem = MIMEMultipart()
        mensagem['From'] = 'seu-email@dominio.com'
        mensagem['To'] = ', '.join(emails_destino)
        mensagem['Subject'] = 'Aviso de manutenção programada para ' + nome_empresa + ' que afetará seus circuitos.'
        corpo_mensagem = 'Prezado(a) ' + cliente_afetado + ',\n\nGostaríamos de informar que iremos realizar uma manutenção programada em nossa rede no dia ' + data_manutencao + ' o que afetará os circuitos que você utiliza em nosso serviço. O objetivo dessa manutenção é melhorar a qualidade e a estabilidade da nossa rede, garantindo que possamos continuar a oferecer um serviço de alta qualidade.\n\nPedimos desculpas por qualquer inconveniente que isso possa causar e estamos à disposição para esclarecer quaisquer dúvidas que possam surgir.\n\nAtenciosamente,\n\nEquipe de suporte técnico'
        mensagem.attach(MIMEText(corpo_mensagem, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login('seu-email@dominio.com', 'sua-senha')
        texto = mensagem.as_string()
        server.sendmail('seu-email@dominio.com', emails_destino, texto)
        server.quit()
