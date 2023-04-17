Este código em Python faz parte de um projeto de análise de impacto e tem como objetivo identificar os clientes afetados por uma falha em um determinado circuito de telecomunicação e gerar uma lista de e-mails de contato para esses clientes.

Como usar
Adicione os arquivos "IMPACTANALYSISS.xlsx" na pasta dist.
Abra o arquivo "main.exe".

Funcionamento
O código utiliza a biblioteca pandas para ler o arquivo "IMPACTANALYSISS.xlsx" e extrair as informações relevantes: o ID do cliente, o circuito afetado e a morada do cliente. Em seguida, é feita uma comparação com a planilha "Contactos CIP Remedy.xlsx" para encontrar os e-mails de contato correspondentes aos clientes afetados.
A comparação é feita utilizando a biblioteca fuzzywuzzy, que utiliza algoritmos de matching de strings para identificar correspondências parciais entre os nomes dos clientes e os nomes registrados na planilha de contatos. Se uma correspondência for encontrada com um score de pelo menos 85, o código assume que o contato encontrado é o correto e adiciona o e-mail correspondente à lista de e-mails de contato para esse cliente. Caso contrário, o código adiciona uma mensagem de "Sem correspondência no CIP." à lista de e-mails de contato.
O resultado final é uma planilha em Excel chamada "Resultado CIP.xlsx", contendo as informações do circuito afetado, do cliente afetado, da morada do cliente e dos e-mails de contato correspondentes. A planilha final é salva na pasta do projeto.