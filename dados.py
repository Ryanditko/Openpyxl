import openpyxl

# Carregar a planilha ou criar uma nova
wb = openpyxl.Workbook()
ws = wb.active

# Definir os títulos das colunas
headers = [
    'Inicio do intervalo', 'Fim do intervalo', 'Tipo de média', 'ID de agente', 'Nome do agente',
    'Atendidas', 'Tratamento', 'Tratamento médio', 'Conversação média', 'Espera média', 'TPC média',
    'Em espera', 'Transferidas'
]

# Adicionar os títulos na primeira linha
ws.append(headers)

# Dados a serem inseridos, organizados como uma lista de dicionários
data = [
    {

# Processo feito de forma manualmente, porem pode ser feito de forma automatizada com o uso de um banco de dados /  integração de uma API, de acordo com o sistema que você utiliza. Fora isso, tambem estou utilizando uma plataforma externa chamada de google magical extension para poder padronizar esse tipo de processo de forma manual, apenas escrevendo os valores que serão recebidos dentro dos headers. 

# Exemplo: 
'Inicio do intervalo':' 10:00',
'Fim do intervalo': '10:30',
'Tipo de média':'Média 2',
'ID de agente': 123,
'Nome do agente': 'Ryan Rodrigues',
'Atendidas': 100,
'Tratamento': 45,
'Tratamento médio': 2.5,
'Conversação média': 3.0,
'Espera média': 1.0,
'TPC média': 1.2,
'Em espera': 5,
'Transferidas': 3
    },

]

# Inserir os dados nas respectivas colunas
for row_data in data:
    row = [row_data[header] for header in headers]
    ws.append(row)

# Salvar a planilha
wb.save('Base_de_dados.xlsx')

# Resumo: Esse codigo promove a inserção de dados dentro de uma planilha excel, com o uso da biblioteca openpyxl.A planilha é criada e os dados são inseridos em linhas, onde cada linha é um dicionário que contém os dados a serem inseridos. A planilha é salva com o nome 'Base_de_dados.xlsx', que foi o nome que eu dei para a mesma, porem ele é alteravel ao seu gosto. 
# O código é bem simples e fácil de entender, e pode ser facilmente adaptado para outros tipos de dados e planilhas.