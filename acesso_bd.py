import pyodbc

# Define a string de conexão com o banco de dados do Access
# Substitua 'caminho_do_arquivo.accdb' pelo caminho do seu arquivo .mdb ou .accdb
conexao_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\GESTÃO\PYTHON\AUTOMACAO\RELATORIOS_SGP\BD\bd_relatorios.accdb;'

# Conecta ao banco de dados
conexao = pyodbc.connect(conexao_str)

# Cria um cursor para executar comandos SQL
cursor = conexao.cursor()

# Define a consulta SQL para selecionar todos os dados da primeira tabela
consulta_sql_enderecos = 'SELECT * FROM Enderecos ORDER BY Indice ASC;'
cursor.execute(consulta_sql_enderecos)

# Obtém os resultados da consulta para a primeira tabela
enderecos = cursor.fetchall()

# Define a consulta SQL para selecionar todos os dados da segunda tabela
consulta_sql_datas = 'SELECT * FROM Datas;'
cursor.execute(consulta_sql_datas)

# Obtém os resultados da consulta para a segunda tabela
datas = cursor.fetchall()

# Fecha a conexão e o cursor
cursor.close()
conexao.close()

# Atualizar dados na tabela
def atualizar_dados():
    cursor.execute("UPDATE Pessoa SET Idade = 35 WHERE Nome = 'João'")
    connection.commit()
    ("Dados atualizados com sucesso.")


periodo = datas[0][1]
dia_inicial = datas[0][1].replace('-', '.').split()[0].split('.')[0]
mes_inicial = datas[0][1].replace('-', '.').split()[0].split('.')[1]
ano_inicial = '20' + datas[0][1].replace('-', '.').split()[0].split('.')[2]
dia_final = datas[0][1].replace('-', '.').split()[2].split('.')[0]
mes_final = datas[0][1].replace('-', '.').split()[2].split('.')[1]
ano_final = '20' + datas[0][1].replace('-', '.').split()[2].split('.')[2]
producao_analitica = enderecos[0][2]
protocolo = enderecos[1][2] + periodo
sgp_geral = enderecos[2][2]
acoes_abertas = enderecos[3][2]
acoes_fechadas = enderecos[4][2] + periodo



# print(f'E periodo definido foi: {periodo}')
# print(f'O dia inicial da apuração é: {dia_inicial}')
# print(f'O mês inicial da apuração é: {mes_inicial}')
# print(f'O ano inicial da apuração é: {ano_inicial}\n')
# print(f'O dia final da apuração é: {dia_final}')
# print(f'O mês final da apuração é: {mes_final}')
# print(f'O ano final da apuração é: {ano_final}')
# print(f'Produção analitica: {producao_analitica}')
# print(f'Protocolo: {protocolo}')
# print(f'Ações Abertas: {acoes_abertas}')
# print(f'Ações Fechadas: {acoes_fechadas}')
# print(f'SGP Geral: {sgp_geral}')
# print(f'Periodo de apuração: {periodo}')
# print(mes_atual)
