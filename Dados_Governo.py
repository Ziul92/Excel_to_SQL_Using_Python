import pandas as pd
import pyodbc
import Funcoes


## Dados da conexão
dados_Conexao = (
    "Driver={SQL Server};"
    "Server=localhost;"
    "Database=Viagens_Governo;"
    "Trusted_Connection=yes;"
)

conexao = pyodbc.connect(dados_Conexao)
cursor = conexao.cursor()

pastaExcel, coluna_chaveP, ano = Funcoes.verificaPlanilha()

caminho = rf"C:\Users\luizh\OneDrive\Aulas Impacta\Prática\Viagens\{ano}_{pastaExcel}.xlsx"
##caminhoCSV = rf"C:\Users\luizh\OneDrive\Aulas Impacta\Prática\Viagens\{ano}_{pastaExcel}.txt"


tabela_sql = f'{pastaExcel}_{ano}'
coluna_chave = coluna_chaveP

resultadoSQL = Funcoes.verificaSQL(conexao, tabela_sql, coluna_chave)
resultadoExcel = Funcoes.ler_Arquivo(caminho)
##resultadoCSV = Funcoes.ler_CSV(caminhoCSV)

comparacaoFinal = Funcoes.comparacaoSQL_Excel(resultadoSQL, resultadoExcel, coluna_chave, conexao, tabela_sql)

Funcoes.atualizar(comparacaoFinal, tabela_sql, cursor, conexao)
cursor.close()
conexao.close()