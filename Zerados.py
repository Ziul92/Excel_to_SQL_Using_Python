import pandas as pd
import pyodbc
import Funcoes

## Dados da conexão 
dados_Conexao = (
    "Driver={SQL Server};"
    "Server=localhost;"
    "Database=Zerados;"
    "Trusted_Connection=yes;"
)
# Caminho onde está o arquivo
caminho = r"C:\Users\luizh\OneDrive\Zerados\Zerados.xlsx"

conexao = pyodbc.connect(dados_Conexao)
cursor = conexao.cursor()

tabela_sql = "Zerados"
coluna_chave = "Numero"

resultadoSQL = Funcoes.verificaSQL(conexao, tabela_sql, coluna_chave)
resultadoExcel = Funcoes.ler_Arquivo(caminho)

comparacaoFinal = Funcoes.comparacaoSQL_Excel(resultadoSQL, resultadoExcel, coluna_chave)

Funcoes.atualizar(comparacaoFinal, tabela_sql, cursor, conexao)
cursor.close()
conexao.close()






