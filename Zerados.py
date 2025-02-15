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

valores_no_banco = resultadoSQL[coluna_chave].unique()
dados_nao_encontrados = resultadoExcel[~resultadoExcel[coluna_chave].isin(valores_no_banco)]


Funcoes.atualizar(dados_nao_encontrados, tabela_sql, cursor, conexao)
cursor.close()
conexao.close()






