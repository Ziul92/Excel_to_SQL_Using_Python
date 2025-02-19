import pandas as pd
import pyodbc


## Dados da conexão
dados_Conexao = (
    "Driver={SQL Server};"
    "Server=localhost;"
    "Database=Viagens_Governo;"
    "Trusted_Connection=yes;"
)

conexao = pyodbc.connect(dados_Conexao)
cursor = conexao.cursor()



pastaExcel = input('Qual a tabela a ser atualizada?')
coluna_chaveP = input('Qual a coluna chave?')

def ler_Excel():
    planilhaExcel = rf"C:\Users\luizh\OneDrive\Aulas Impacta\Prática\Viagens\2025\{pastaExcel}.xlsx"
    leituraDoExcel = pd.read_excel(planilhaExcel)
    return leituraDoExcel

def verificaSQL(conexao, tabela_sql):
    consulta_sql = f"SELECT {coluna_chave} FROM {tabela_sql}"
    leituraDoSql = pd.read_sql(consulta_sql, conexao)
    return leituraDoSql

def verificaColunas(conexao, tabela_sql):
    showColunas = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tabela_sql}'"
    lerColunas = pd.read_sql(showColunas, conexao)
    return lerColunas

def formatar_valores(valor):
    if isinstance(valor, list):
        return f"'{', '. join(map(str, valor))}'"
    elif pd.notnull(valor):
        return f"'{str(valor).replace('\'', '\'\'')}'"
    else:
        return "NULL"

def inserirDados():
    for index, row in dados_nao_encontrados.iterrows():
        colunas = ', '.join(f"[{col.strip()}]" for col in dados_nao_encontrados.columns)
        valores = ', '.join([formatar_valores(x) for x in row])
        query_insert = f"INSERT INTO [{tabela_sql}] ({colunas}) VALUES ({valores})"
        
        print("Query gerada:", query_insert)

        cursor.execute(query_insert)

    conexao.commit()
    print("Dados inseridos com sucesso no SQL Server")

def atualizar():
    if dados_nao_encontrados.empty: 
        print("Nenhum dado do Excel está faltando no banco de dados.") 
    else: 
        dados_nao_encontrados.to_excel("dados_nao_encontrados.xlsx", index=False) 
        print("Dados do Excel não estão no banco de dados: ") 
        print(dados_nao_encontrados)
        verificaSeAtualiza = input("Devemos atualizar as informações? (y/n)")
        if verificaSeAtualiza == "y":
            inserirDados()
        else:
            print("Dados não foram inseridos no Banco de Dados")


tabela_sql = pastaExcel
coluna_chave = coluna_chaveP

resultadoSQL = verificaSQL(conexao, tabela_sql)
resultadoExcel = ler_Excel()
resultadoColunas = verificaColunas(conexao, tabela_sql)


mapeamento_colunas = {coluna_excel: coluna_sql for coluna_excel, coluna_sql in zip(resultadoExcel.columns, resultadoColunas['COLUMN_NAME'])}
resultadoExcel.rename(columns=mapeamento_colunas, inplace=True)

valores_no_banco = resultadoSQL[coluna_chave].unique()
dados_nao_encontrados = resultadoExcel[~resultadoExcel[coluna_chave].isin(valores_no_banco)]



atualizar()
cursor.close()
conexao.close()