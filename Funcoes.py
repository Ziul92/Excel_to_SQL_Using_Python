import pandas as pd
import pyodbc

def verificaSQL(conexao, tabela_sql, coluna_chave):
    consulta_sql = f"SELECT {coluna_chave} FROm {tabela_sql}"
    leituraDoSql = pd.read_sql(consulta_sql, conexao)
    return leituraDoSql

def ler_Arquivo(caminho):
    planilha = caminho
    leituraPlanilha = pd.read_excel(planilha)
    return leituraPlanilha

def ler_CSV(caminho, limite_linhas=500000):
    planilhaCSV = caminho
    leituraDoCSV = pd.read_csv(planilhaCSV, delimiter=',', nrows=limite_linhas)
    return leituraDoCSV

def formatar_valores(valor):
    if isinstance(valor, list):
        return f"'{', '. join(map(str, valor))}'"
    elif pd.notnull(valor):
        return f"'{str(valor).replace('\'', '\'\'')}'"
    else:
        return "NULL"
    
def inserirDados(dados_nao_encontrados, tabela_sql, cursor, conexao):
    linhas_ignoradas = []
    erros = {}
    for index, row in dados_nao_encontrados.iterrows():
        colunas = ', '.join(f"[{col.strip()}]" for col in dados_nao_encontrados.columns)
        valores = ', '.join([formatar_valores(x) for x in row])
        query_insert = f"INSERT INTO [{tabela_sql}] ({colunas}) VALUES ({valores})"
        
        print("Query gerada:", query_insert)
        try:
            cursor.execute(query_insert)
        except pyodbc.Error as e:
            linhas_ignoradas.append((row))
            erros[index] = str(e)
            print(f"Erro ao inserir a linha {index}: {e}")

    conexao.commit()
    print("Dados inseridos com sucesso no SQL Server")

    if linhas_ignoradas:
        dados_linhas_ignoradas = pd.DataFrame([row for index in linhas_ignoradas])
        dados_linhas_ignoradas.to_excel("linhas_ignoradas.xlsx", index=False)
        print("Linhas ingnoradas pelo Truncamento foram salvas no arquivo 'Linhas_ignoradas.xlsx")
    
    if erros:
        print("Erros encontrados durante a inserção:")
        for linha, mensagem in erros.items():
            print(f"Linha {linha}: {mensagem}")

def verificaColunas(conexao, tabela_sql, planilha):
    showColunas = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tabela_sql}'"
    lerColunas = pd.read_sql(showColunas, conexao)
    mapeamento_colunas = {coluna_excel: coluna_sql for coluna_excel, coluna_sql in zip(planilha.columns, lerColunas['COLUMN_NAME'])}
    planilha.rename(columns=mapeamento_colunas, inplace=True)
    return planilha

def atualizar(dados_nao_encontrados, tabela_sql, cursor, conexao):
    if dados_nao_encontrados.empty: 
        print("Nenhum dado do Excel está faltando no banco de dados.") 
    else: 
        dados_nao_encontrados.to_excel("dados_nao_encontrados.xlsx", index=False) 
        print("Dados do Excel não estão no banco de dados: ") 
        print(dados_nao_encontrados)
        verificaSeAtualiza = input("Devemos atualizar as informações? (y/n)")
        if verificaSeAtualiza == "y":
            inserirDados(dados_nao_encontrados, tabela_sql, cursor, conexao)
        else:
            print("Dados não foram inseridos no Banco de Dados")

def comparacaoSQL_Excel(resultadoSQL, planilha, coluna_chave, conexao, tabela_sql):
    renomeiaColunas = verificaColunas(conexao, tabela_sql, planilha)
    valores_no_banco = resultadoSQL[coluna_chave].unique()
    dados_nao_encontrados = renomeiaColunas[~renomeiaColunas[coluna_chave].isin(valores_no_banco)]
    return dados_nao_encontrados


