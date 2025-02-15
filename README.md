# Excel_to_SQL_Using_Python

# Arquivo Funções.py

--> VerificaSQL <--
Faz a consulta com o servidor SQL
Parâmetros: 
- conexao: Dados da conexão usando pyodbc com os dados
Ex de dados: dados_Conexao = (
    "Driver={SQL Server};"
    "Server=localhost;"
    "Database=Database;"
    "Trusted_Connection=yes;"
)
- tabela_sql: Informação de qual a tabela que será conectada no Banco de Dados
Ex: tabela_sql = "tabela"
- coluna_chave: Informação de qual a coluna principal que será utilizada para verificar os valores unicos entre SQL e Excel/CSV


--> ler_Arquivo <--
Faz a consulta com Excel
Parâmetros:
- caminho: informação de qual caminho que está o arquivo Excel/CSV


--> ler_CSV <--
Faz a consulta com CSV
Parâmetros:
- caminho: Mesmo exemplo do Excel acima
- limite_linhas: Informação de qual o limite de linhas que será verificado no arquivo CSV. 
OBS: Não é obrigatorio, mas possui por padrão 500000 (quinhentas mil linhas) para pesquisa, caso tenha mais, informar no parâmetro.

--> atualizar <--
Utilizado para verificar se a variável dados_nao_encontrados se encontra vazia, caso sim, apenas mostra que não a dados a serem atualizados, caso esteja com informações, mostra os valores a serem atualizados, questiona se quer atualizar e se y será feita a inicialização da função inserirDados
Parâmetros:
- dados_nao_encontrados: Variavel para juntar os dados que não constam no Banco depois da verificação, seguindo informações/return da função ComparacaoSQL_Excel.
- tabela_sql: Mesmo citado acima
- cursor: váriavel criado para inicializar a conexão com o servidor.
- conexao: Mesmo citado mais acima.
Obs: A função inserirDados é chamada apenas após a confimação de que os valores podem ser inseridos.


--> inserirDados <--
Após a verificação do SQL com o Excel/CSV, são feitas as querys para insersão dos dados de acordo com a variavel dados_nao_encontrados
Parâmetros:
- dados_nao_encontrados: Variavel cidada acima
- tabela_sql: Mesmo citado acima
- cursor: váriavel criado para inicializar a conexão com o servidor.
- conexao: Mesmo citado mais acima.
Obs: Possui um tratamento de erro na Query e ao final caso alguma das informações tenha Truncado ou algo do tipo, será informado em qual coluna foi o erro sem evitar que milhares de linhas precisem ser iniciadas novamente.
Obs 2: A variavel linhas_ignoradas ainda está sendo tratada, está apenas mostrando as informações da última linha que consta erro, repetindo ela pelo numero de erros.

 --> comparacaoSQL_Excel
 Faz a verificação entre o banco de dados SQL e o arquivo Excel/CSV e adiciona a variavel dados_nao_encontrados, fazendo return da mesma
