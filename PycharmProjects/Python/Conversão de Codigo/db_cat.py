import pandas as pd
import mysql.connector

# Função para criar a conexão com o banco de dados
def criar_conexao():
    return mysql.connector.connect(
        host="127.0.0.1",
        user="mateus_ramos",
        password="flamengo4422",
        database="db_cat"
    )

# Função para importar dados do Excel para o banco de dados
def importar_dados_do_excel(arquivo_excel, tabela):
    # Ler a planilha Excel
    df = pd.read_excel(arquivo_excel)

    # Conectar-se ao banco de dados
    cnx = criar_conexao()
    cursor = cnx.cursor()

    try:
        # Inserir os dados no banco de dados
        for _, row in df.iterrows():
            valores = tuple(row)
            placeholders = ', '.join(['%s'] * len(valores))
            sql = f"INSERT INTO {tabela} VALUES ({placeholders})"
            cursor.execute(sql, valores)

        # Confirmar as alterações e fechar a conexão
        cnx.commit()
        print("Dados importados com sucesso!")
    except mysql.connector.Error as error:
        print("Erro ao importar dados:", error)
    finally:
        cursor.close()
        cnx.close()

# Exemplo de uso
arquivo_excel = 'produtosconv.xlsx'
tabela_destino = 'nome_da_tabela'

importar_dados_do_excel(arquivo_excel, tabela_destino)
