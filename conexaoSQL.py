import pyodbc

dados_conexao = (
    "DRIVER={SQL Server};"
    "SERVER=NEDF9906;"
    "Database=DW_SFCRI;"
)

conexao = pyodbc.connect(dados_conexao)
print("conexao bem sucedida")
